# FormularioISO.py
import streamlit as st
import pandas as pd
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import json
import io
import os
from dotenv import load_dotenv
import openai
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from PIL import Image
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from datetime import datetime

# ---------------------------
# CONFIG
# ---------------------------
load_dotenv()
st.set_page_config(page_title="Formulario ISO 9001 — Inteligente", layout="wide", page_icon="📄")

# CSS / Diseño visual (todo texto azul)
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    html, body, [class*="css"]  { font-family: 'Inter', sans-serif; color: #0033cc !important; }
    .header { border-radius:12px; padding:8px 14px; background: linear-gradient(90deg,#f7fbff, #ffffff); box-shadow: 0 6px 20px rgba(13,38,66,0.06); color:#0033cc !important;}
    .card { background:#fff; padding:12px; border-radius:10px; box-shadow:0 6px 18px rgba(12,40,80,0.04); margin-bottom:10px; color:#0033cc !important;}
    .chip { display:inline-block; padding:6px 10px; margin:4px; border-radius:18px; background:#f1f7ff; border:1px solid #e1efff; font-size:14px; color:#0033cc !important;}
    .small{ font-size:13px; color:#0033cc !important; }
    label, input, select, textarea, .stTextInput, .stDateInput { color: #0033cc !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

def load_image_try(path):
    try:
        return Image.open(path)
    except Exception:
        return None

# ---------------------------
# OPENAI API (leer key desde secrets/.env)
# ---------------------------
OPENAI_KEY = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
if OPENAI_KEY:
    openai.api_key = OPENAI_KEY

def query_openai(prompt, model="gpt-3.5-turbo", temperature=0.2, max_tokens=700):
    if not OPENAI_KEY:
        raise RuntimeError("OPENAI API key no configurada.")
    try:
        if hasattr(openai, "chat") and hasattr(openai.chat, "completions"):
            resp = openai.chat.completions.create(
                model=model,
                messages=[{"role":"user","content": prompt}],
                temperature=temperature,
                max_tokens=max_tokens
            )
            try:
                return resp.choices[0].message.content
            except Exception:
                return getattr(resp.choices[0].message, "content", str(resp))
        else:
            resp = openai.ChatCompletion.create(
                model=model,
                messages=[{"role":"user","content": prompt}],
                temperature=temperature,
                max_tokens=max_tokens
            )
            try:
                return resp.choices[0].message['content']
            except Exception:
                try:
                    return resp.choices[0].message.content
                except Exception:
                    return str(resp)
    except Exception as e:
        raise

# ---------------------------
# ---------------------------
# GSPREAD CLIENT (OAuth desde Streamlit Secrets)
# ---------------------------
def get_gspread_client_oauth_secrets():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    
    # Leer credenciales desde Streamlit secrets
    creds_json_str = st.secrets.get("SERVICE_ACCOUNT_JSON")
    if not creds_json_str:
        st.error("No se encontró SERVICE_ACCOUNT_JSON en Streamlit Secrets.")
        st.stop()
    
    creds_dict = json.loads(creds_json_str)
    
    # Autorizar
    gc = gspread.service_account_from_dict(creds_dict)
    return gc

# ---------------------------
# GOOGLE DRIVE SERVICE (OAuth desde secrets)
# ---------------------------
def get_drive_service_oauth_secrets():
    creds_json_str = st.secrets.get("SERVICE_ACCOUNT_JSON")
    if not creds_json_str:
        st.error("No se encontró SERVICE_ACCOUNT_JSON en Streamlit Secrets.")
        st.stop()
    
    creds_dict = json.loads(creds_json_str)
    gc = gspread.service_account_from_dict(creds_dict)
    creds = gc.auth  # credenciales de google-auth
    
    drive_service = build("drive", "v3", credentials=creds)
    return drive_service

# ---------------------------
# Uso
# ---------------------------
gc = get_gspread_client_oauth_secrets()
try:
    sh = gc.open_by_url(SHEET_URL)
except Exception as e:
    st.error(f"Error accediendo a Google Sheets: {e}")
    st.stop()

drive_service = get_drive_service_oauth_secrets()

# ---------------------------
# LEER SHEETS
# ---------------------------
SHEET_URL = "https://docs.google.com/spreadsheets/d/1mQY0_MEjluVT95iat5_5qGyffBJGp2n0hwEChvp2Ivs"
gc = get_gspread_client_oauth()
try:
    sh = gc.open_by_url(SHEET_URL)
except Exception as e:
    st.error(f"Error accediendo a Google Sheets: {e}")
    st.stop()

def load_sheets():
    df_areas = pd.DataFrame(sh.worksheet("Areas").get_all_records())
    df_claus = pd.DataFrame(sh.worksheet("Clausulas").get_all_records())
    df_ent = pd.DataFrame(sh.worksheet("Entregables").get_all_records())
    try:
        df_carga = pd.DataFrame(sh.worksheet("Carga").get_all_records())
    except Exception:
        df_carga = pd.DataFrame(columns=["Area","Categoria","Entregable","Fecha Entrega","Link Documento"])
    return df_areas, df_claus, df_ent, df_carga

df_areas, df_claus, df_ent, df_carga = load_sheets()

# ---------------------------
# VALIDAR HOJA AREAS
# ---------------------------
required_areas_cols = ["Area", "Dueño del Proceso", "Puesto", "Correo"]
actual_cols_norm = [c.strip().lower() for c in df_areas.columns]
required_cols_norm = [c.strip().lower() for c in required_areas_cols]
if not set(required_cols_norm).issubset(actual_cols_norm):
    st.error(f"La hoja 'Areas' debe contener columnas: {required_areas_cols}. Revisa nombres exactos.")
    st.stop()

col_mapping = {}
for req_col in required_areas_cols:
    for actual_col in df_areas.columns:
        if req_col.strip().lower() == actual_col.strip().lower():
            col_mapping[actual_col] = req_col
df_areas.rename(columns=col_mapping, inplace=True)

# ---------------------------
# HEADER
# ---------------------------
header_img = load_image_try("assets/Encabezado.png") or load_image_try("Encabezado.png")
hcol1, hcol2 = st.columns([3,1])
with hcol1:
    if header_img:
        st.image(header_img, width=450)
    else:
        st.markdown("<div class='header'><h2>📄 Formulario ISO 9001 — Inteligente</h2></div>", unsafe_allow_html=True)
with hcol2:
    st.markdown("<div style='text-align:left'><h3 style='color:#FFFFFF;margin-top:18px'>MANTENIMIENTO ISO 9001:2015</h3></div>", unsafe_allow_html=True)

st.write("")

# ---------------------------
# SELECTOR DE AREA
# ---------------------------
left, right = st.columns([2,1])
with left:
    area = st.selectbox("Selecciona tu área", options=sorted(df_areas["Area"].unique()))
with right:
    if st.button("🔄 Refrescar datos"):
        df_areas, df_claus, df_ent, df_carga = load_sheets()
        st.experimental_rerun()

info = df_areas[df_areas["Area"].str.strip().str.lower() == area.strip().lower()].iloc[0]
st.markdown(f"<div class='card'><strong>{area}</strong><br><span class='small'>Dueño: {info['Dueño del Proceso']} | Puesto: {info['Puesto']} | {info.get('Correo','')}</span></div>", unsafe_allow_html=True)

# ---------------------------
# CLÁUSULAS
# ---------------------------
st.subheader("Cláusulas ISO aplicables")
cl_area = df_claus[df_claus["Area"].str.strip().str.lower() == area.strip().lower()] if not df_claus.empty else pd.DataFrame()
if cl_area.empty:
    st.info("No hay cláusulas registradas para esta área.")
else:
    for _, r in cl_area.iterrows():
        st.markdown(f"<span class='chip'>{r.get('Clausula','')} — {r.get('Descripcion', r.get('Descripción',''))}</span>", unsafe_allow_html=True)

# ---------------------------
# ENTREGABLES
# ---------------------------
st.subheader("Entregables asignados")
ent_area = df_ent[df_ent["Area"].str.strip().str.lower() == area.strip().lower()] if not df_ent.empty else pd.DataFrame()
if ent_area.empty:
    st.info("No hay entregables asignados para esta área.")
else:
    for _, r in ent_area.iterrows():
        st.markdown(f"<div class='card'><strong>{r.get('Categoria','')}</strong><br>{r.get('Entregable','')}<br><span class='small'>Estado: {r.get('Estado','')}</span></div>", unsafe_allow_html=True)

# ---------------------------
# NUEVO ENTREGABLE
# ---------------------------
st.markdown("### Registrar / Analizar un entregable")
col_a, col_b = st.columns([2,1])

cats_area = []
if not ent_area.empty:
    cats_area = sorted([str(x).strip() for x in ent_area["Categoria"].dropna().unique()])
    if "" in cats_area:
        cats_area = [c for c in cats_area if c]
if not cats_area:
    cats_area = ["(Sin categorías en Entregables)"]

with col_a:
    nueva_categoria = st.selectbox("Categoría", options=cats_area)
    if nueva_categoria and nueva_categoria != "(Sin categorías en Entregables)":
        mask = (df_ent["Area"].str.strip().str.lower() == area.strip().lower()) & (df_ent["Categoria"].str.strip().str.lower() == str(nueva_categoria).strip().lower())
        posibles = df_ent[mask]["Entregable"].dropna().unique().tolist()
        posibles = [str(x) for x in posibles]
        if not posibles:
            posibles = ["(Sin entregables en esta categoría)"]
    else:
        posibles = ["(Sin entregables disponibles)"]

    nuevo_entregable = st.selectbox("Entregable / Tarea", options=posibles)
    nota_descr = st.text_area("Descripción / Comentarios (libre)", value="", height=140)
    archivo = st.file_uploader("Subir archivo entregable (PDF/Word/Excel)", type=["pdf","docx","xlsx"], key="uploader")

with col_b:
    prioridad_options = []
    if nueva_categoria and nueva_categoria != "(Sin categorías en Entregables)":
        pri = df_ent[ (df_ent["Area"].str.strip().str.lower() == area.strip().lower()) & (df_ent["Categoria"].str.strip().str.lower() == str(nueva_categoria).strip().lower()) ]["Prioridad"].dropna().unique().tolist()
        prioridad_options = [str(x) for x in pri if str(x).strip()]
    if not prioridad_options:
        prioridad_options = ["Baja","Media","Alta"]

    default_date = None
    if nuevo_entregable and nuevo_entregable not in ("(Sin entregables en esta categoría)","(Sin entregables disponibles)"):
        row_match = df_ent[
            (df_ent["Area"].str.strip().str.lower() == area.strip().lower()) &
            (df_ent["Categoria"].str.strip().str.lower() == str(nueva_categoria).strip().lower()) &
            (df_ent["Entregable"].str.strip().str.lower() == str(nuevo_entregable).strip().lower())
        ]
        if not row_match.empty:
            date_col = None
            for col_try in ["Fecha Compromiso","Fecha compromiso","Fecha Entrega","FechaEntrega","Fecha"]:
                if col_try in row_match.columns:
                    date_col = col_try
                    break
            if date_col:
                val = row_match.iloc[0].get(date_col)
                if pd.notna(val) and val != "":
                    try:
                        dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
                        if pd.notna(dt):
                            default_date = dt.date()
                    except Exception:
                        default_date = None
            try:
                p = row_match.iloc[0].get("Prioridad") or row_match.iloc[0].get("prioridad") or None
                if p and str(p).strip() and str(p).strip() in prioridad_options:
                    prioridad_default = str(p).strip()
                else:
                    prioridad_default = prioridad_options[0]
            except Exception:
                prioridad_default = prioridad_options[0]
        else:
            prioridad_default = prioridad_options[0]
    else:
        prioridad_default = prioridad_options[0]

    prioridad = st.selectbox("Prioridad", options=prioridad_options, index=max(0, prioridad_options.index(prioridad_default) if prioridad_default in prioridad_options else 0))
    fecha_compromiso = st.date_input("Fecha compromiso", value=default_date if default_date else datetime.today().date())
    responsable = st.text_input("Responsable", value=info.get("Dueño del Proceso",""))
    estado = st.text_input("Estado", value="Pendiente")

# ---------------------------
# CHAT / CONSULTA IA
# ---------------------------
st.subheader("💬 Consultar IA sobre cláusulas o entregables")
pregunta_ia = st.text_input("Escribe tu duda o consulta sobre ISO 9001 para tu área:", key="pregunta_ia")

if st.button("Preguntar a la IA"):
    if not OPENAI_KEY:
        st.error("No se detectó clave de OpenAI. Añade OPENAI_API_KEY en Streamlit Secrets o .env.")
    elif not pregunta_ia.strip():
        st.warning("Escribe tu consulta antes de enviar.")
    else:
        try:
            contexto = f"""
Eres un experto en Sistemas de Gestión de Calidad ISO 9001.
Área: {area}
Dueño del proceso: {info.get('Dueño del Proceso')}
Puesto: {info.get('Puesto')}
Cláusulas aplicables: {', '.join([str(x.get('Clausula','')) for x in cl_area.to_dict('records')]) if not cl_area.empty else 'N/A'}
Entregables asignados: {', '.join([str(x.get('Entregable','')) for x in ent_area.to_dict('records')]) if not ent_area.empty else 'N/A'}

Consulta del usuario: {pregunta_ia}

Responde de manera clara, práctica y breve, indicando si aplica y qué acciones recomendarías.
"""
            respuesta = query_openai(contexto, model="gpt-3.5-turbo", temperature=0.2, max_tokens=500)
            st.markdown(f"<div class='card'>{respuesta}</div>", unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Ocurrió un error inesperado al consultar la IA: {e}")

# ---------------------------
# GUARDAR ENTREGABLE
# ---------------------------
DRIVE_FOLDER_ID = "1ueBPvyVPoSkz0VoLXIkulnwLw3am3WYX"
drive_service = None

if "credentials.json" in os.listdir("."):
    try:
        drive_service = get_drive_service_oauth()
    except Exception:
        drive_service = None

def subir_archivo_drive(service_drive, archivo, carpeta_id):
    archivo_bytes = archivo.read()
    fh = io.BytesIO(archivo_bytes)
    media = MediaIoBaseUpload(fh, mimetype=archivo.type, resumable=False)
    file_metadata = {"name": archivo.name, "parents": [carpeta_id]}
    created = service_drive.files().create(body=file_metadata, media_body=media, fields="id").execute()
    return created.get("id")

def hacer_publico(service_drive, file_id):
    try:
        service_drive.permissions().create(
            fileId=file_id,
            body={"role": "reader", "type": "anyone"},
            fields="id"
        ).execute()
    except Exception:
        pass

def asegurar_hoja_carga(sh_obj):
    try:
        ws = sh_obj.worksheet("Carga")
    except Exception:
        try:
            sh_obj.add_worksheet(title="Carga", rows="1000", cols="10")
            ws = sh_obj.worksheet("Carga")
            ws.append_row(["Area","Categoria","Entregable","Fecha Entrega","Link Documento"])
        except Exception:
            ws = None
    return ws

if st.button("💾 Guardar entregable"):
    if not nuevo_entregable or nuevo_entregable in ("(Sin entregables en esta categoría)","(Sin entregables disponibles)"):
        st.warning("Selecciona un entregable válido antes de guardar.")
    else:
        file_link = ""
        if archivo is not None:
            try:
                if drive_service is None:
                    drive_service = get_drive_service_oauth()
                file_id = subir_archivo_drive(drive_service, archivo, DRIVE_FOLDER_ID)
                try:
                    hacer_publico(drive_service, file_id)
                except Exception:
                    pass
                file_link = f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"
                st.success("Archivo subido a Google Drive correctamente.")
            except Exception as e:
                st.error(f"No se pudo subir el archivo a Drive: {e}")
                file_link = ""

        ws_carga = asegurar_hoja_carga(sh)
        if ws_carga is not None:
            try:
                ws_carga.append_row([area, nueva_categoria, nuevo_entregable, fecha_compromiso.strftime("%d/%m/%Y"), file_link])
                st.success("Entregable guardado en la hoja de Carga ✅")
            except Exception as e:
                st.error(f"No se pudo guardar en Google Sheets: {e}")
        else:
            st.error("No se pudo asegurar la hoja 'Carga' para registrar datos.")
# ---------------------------
# GENERAR PDF (descarga)
# ---------------------------
def build_pdf_bytes(area, info, nuevo_entregable, nota_descr, resumen_ia=""):
    """
    Genera un PDF en memoria con el reporte de ISO 9001 para el área seleccionada.
    Devuelve un objeto BytesIO listo para descargar.
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, topMargin=40, bottomMargin=80)
    styles = getSampleStyleSheet()
    blue_style = ParagraphStyle(
        'blue', parent=styles['Normal'], textColor=colors.blue
    )
    story = []

    # Encabezado
    header_path = "assets/Encabezado.png"
    if os.path.exists(header_path):
        try:
            story.append(RLImage(header_path, width=500, height=60))
            story.append(Spacer(1, 8))
        except Exception:
            pass

    # Título y datos del área
    story.append(Paragraph(f"<b>Reporte ISO 9001 — {area}</b>", blue_style))
    story.append(Spacer(1, 8))
    story.append(
        Paragraph(
            f"Dueño: {info.get('Dueño del Proceso')} — Puesto: {info.get('Puesto')} — Email: {info.get('Correo')}",
            blue_style,
        )
    )
    story.append(Spacer(1, 6))

    # Entregable
    story.append(Paragraph("<b>Entregable</b>", blue_style))
    story.append(Paragraph(f"{nuevo_entregable}", blue_style))
    story.append(Spacer(1, 6))

    # Descripción / comentarios
    story.append(Paragraph("<b>Descripción</b>", blue_style))
    story.append(Paragraph(nota_descr or "-", blue_style))
    story.append(Spacer(1, 8))

    # Resumen IA (opcional)
    if resumen_ia:
        story.append(Paragraph("<b>Resumen IA</b>", blue_style))
        story.append(Paragraph(resumen_ia, blue_style))
        story.append(Spacer(1, 8))

    # Pie de página (imagen)
    footer_path = "assets/Pie.png"
    if os.path.exists(footer_path):
        try:
            story.append(Spacer(1, 20))
            story.append(RLImage(footer_path, width=500, height=50))
        except Exception:
            pass

    # Construir PDF
    doc.build(story)
    buf.seek(0)
    return buf


# Botón para generar y descargar PDF
if st.button("📥 Generar y descargar PDF"):
    pdf_buf = build_pdf_bytes(area, info, nuevo_entregable, nota_descr)
    st.download_button(
        "Descargar Reporte PDF",
        data=pdf_buf,
        file_name=f"Reporte_ISO_{area}.pdf",
        mime="application/pdf",
    )


# ---------------------------
# FOOTER STREAMLIT
# ---------------------------
footer_img = load_image_try("assets/Pie.png") or load_image_try("Pie.png")
if footer_img:
    try:
        st.image(footer_img, width=800)
    except Exception:
        st.markdown(
            "<div class='small' style='text-align:center;margin-top:20px;color:#0033cc;'>"
            "Formulario automatizado · Mantenimiento ISO · Generado con IA</div>",
            unsafe_allow_html=True,
        )
else:
    st.markdown(
        "<div class='small' style='text-align:center;margin-top:20px;color:#0033cc;'>"
        "Formulario automatizado · Mantenimiento ISO · Generado con IA</div>",
        unsafe_allow_html=True,
    )


