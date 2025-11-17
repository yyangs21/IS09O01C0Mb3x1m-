# FormularioISO.py
import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import io
import os
from dotenv import load_dotenv
import openai
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter
from PIL import Image
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ---------------------------
# CONFIG
# ---------------------------
load_dotenv()
st.set_page_config(page_title="Formulario ISO 9001 — Inteligente", layout="wide", page_icon="📄")

# CSS / Diseño visual
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }
    .header { border-radius:12px; padding:14px; background: linear-gradient(90deg,#f7fbff, #ffffff); box-shadow: 0 6px 20px rgba(13,38,66,0.06); text-align:center;}
    .card { background:#fff; padding:12px; border-radius:10px; box-shadow:0 6px 18px rgba(12,40,80,0.04); margin-bottom:10px; }
    .chip { display:inline-block; padding:6px 10px; margin:4px; border-radius:18px; background:#f1f7ff; border:1px solid #e1efff; font-size:14px; }
    .small{ font-size:13px; color:#666; }
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
# CARGAR CREDENCIALES
# ---------------------------
OPENAI_KEY = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
if OPENAI_KEY:
    openai.api_key = OPENAI_KEY

def get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = None
    if "SERVICE_ACCOUNT_JSON" in st.secrets:
        sa_info = st.secrets["SERVICE_ACCOUNT_JSON"]
        sa_json = json.loads(sa_info) if isinstance(sa_info, str) else sa_info
        creds = ServiceAccountCredentials.from_json_keyfile_dict(sa_json, scope)
    elif os.path.exists("service_account.json"):
        creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
    else:
        st.error("No se encontró credencial de Google Sheets.")
        st.stop()
    return gspread.authorize(creds)

def get_drive_service():
    scope = ["https://www.googleapis.com/auth/drive"]
    creds = None
    if "SERVICE_ACCOUNT_JSON" in st.secrets:
        sa_info = st.secrets["SERVICE_ACCOUNT_JSON"]
        sa_json = json.loads(sa_info) if isinstance(sa_info, str) else sa_info
        creds = ServiceAccountCredentials.from_json_keyfile_dict(sa_json, scope)
    elif os.path.exists("service_account.json"):
        creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
    else:
        st.error("No se encontró credencial de Google Drive.")
        st.stop()
    return build('drive', 'v3', credentials=creds)

# ---------------------------
# LEER SHEETS
# ---------------------------
SHEET_URL = "https://docs.google.com/spreadsheets/d/1mQY0_MEjluVT95iat5_5qGyffBJGp2n0hwEChvp2Ivs"
gc = get_gspread_client()
sh = gc.open_by_url(SHEET_URL)

def load_sheets():
    df_areas = pd.DataFrame(sh.worksheet("Areas").get_all_records())
    df_claus = pd.DataFrame(sh.worksheet("Clausulas").get_all_records())
    df_ent = pd.DataFrame(sh.worksheet("Entregables").get_all_records())
    return df_areas, df_claus, df_ent

df_areas, df_claus, df_ent = load_sheets()

# ---------------------------
# VALIDAR HOJA AREAS
# ---------------------------
required_areas_cols = ["Area", "Dueño del Proceso", "Puesto", "Correo"]
actual_cols_norm = [c.strip().lower() for c in df_areas.columns]
required_cols_norm = [c.strip().lower() for c in required_areas_cols]

if not set(required_cols_norm).issubset(actual_cols_norm):
    st.error(f"La hoja 'Areas' debe contener columnas: {required_areas_cols}. Revisa nombres exactos.")
    st.stop()

# Renombrar columnas
col_mapping = {actual_col: req_col for req_col in required_areas_cols for actual_col in df_areas.columns if req_col.strip().lower() == actual_col.strip().lower()}
df_areas.rename(columns=col_mapping, inplace=True)

# ---------------------------
# HEADER
# ---------------------------
header_img = load_image_try("assets/Encabezado.png") or load_image_try("Encabezado.png")
if header_img:
    try:
        st.image(header_img, width=800, output_format="PNG")
    except:
        st.markdown("<div class='header'><h2>📄 Formulario ISO 9001 — Inteligente</h2></div>", unsafe_allow_html=True)
else:
    st.markdown("<div class='header'><h2>📄 Formulario ISO 9001 — Inteligente</h2></div>", unsafe_allow_html=True)

st.write("")

# ---------------------------
# SELECTOR DE AREA
# ---------------------------
left, right = st.columns([2,1])
with left:
    area = st.selectbox("Selecciona tu área", options=df_areas["Area"].unique())
with right:
    if st.button("🔄 Refrescar datos"):
        df_areas, df_claus, df_ent = load_sheets()
        st.experimental_rerun()

info = df_areas[df_areas["Area"].str.strip().str.lower() == area.strip().lower()].iloc[0]
st.markdown(f"<div class='card'><strong>{area}</strong><br><span class='small'>Dueño: {info['Dueño del Proceso']} | Puesto: {info['Puesto']} | {info.get('Correo','')}</span></div>", unsafe_allow_html=True)

# ---------------------------
# CLÁUSULAS
# ---------------------------
st.subheader("Cláusulas ISO aplicables")
cl_area = df_claus[df_claus["Area"].str.strip().str.lower() == area.strip().lower()]
if cl_area.empty:
    st.info("No hay cláusulas registradas para esta área.")
else:
    for _, r in cl_area.iterrows():
        st.markdown(f"<span class='chip'>{r.get('Clausula','')} — {r.get('Descripción','')}</span>", unsafe_allow_html=True)

# ---------------------------
# ENTREGABLES ASIGNADOS
# ---------------------------
st.subheader("Entregables asignados")
ent_area = df_ent[df_ent["Area"].str.strip().str.lower() == area.strip().lower()]
if ent_area.empty:
    st.info("No hay entregables asignados para esta área.")
else:
    for _, r in ent_area.iterrows():
        st.markdown(f"<div class='card'><strong>{r.get('Categoría','')}</strong><br>{r.get('Entregable','')}<br><span class='small'>Estado: {r.get('Estado','')}</span></div>", unsafe_allow_html=True)

# ---------------------------
# NUEVO ENTREGABLE
# ---------------------------
st.markdown("### Registrar / Analizar un entregable")
col_a, col_b = st.columns([2,1])
with col_a:
    nueva_categoria = st.text_input("Categoría", value="")
    nuevo_entregable = st.text_input("Entregable / Tarea", value="")
    nota_descr = st.text_area("Descripción / Comentarios", value="", height=120)
    archivo = st.file_uploader("Subir archivo entregable (PDF/Word/Excel)", type=["pdf","docx","xlsx"])
with col_b:
    prioridad = st.selectbox("Prioridad", ["Baja","Media","Alta"])
    fecha_compromiso = st.date_input("Fecha compromiso")
    responsable = st.text_input("Responsable", value=info.get("Dueño del Proceso",""))

# ---------------------------
# ---------------------------
# # ---------------------------
# FUNCIÓN IA
# ---------------------------
resumen_ia = None

def make_prompt(area, info, clausulas_records, entregables_records, descripcion, prioridad):
    prompt = f"""
Eres un experto en Sistemas de Gestión de Calidad ISO 9001.
Área: {area}
Dueño del proceso: {info.get('Dueño del Proceso')}
Puesto: {info.get('Puesto')}
Cláusulas aplicables: {', '.join([str(x.get('Clausula','')) for x in clausulas_records])}
Entregable a analizar: {entregables_records}
Descripción: {descripcion}
Prioridad: {prioridad}

Entrega:
1) Resumen ejecutivo (2-3 líneas)
2) 3 riesgos principales con impacto
3) 4 acciones recomendadas, priorizadas
4) Checklist de 3 ítems
"""
    return prompt

if st.button("🤖 Consultar IA"):
    if not OPENAI_KEY:
        st.error("No se detectó clave de OpenAI. Verifica tus secrets o variable de entorno.")
    else:
        clausulas_records = cl_area.to_dict("records")
        entregables_records = {"entregable": nuevo_entregable, "descripcion": nota_descr}
        prompt = make_prompt(area, info, clausulas_records, entregables_records, nota_descr, prioridad)
        try:
            resp = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[{"role":"user","content": prompt}],
                temperature=0.2,
                max_tokens=700
            )
            resumen_ia = resp.choices[0].message['content'].strip()
            st.markdown(f"<div class='card'>{resumen_ia}</div>", unsafe_allow_html=True)
        except openai.error.OpenAIError as e:
            st.error(f"Error al consultar OpenAI: {e}")
        except Exception as e:
            st.error(f"Ocurrió un error inesperado al consultar la IA: {e}")



# ---------------------------
# SUBIR ENTREGABLE A SHEET + DRIVE
# ---------------------------
DRIVE_FOLDER_ID = "1ueBPvyVPoSkz0VoLXIkulnwLw3YX"
drive_service = get_drive_service()

if st.button("💾 Guardar entregable"):
    if not nuevo_entregable:
        st.warning("Agrega texto en 'Entregable / Tarea'.")
    else:
        file_url = ""
        if archivo:
            try:
                file_metadata = {"name": archivo.name, "parents": [DRIVE_FOLDER_ID]}
                media = MediaIoBaseUpload(archivo, mimetype=archivo.type, resumable=True)
                file_drive = drive_service.files().create(body=file_metadata, media_body=media, fields="id, webViewLink").execute()
                file_url = file_drive.get("webViewLink","")
            except Exception as e:
                st.error(f"Error subiendo archivo a Drive: {e}")

        row = [area, nueva_categoria, nuevo_entregable, str(fecha_compromiso), prioridad, responsable, "Pendiente", nota_descr, file_url]
        try:
            sh.worksheet("Entregables").append_row(row)
            st.success("Entregable registrado correctamente.")
        except Exception as e:
            st.error(f"Error guardando en Sheets: {e}")

# ---------------------------
# GENERAR PDF
# ---------------------------
def build_pdf_bytes(area, info, nuevo_entregable, nota_descr, resumen_ia):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, topMargin=40, bottomMargin=80)
    styles = getSampleStyleSheet()
    story = []

    header_path = "assets/Encabezado.png"
    if os.path.exists(header_path):
        story.append(RLImage(header_path, width=500, height=60))
        story.append(Spacer(1, 8))

    story.append(Paragraph(f"<b>Reporte ISO 9001 — {area}</b>", styles["Title"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"Dueño: {info.get('Dueño del Proceso')} — Puesto: {info.get('Puesto')} — Email: {info.get('Correo')}", styles["Normal"]))
    story.append(Spacer(1, 6))
    story.append(Paragraph("<b>Entregable</b>", styles["Heading3"]))
    story.append(Paragraph(f"{nuevo_entregable}", styles["Normal"]))
    story.append(Spacer(1, 6))
    story.append(Paragraph("<b>Descripción</b>", styles["Normal"]))
    story.append(Paragraph(nota_descr or "-", styles["Normal"]))
    story.append(Spacer(1, 8))
    if resumen_ia:
        story.append(Paragraph("<b>Resumen IA</b>", styles["Heading3"]))
        story.append(Paragraph(resumen_ia, styles["Normal"]))
        story.append(Spacer(1, 8))

    footer_path = "assets/Pie.png"
    if os.path.exists(footer_path):
        story.append(Spacer(1, 20))
        story.append(RLImage(footer_path, width=500, height=50))

    doc.build(story)
    buf.seek(0)
    return buf

if st.button("📥 Generar y descargar PDF"):
    pdf_buf = build_pdf_bytes(area, info, nuevo_entregable, nota_descr, resumen_ia or "")
    st.download_button("Descargar Reporte PDF", data=pdf_buf, file_name=f"Reporte_ISO_{area}.pdf", mime="application/pdf")

# ---------------------------
# FOOTER
# ---------------------------
footer_img = load_image_try("assets/Pie.png") or load_image_try("Pie.png")
if footer_img:
    try:
        st.image(footer_img, width=800, output_format="PNG")
    except:
        st.markdown("<div class='small' style='text-align:center;margin-top:20px;color:#777;'>Formulario automatizado · Mantenimiento ISO · Generado con IA</div>", unsafe_allow_html=True)
else:
    st.markdown("<div class='small' style='text-align:center;margin-top:20px;color:#777;'>Formulario automatizado · Mantenimiento ISO · Generado con IA</div>", unsafe_allow_html=True)



