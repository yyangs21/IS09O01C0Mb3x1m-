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

# ---------------------------
# CONFIG
# ---------------------------
load_dotenv()  # carga .env en local si existe
st.set_page_config(page_title="Formulario ISO 9001 — Inteligente", layout="wide", page_icon="📄")

# CSS / Diseño visual mejorado
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }
    .header { border-radius:12px; padding:14px; background: linear-gradient(90deg,#f7fbff, #ffffff); box-shadow: 0 6px 20px rgba(13,38,66,0.06); }
    .card { background:#fff; padding:12px; border-radius:10px; box-shadow:0 6px 18px rgba(12,40,80,0.04); margin-bottom:10px; }
    .chip { display:inline-block; padding:6px 10px; margin:4px; border-radius:18px; background:#f1f7ff; border:1px solid #e1efff; font-size:14px; }
    .small{ font-size:13px; color:#666; }
    </style>
    """,
    unsafe_allow_html=True,
)

# Header image (intentará cargar local/asset; si está en repo público, Streamlit lo servirá)
def load_image_try(path):
    try:
        return Image.open(path)
    except Exception:
        return None

# Mostrar header (intento local o desde assets en repo)
header_img = load_image_try("assets/Encabezado.png") or load_image_try("Encabezado.png")
if header_img:
    st.image(header_img, use_column_width=True)
else:
    st.markdown("<div class='header'><h2>📄 Formulario ISO 9001 — Inteligente</h2><p class='small'>Actualiza Google Sheets → la app se actualiza automáticamente</p></div>", unsafe_allow_html=True)

st.write("")

# ---------------------------
# CARGAR CREDENCIALES: modo cloud (st.secrets) o local (service_account.json)
# - En Streamlit Cloud: poner SERVICE_ACCOUNT_JSON en Secrets (valor = JSON completo como string)
# - También configurar OPENAI_API_KEY en Secrets
# ---------------------------

# 1) OpenAI key: prioridad Secrets (streamlit cloud) -> .env local
OPENAI_KEY = None
if "OPENAI_API_KEY" in st.secrets:
    OPENAI_KEY = st.secrets["OPENAI_API_KEY"]
else:
    OPENAI_KEY = os.getenv("OPENAI_API_KEY")

if OPENAI_KEY:
    openai.api_key = OPENAI_KEY
else:
    st.warning("OPENAI API key no detectada. Algunas funciones de IA no estarán disponibles. (Setea OPENAI_API_KEY en Streamlit Secrets o .env).")

# 2) Google Sheets auth: prioridad st.secrets["SERVICE_ACCOUNT_JSON"] (string) -> local service_account.json
def get_gspread_client():
    # scope para Sheets/Drive
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    # preferir secrets de Streamlit (SERVICE_ACCOUNT_JSON)
    if "SERVICE_ACCOUNT_JSON" in st.secrets:
        try:
            sa_info = st.secrets["SERVICE_ACCOUNT_JSON"]
            if isinstance(sa_info, str):
                sa_json = json.loads(sa_info)
            else:
                sa_json = sa_info
            creds = ServiceAccountCredentials.from_json_keyfile_dict(sa_json, scope)
            return gspread.authorize(creds)
        except Exception as e:
            st.error(f"Error autenticando con SERVICE_ACCOUNT_JSON en secrets: {e}")
            raise e
    # fallback: archivo local service_account.json
    elif os.path.exists("service_account.json"):
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
            return gspread.authorize(creds)
        except Exception as e:
            st.error(f"Error autenticando con service_account.json local: {e}")
            raise e
    else:
        st.error("No se encontró credencial de Google Sheets. Añade SERVICE_ACCOUNT_JSON en Streamlit Secrets o sube service_account.json local.")
        st.stop()

# ---------------------------
# Conectar y leer hojas
# ---------------------------
SHEET_URL = "https://docs.google.com/spreadsheets/d/1mQY0_MEjluVT95iat5_5qGyffBJGp2n0hwEChvp2Ivs"

try:
    gc = get_gspread_client()
    sh = gc.open_by_url(SHEET_URL)
    df_areas = pd.DataFrame(sh.worksheet("Areas").get_all_records())
    df_claus = pd.DataFrame(sh.worksheet("Clausulas").get_all_records())
    df_ent = pd.DataFrame(sh.worksheet("Entregables").get_all_records())
except Exception as e:
    st.error(f"Error leyendo Google Sheets: {e}")
    st.stop()

# Validaciones mínimas
required_areas_cols = ["Area", "Dueño del Proceso", "Puesto", "Correo"]
if not set(required_areas_cols).issubset(df_areas.columns):
    st.error(f"La hoja 'Areas' debe contener columnas: {required_areas_cols}. Revisa nombres y mayúsculas.")
    st.stop()

# ---------------------------
# UI: selector área + info
# ---------------------------
left, right = st.columns([2,1])
with left:
    area = st.selectbox("Selecciona tu área", options=df_areas["Area"].unique())
with right:
    st.markdown("**Acciones**")
    if st.button("Refrescar datos"):
        st.experimental_rerun()

st.write("")

# Info del dueño/proceso
info = df_areas[df_areas["Area"] == area].iloc[0]
st.markdown(f"<div class='card'><strong>{area}</strong><br><span class='small'>Dueño: {info['Dueño del Proceso']} &nbsp; | &nbsp; Puesto: {info['Puesto']} &nbsp; | &nbsp; {info.get('Correo','')}</span></div>", unsafe_allow_html=True)

# Cláusulas aplicables (chips)
st.subheader("Cláusulas ISO aplicables")
cl_area = df_claus[df_claus["Area"] == area]
if cl_area.empty:
    st.info("No hay cláusulas registradas para esta área.")
else:
    for _, r in cl_area.iterrows():
        st.markdown(f"<span class='chip'>{r.get('Clausula','')} — {r.get('Descripción','')}</span>", unsafe_allow_html=True)

st.write("")
# Entregables (cards)
st.subheader("Entregables asignados")
ent_area = df_ent[df_ent["Area"] == area]
if ent_area.empty:
    st.info("No hay entregables asignados para esta área.")
else:
    for _, r in ent_area.iterrows():
        st.markdown(f"<div class='card'><strong>{r.get('Categoría','')}</strong><br>{r.get('Entregable','')}<br><span class='small'>Estado: {r.get('Estado','')}</span></div>", unsafe_allow_html=True)

st.write("")
# Inputs para nuevo entregable / comentario
st.markdown("### Registrar / Analizar un entregable")
col_a, col_b = st.columns([2,1])
with col_a:
    nueva_categoria = st.text_input("Categoría", value="")
    nuevo_entregable = st.text_input("Entregable / Tarea", value="")
    nota_descr = st.text_area("Descripción / Comentarios", value="", height=120)
with col_b:
    prioridad = st.selectbox("Prioridad", ["Baja","Media","Alta"])
    fecha_compromiso = st.date_input("Fecha compromiso")
    responsable = st.text_input("Responsable (si aplica)", value=info.get("Dueño del Proceso",""))

# ---------------------------
# IA: Generar resumen / checklist / acciones
# ---------------------------
def make_prompt(area, info, clausulas_records, entregables_records, descripcion, prioridad):
    prompt = f"""
Eres un experto en Sistemas de Gestión de Calidad ISO 9001. 
Genera un resumen ejecutivo y acciones concretas.

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
4) Checklist de 3 ítems para la próxima reunión
"""
    return prompt

resumen_ia = None
if st.button("🤖 Generar resumen IA"):
    if not OPENAI_KEY:
        st.error("No se detectó clave de OpenAI. Añádela a Streamlit Secrets como OPENAI_API_KEY o a .env local.")
    else:
        clausulas_records = cl_area.to_dict("records")
        entregables_records = {"entregable": nuevo_entregable, "descripcion": nota_descr}
        prompt = make_prompt(area, info, clausulas_records, entregables_records, nota_descr, prioridad)
        try:
            # Llamada a OpenAI (ChatCompletion)
            resp = openai.ChatCompletion.create(
                model="gpt-5.1-mini",
                messages=[{"role":"user","content": prompt}],
                temperature=0.2,
                max_tokens=700
            )
            resumen_ia = resp["choices"][0]["message"]["content"].strip()
            st.success("Resumen IA generado")
            st.markdown(f"<div class='card'>{resumen_ia}</div>", unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Error en llamada a OpenAI: {e}")

# ---------------------------
# Guardar nuevo entregable en Sheets
# ---------------------------
if st.button("💾 Guardar entregable en Sheets"):
    if not nuevo_entregable:
        st.warning("Agrega texto en 'Entregable / Tarea' para guardar.")
    else:
        try:
            row = [area, nueva_categoria, nuevo_entregable, str(fecha_compromiso), prioridad, responsable, "Pendiente"]
            sh.worksheet("Entregables").append_row(row)
            st.success("Entregable agregado en Google Sheets ✔️")
        except Exception as e:
            st.error(f"Error guardando en Sheets: {e}")

# ---------------------------
# Generar PDF con header/footer (si hay resumen_ia o nota)
# ---------------------------
def build_pdf_bytes(area, info, nuevo_entregable, nota_descr, resumen_ia):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, topMargin=40, bottomMargin=80)
    styles = getSampleStyleSheet()
    story = []
    # Header image - intentamos cargar local/assets
    header_path = "assets/Encabezado.png"
    if not os.path.exists(header_path) and os.path.exists("Encabezado.png"):
        header_path = "Encabezado.png"
    if os.path.exists(header_path):
        try:
            story.append(RLImage(header_path, width=500, height=60))
            story.append(Spacer(1, 8))
        except Exception:
            pass

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

    # Footer image
    footer_path = "assets/Pie.png"
    if not os.path.exists(footer_path) and os.path.exists("Pie.png"):
        footer_path = "Pie.png"

    # add footer image later using canvas if needed; for now append image at end
    if os.path.exists(footer_path):
        try:
            story.append(Spacer(1, 20))
            story.append(RLImage(footer_path, width=500, height=50))
        except Exception:
            pass

    doc.build(story)
    buf.seek(0)
    return buf

if st.button("📥 Generar y descargar PDF"):
    pdf_buf = build_pdf_bytes(area, info, nuevo_entregable, nota_descr, resumen_ia or "")
    st.download_button("Descargar Reporte PDF", data=pdf_buf, file_name=f"Reporte_ISO_{area}.pdf", mime="application/pdf")

# Footer visual
footer_img = load_image_try("assets/Pie.png") or load_image_try("Pie.png")
if footer_img:
    st.image(footer_img, use_column_width=True)
else:
    st.markdown("<div class='small' style='text-align:center;margin-top:20px;color:#777;'>Formulario automatizado · Mantenimiento ISO · Generado con IA</div>", unsafe_allow_html=True)
