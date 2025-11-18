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

# Helper para consultar OpenAI con compatibilidad (intenta nueva interfaz, luego fallback)
def query_openai(prompt, model="gpt-3.5-turbo", temperature=0.2, max_tokens=700):
    if not OPENAI_KEY:
        raise RuntimeError("OPENAI API key no configurada.")
    # Intentar la nueva interfaz (openai.chat.completions.create)
    try:
        if hasattr(openai, "chat") and hasattr(openai.chat, "completions"):
            resp = openai.chat.completions.create(
                model=model,
                messages=[{"role":"user","content": prompt}],
                temperature=temperature,
                max_tokens=max_tokens
            )
            # lectura flexible del contenido
            try:
                return resp.choices[0].message.content
            except Exception:
                return getattr(resp.choices[0].message, "content", str(resp))
        else:
            # fallback a la API antigua si está disponible
            resp = openai.ChatCompletion.create(
    if not nuevo_entregable or nuevo_entregable in ("(Sin entregables en esta categoría)","(Sin entregables disponibles)"):
        st.warning("Selecciona un entregable válido antes de guardar.")
    else:
        file_url = ""
        # Subir archivo a Drive si viene
        if archivo is not None:
            try:
                # archivo es UploadedFile; convertir a BytesIO
                archivo_bytes = archivo.read()
                fh = io.BytesIO(archivo_bytes)
                if drive_service is None:
                    # intentar iniciar
                    drive_service = get_drive_service()
                file_metadata = {"name": archivo.name, "parents": [DRIVE_FOLDER_ID]}
                media = MediaIoBaseUpload(fh, mimetype=archivo.type, resumable=True)
                file_drive = drive_service.files().create(body=file_metadata, media_body=media, fields="id, webViewLink").execute()
                file_url = file_drive.get("webViewLink","")
            except Exception as e:
                st.warning(f"Error subiendo archivo a Drive (el entregable se guardará en Sheets sin link): {e}")
                file_url = ""

        # FORMATO de fecha a string
        fecha_str = fecha_compromiso.strftime("%Y-%m-%d")

        # Append a hoja Entregables (fila base)
        row_ent = [area, nueva_categoria, nuevo_entregable, fecha_str, prioridad, responsable, estado, nota_descr, file_url]
        try:
            ws_ent = sh.worksheet("Entregables")
            ws_ent.append_row(row_ent)
            st.success("Entregable guardado en hoja 'Entregables'.")
        except Exception as e:
            st.error(f"Error guardando en hoja 'Entregables': {e}")

        # Append a hoja Carga (si existe) con: Area, Categoria, Entregable, Fecha Entrega, Link documento
        try:
            try:
                ws_carga = sh.worksheet("Carga")
            except Exception:
                # intentar crear la hoja si no existe (sheet add)
                try:
                    sh.add_worksheet(title="Carga", rows="1000", cols="10")
                    ws_carga = sh.worksheet("Carga")
                    # escribir encabezado
                    ws_carga.append_row(["Area","Categoria","Entregable","Fecha Entrega","Link documento"])
                except Exception:
                    ws_carga = None
            if ws_carga:
                row_carga = [area, nueva_categoria, nuevo_entregable, fecha_str, file_url]
                ws_carga.append_row(row_carga)
                st.success("Registro añadido en hoja 'Carga'.")
        except Exception as e:
            st.warning(f"No se pudo registrar en hoja 'Carga': {e}")

# ---------------------------
# GENERAR PDF (descarga)
# ---------------------------
def build_pdf_bytes(area, info, nuevo_entregable, nota_descr, resumen_ia=""):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, topMargin=40, bottomMargin=80)
    styles = getSampleStyleSheet()
    blue_style = ParagraphStyle('blue', parent=styles['Normal'], textColor=colors.blue)
    story = []

    header_path = "assets/Encabezado.png"
    if os.path.exists(header_path):
        try:
            story.append(RLImage(header_path, width=500, height=60))
            story.append(Spacer(1, 8))
        except Exception:
            pass

    story.append(Paragraph(f"<b>Reporte ISO 9001 — {area}</b>", blue_style))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"Dueño: {info.get('Dueño del Proceso')} — Puesto: {info.get('Puesto')} — Email: {info.get('Correo')}", blue_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph("<b>Entregable</b>", blue_style))
    story.append(Paragraph(f"{nuevo_entregable}", blue_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph("<b>Descripción</b>", blue_style))
    story.append(Paragraph(nota_descr or "-", blue_style))
    story.append(Spacer(1, 8))
    if resumen_ia:
        story.append(Paragraph("<b>Resumen IA</b>", blue_style))
        story.append(Paragraph(resumen_ia, blue_style))
        story.append(Spacer(1, 8))

    footer_path = "assets/Pie.png"
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
    pdf_buf = build_pdf_bytes(area, info, nuevo_entregable, nota_descr)
    st.download_button("Descargar Reporte PDF", data=pdf_buf, file_name=f"Reporte_ISO_{area}.pdf", mime="application/pdf")

# ---------------------------
# FOOTER
# ---------------------------
footer_img = load_image_try("assets/Pie.png") or load_image_try("Pie.png")
if footer_img:
    try:
        st.image(footer_img, width=800)
    except Exception:
        st.markdown("<div class='small' style='text-align:center;margin-top:20px;color:#0033cc;'>Formulario automatizado · Mantenimiento ISO · Generado con IA</div>", unsafe_allow_html=True)
else:
    st.markdown("<div class='small' style='text-align:center;margin-top:20px;color:#0033cc;'>Formulario automatizado · Mantenimiento ISO · Generado con IA</div>", unsafe_allow_html=True)
