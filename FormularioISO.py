import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import json
import io
import os
def load_image_try(path):
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
def query_openai(prompt, model="gpt-3.5-turbo", temperature=0.2, max_tokens=700)
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
                model=model,
                messages=[{"role":"user","content": prompt}],
                temperature=temperature,
                max_tokens=max_tokens
            )
            # intento de lectura flexible
            try:
                return resp.choices[0].message['content']
            except Exception:
def query_openai(prompt, model="gpt-3.5-turbo", temperature=0.2, max_tokens=700)
                except Exception:
                    return str(resp)
    except Exception as e:
        # rebota la excepción para manejo en UI
        raise

# ---------------------------
# GSPREAD CLIENT
# GSPREAD CLIENT (OAuth)
# ---------------------------
def get_gspread_client_oauth():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds = None
    # si ya tenemos token previo
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token_file:
            creds = pickle.load(token_file)
    # si no es válido, iniciar flujo OAuth
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes=scopes)
            creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as token_file:
            pickle.dump(creds, token_file)
    return gspread.authorize(creds)

# ---------------------------
# GOOGLE DRIVE SERVICE
# GOOGLE DRIVE SERVICE (OAuth)
# ---------------------------
def get_drive_service_oauth():
    scopes = ["https://www.googleapis.com/auth/drive"]
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token_file:
            creds = pickle.load(token_file)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes=scopes)
            creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as token_file:
            pickle.dump(creds, token_file)
    return build("drive", "v3", credentials=creds)

# ---------------------------
# LEER SHEETS
# ---------------------------
SHEET_URL = "https://docs.google.com/spreadsheets/d/1mQY0_MEjluVT95iat5_5qGyffBJGp2n0hwEChvp2Ivs"
gc = get_gspread_client()
gc = get_gspread_client_oauth()
try:
    sh = gc.open_by_url(SHEET_URL)
except Exception as e:
    st.error(f"Error accediendo a Google Sheets: {e}")
    st.stop()

def load_sheets():
    # lee hojas principales; si no existe 'Carga' se notificará más adelante al guardar
    df_areas = pd.DataFrame(sh.worksheet("Areas").get_all_records())
    df_claus = pd.DataFrame(sh.worksheet("Clausulas").get_all_records())
    df_ent = pd.DataFrame(sh.worksheet("Entregables").get_all_records())
    # si existe hoja 'Carga' la leemos (no es obligatorio)
    try:
        df_carga = pd.DataFrame(sh.worksheet("Carga").get_all_records())
    except Exception:
        # aseguramos encabezado correcto (Link Documento con D mayúscula)
        df_carga = pd.DataFrame(columns=["Area","Categoria","Entregable","Fecha Entrega","Link Documento"])
    return df_areas, df_claus, df_ent, df_carga

@@ -156,12 +153,10 @@ def load_sheets():
required_areas_cols = ["Area", "Dueño del Proceso", "Puesto", "Correo"]
actual_cols_norm = [c.strip().lower() for c in df_areas.columns]
required_cols_norm = [c.strip().lower() for c in required_areas_cols]

if not set(required_cols_norm).issubset(actual_cols_norm):
    st.error(f"La hoja 'Areas' debe contener columnas: {required_areas_cols}. Revisa nombres exactos.")
    st.stop()

# Renombrar columnas (normalizar nombres)
col_mapping = {}
for req_col in required_areas_cols:
    for actual_col in df_areas.columns:
@@ -170,10 +165,9 @@ def load_sheets():
df_areas.rename(columns=col_mapping, inplace=True)

# ---------------------------
# HEADER con titulo adicional a la derecha
# HEADER
# ---------------------------
header_img = load_image_try("assets/Encabezado.png") or load_image_try("Encabezado.png")

hcol1, hcol2 = st.columns([3,1])
with hcol1:
    if header_img:
@@ -196,7 +190,6 @@ def load_sheets():
        df_areas, df_claus, df_ent, df_carga = load_sheets()
        st.experimental_rerun()

# Info dueño proceso
info = df_areas[df_areas["Area"].str.strip().str.lower() == area.strip().lower()].iloc[0]
st.markdown(f"<div class='card'><strong>{area}</strong><br><span class='small'>Dueño: {info['Dueño del Proceso']} | Puesto: {info['Puesto']} | {info.get('Correo','')}</span></div>", unsafe_allow_html=True)

@@ -212,7 +205,7 @@ def load_sheets():
        st.markdown(f"<span class='chip'>{r.get('Clausula','')} — {r.get('Descripcion', r.get('Descripción',''))}</span>", unsafe_allow_html=True)

# ---------------------------
# ENTREGABLES ASIGNADOS (vista)
# ENTREGABLES
# ---------------------------
st.subheader("Entregables asignados")
ent_area = df_ent[df_ent["Area"].str.strip().str.lower() == area.strip().lower()] if not df_ent.empty else pd.DataFrame()
@@ -223,27 +216,21 @@ def load_sheets():
        st.markdown(f"<div class='card'><strong>{r.get('Categoria','')}</strong><br>{r.get('Entregable','')}<br><span class='small'>Estado: {r.get('Estado','')}</span></div>", unsafe_allow_html=True)

# ---------------------------
# NUEVO ENTREGABLE: CAMPOS POBLADOS DESDE SHEET (solo descripcion y estado libres)
# NUEVO ENTREGABLE
# ---------------------------
st.markdown("### Registrar / Analizar un entregable")
col_a, col_b = st.columns([2,1])

# --- Obtener listas según area
# Categorias disponibles para el area
cats_area = []
if not ent_area.empty:
    cats_area = sorted([str(x).strip() for x in ent_area["Categoria"].dropna().unique()])
    if "" in cats_area:
        cats_area = [c for c in cats_area if c]
# Si no hay categorias, dejar una opción vacía
if not cats_area:
    cats_area = ["(Sin categorías en Entregables)"]

with col_a:
    # Categoria: selectbox poblado desde sheet Entregables filtrado por area
    nueva_categoria = st.selectbox("Categoría", options=cats_area)
    # Entregables para la categoria y area
    # si la hoja contiene columnas con mayúsculas distintas; usar get y fillna
    if nueva_categoria and nueva_categoria != "(Sin categorías en Entregables)":
        mask = (df_ent["Area"].str.strip().str.lower() == area.strip().lower()) & (df_ent["Categoria"].str.strip().str.lower() == str(nueva_categoria).strip().lower())
        posibles = df_ent[mask]["Entregable"].dropna().unique().tolist()
@@ -255,19 +242,16 @@ def load_sheets():

    nuevo_entregable = st.selectbox("Entregable / Tarea", options=posibles)
    nota_descr = st.text_area("Descripción / Comentarios (libre)", value="", height=140)

    archivo = st.file_uploader("Subir archivo entregable (PDF/Word/Excel)", type=["pdf","docx","xlsx"], key="uploader")

with col_b:
    # Prioridad: intentar obtener opciones desde el sheet (para area+categoria), si no usar estándar
    prioridad_options = []
    if nueva_categoria and nueva_categoria != "(Sin categorías en Entregables)":
        pri = df_ent[ (df_ent["Area"].str.strip().str.lower() == area.strip().lower()) & (df_ent["Categoria"].str.strip().str.lower() == str(nueva_categoria).strip().lower()) ]["Prioridad"].dropna().unique().tolist()
        prioridad_options = [str(x) for x in pri if str(x).strip()]
    if not prioridad_options:
        prioridad_options = ["Baja","Media","Alta"]

    # Fecha compromiso: si el entregable existe en sheet, intentar prellenar fecha
    default_date = None
    if nuevo_entregable and nuevo_entregable not in ("(Sin entregables en esta categoría)","(Sin entregables disponibles)"):
        row_match = df_ent[
@@ -276,7 +260,6 @@ def load_sheets():
            (df_ent["Entregable"].str.strip().str.lower() == str(nuevo_entregable).strip().lower())
        ]
        if not row_match.empty:
            # intentar leer 'Fecha Compromiso' o 'Fecha Compromiso' variantes
            date_col = None
            for col_try in ["Fecha Compromiso","Fecha compromiso","Fecha Entrega","FechaEntrega","Fecha"]:
                if col_try in row_match.columns:
@@ -286,13 +269,11 @@ def load_sheets():
                val = row_match.iloc[0].get(date_col)
                if pd.notna(val) and val != "":
                    try:
                        # convertir con pandas
                        dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
                        if pd.notna(dt):
                            default_date = dt.date()
                    except Exception:
                        default_date = None
            # prioridad por defecto desde hoja
            try:
                p = row_match.iloc[0].get("Prioridad") or row_match.iloc[0].get("prioridad") or None
                if p and str(p).strip() and str(p).strip() in prioridad_options:
@@ -306,22 +287,13 @@ def load_sheets():
    else:
        prioridad_default = prioridad_options[0]

    # Mostrar los controles con valores por defecto si existían
    prioridad = st.selectbox("Prioridad", options=prioridad_options, index=max(0, prioridad_options.index(prioridad_default) if prioridad_default in prioridad_options else 0))
    # Fecha compromiso control
    if default_date:
        fecha_compromiso = st.date_input("Fecha compromiso", value=default_date)
    else:
        fecha_compromiso = st.date_input("Fecha compromiso", value=datetime.today().date())

    # Responsable (por defecto dueño del proceso)
    fecha_compromiso = st.date_input("Fecha compromiso", value=default_date if default_date else datetime.today().date())
    responsable = st.text_input("Responsable", value=info.get("Dueño del Proceso",""))

    # Estado: libre (texto)
    estado = st.text_input("Estado", value="Pendiente")

# ---------------------------
# CHAT / CONSULTA IA (asistente) - usa contexto del area
# CHAT / CONSULTA IA
# ---------------------------
st.subheader("💬 Consultar IA sobre cláusulas o entregables")
pregunta_ia = st.text_input("Escribe tu duda o consulta sobre ISO 9001 para tu área:", key="pregunta_ia")
@@ -351,19 +323,18 @@ def load_sheets():
            st.error(f"Ocurrió un error inesperado al consultar la IA: {e}")

# ---------------------------
# GUARDAR ENTREGABLE: subir a Drive y registrar SOLO en hoja Carga (A,E columnas exactas)
# GUARDAR ENTREGABLE
# ---------------------------
DRIVE_FOLDER_ID = "1ueBPvyVPoSkz0VoLXIkulnwLw3am3WYX"
drive_service = None
# inicializar drive_service solo si hay archivo o al intentar usar
if "service_account.json" in os.listdir(".") or "SERVICE_ACCOUNT_JSON" in st.secrets:

if "credentials.json" in os.listdir("."):
    try:
        drive_service = get_drive_service()
        
    except Exception:
        drive_service = None

def subir_archivo_drive(service_drive, archivo, carpeta_id):
    # archivo: st.uploaded_file object
    archivo_bytes = archivo.read()
    fh = io.BytesIO(archivo_bytes)
    media = MediaIoBaseUpload(fh, mimetype=archivo.type, resumable=False)
@@ -373,14 +344,12 @@ def subir_archivo_drive(service_drive, archivo, carpeta_id):

def hacer_publico(service_drive, file_id):
    try:
        # permiso público lectura
        service_drive.permissions().create(
            fileId=file_id,
            body={"role": "reader", "type": "anyone"},
            fields="id"
        ).execute()
    except Exception:
        # si falla por permisos insuficientes, seguimos sin detener el flujo
        pass

def asegurar_hoja_carga(sh_obj):
@@ -390,7 +359,6 @@ def asegurar_hoja_carga(sh_obj):
        try:
            sh_obj.add_worksheet(title="Carga", rows="1000", cols="10")
            ws = sh_obj.worksheet("Carga")
            # encabezado exacto solicitado
            ws.append_row(["Area","Categoria","Entregable","Fecha Entrega","Link Documento"])
        except Exception:
            ws = None
@@ -401,53 +369,47 @@ def asegurar_hoja_carga(sh_obj):
        st.warning("Selecciona un entregable válido antes de guardar.")
    else:
        file_link = ""
        # Subir archivo a Drive si se cargó
        if archivo is not None:
            try:
                if drive_service is None:
                    drive_service = get_drive_service()

                    drive_service = get_drive_service_oauth()
                file_id = subir_archivo_drive(drive_service, archivo, DRIVE_FOLDER_ID)
                # intentar hacer público (anyone with link)
                try:
                    hacer_publico(drive_service, file_id)
                except Exception:
                    pass

                file_link = f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"

                st.success("Archivo subido a Google Drive correctamente.")
                st.markdown(f"**Link del archivo:** {file_link}")

            except Exception as e:
                st.warning(f"Error subiendo archivo a Drive: {e}")
                st.error(f"No se pudo subir el archivo a Drive: {e}")
                file_link = ""

        # FORMATO de fecha a string
        fecha_str = fecha_compromiso.strftime("%Y-%m-%d")

        # Solo guardamos en la hoja "Carga" con EXACTAMENTE estas 5 columnas
        try:
            ws_carga = asegurar_hoja_carga(sh)
            if ws_carga:
                row_carga = [area, nueva_categoria, nuevo_entregable, fecha_str, file_link]
                ws_carga.append_row(row_carga)
                st.success("Registro añadido en hoja 'Carga'.")
            else:
                st.error("No se pudo acceder o crear la hoja 'Carga'. Verifica permisos.")
        except Exception as e:
            st.error(f"No se pudo registrar en hoja 'Carga': {e}")

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
    blue_style = ParagraphStyle('blue', parent=styles['Normal'], textColor=colors.blue)
    blue_style = ParagraphStyle(
        'blue', parent=styles['Normal'], textColor=colors.blue
    )
    story = []

    # Encabezado
    header_path = "assets/Encabezado.png"
    if os.path.exists(header_path):
        try:
@@ -456,21 +418,34 @@ def build_pdf_bytes(area, info, nuevo_entregable, nota_descr, resumen_ia=""):
        except Exception:
            pass

    # Título y datos del área
    story.append(Paragraph(f"<b>Reporte ISO 9001 — {area}</b>", blue_style))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"Dueño: {info.get('Dueño del Proceso')} — Puesto: {info.get('Puesto')} — Email: {info.get('Correo')}", blue_style))
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
@@ -479,25 +454,41 @@ def build_pdf_bytes(area, info, nuevo_entregable, nota_descr, resumen_ia=""):
        except Exception:
            pass

    # Construir PDF
    doc.build(story)
    buf.seek(0)
    return buf


# Botón para generar y descargar PDF
if st.button("📥 Generar y descargar PDF"):
    pdf_buf = build_pdf_bytes(area, info, nuevo_entregable, nota_descr)
    st.download_button("Descargar Reporte PDF", data=pdf_buf, file_name=f"Reporte_ISO_{area}.pdf", mime="application/pdf")
    st.download_button(
        "Descargar Reporte PDF",
        data=pdf_buf,
        file_name=f"Reporte_ISO_{area}.pdf",
        mime="application/pdf",
    )


# ---------------------------
# FOOTER
# FOOTER STREAMLIT
# ---------------------------
footer_img = load_image_try("assets/Pie.png") or load_image_try("Pie.png")
if footer_img:
    try:
        st.image(footer_img, width=800)
    except Exception:
        st.markdown("<div class='small' style='text-align:center;margin-top:20px;color:#0033cc;'>Formulario automatizado · Mantenimiento ISO · Generado con IA</div>", unsafe_allow_html=True)
        st.markdown(
            "<div class='small' style='text-align:center;margin-top:20px;color:#0033cc;'>"
            "Formulario automatizado · Mantenimiento ISO · Generado con IA</div>",
            unsafe_allow_html=True,
        )
else:
    st.markdown("<div class='small' style='text-align:center;margin-top:20px;color:#0033cc;'>Formulario automatizado · Mantenimiento ISO · Generado con IA</div>", unsafe_allow_html=True)

    st.markdown(
        "<div class='small' style='text-align:center;margin-top:20px;color:#0033cc;'>"
        "Formulario automatizado · Mantenimiento ISO · Generado con IA</div>",
        unsafe_allow_html=True,
    )

