import streamlit as st
import fitz
import re
import pandas as pd
from zipfile import ZipFile
from io import BytesIO
from datetime import datetime

# =========================
# CONFIGURACIÓN DE PÁGINA
# =========================
st.set_page_config(
    page_title="Renombrador de Certificados",
    page_icon="📜",
    layout="wide"
)

# =========================
# DICCIONARIOS BASE
# =========================
base_abrev = {
    "SAN ANDRES": "ADZ",
    "ARMENIA": "AXM",
    "CALI": "CLO",
    "BARRANQUILLA": "BAQ",
    "BUCARAMANGA": "BGA",
    "SANTA MARTA": "SMR",
    "CARTAGENA": "CTG",
    "PEREIRA": "PEI"
}

cursos_validos = {
    "SMS ESP": "SMS ESP",
    "SEGURIDAD EN RAMPA PAX": "SEGURIDAD EN RAMPA",
    "SEGURIDAD EN RAMPA OT": "SEGURIDAD EN RAMPA",
    "FACTORES HUMANOS": "FACTORES HUMANOS",
    "ER 201": "ER 201",
    "EQUIPAJES": "EQUIPAJES",
    "DESPACHO CENTRALIZADO": "DESPACHO",
    "ATENCIÓN A PASAJEROS": "ATENCIÓN A PASAJEROS",
    "BRS": "BRS",
    "MODELO DE EXPERIENCIA": "MODELO DE EXPERIENCIA",
    "PROCESOS PARA LA ATENCION DE AERONAVE": "PROCESOS PARA LA ATENCION DE AERONAVE"
}

# =========================
# OCR PDF
# =========================
def obtener_texto_con_ocr(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = "".join([page.get_text() for page in doc])
    doc.close()
    return texto.upper()

# =========================
# DETECCIÓN DE CURSO
# =========================
def detectar_curso(texto):
    for linea in texto.splitlines():
        for clave in cursos_validos:
            if clave in linea:
                return cursos_validos[clave]
    return "CURSO"

# =========================
# DETECCIÓN DE BASE
# =========================
def detectar_base(texto):
    for base in base_abrev:
        if base in texto:
            return base
    return "XXX"

# =========================
# DETECCIÓN DE CARGO
# =========================
def detectar_tipo(texto):
    texto = texto.upper()
    claves_ot = ["OT", "OPERACIONES TERRESTRES", "AGENTE DE RAMPA", "OPERADOR DE RAMPA", "OPERARIO", "OPERACIÓN TERRESTRE"]
    claves_sap = ["SAP", "PAX", "PASAJEROS", "SERVICIO AL PASAJERO", "ATENCIÓN A PASAJEROS", "CHECK IN", "PASAJERO"]
    for palabra in claves_ot:
        if palabra in texto:
            return "OT"
    for palabra in claves_sap:
        if palabra in texto:
            return "SAP"
    return "SAP"

# =========================
# DETECCIÓN DE NOMBRE
# =========================
def detectar_nombre_con_flexibilidad(texto):
    patrones = [
        r"NOMBRE\s+DEL\s+ALUMNO\s*:?[\s]*([A-Z\s]{5,})\s+IDENTIFICACIÓN",
        r"NOMBRE\s+ALUMNO\s*:?[\s]*([A-Z\s]{5,})\s+IDENTIFICACIÓN",
        r"NOMBRE\s+DEL\s+ALUMNO\s*:?[\s]*([A-Z\s]{5,})"
    ]
    for patron in patrones:
        coincidencias = re.findall(patron, texto)
        for match in coincidencias:
            posible = match.strip()
            if len(posible.split()) >= 2:
                return posible
    return ""

# =========================
# EXTRAER PRIMER NOMBRE + PRIMER APELLIDO
# =========================
def extraer_primer_nombre_apellido(nombre_completo):
    if not nombre_completo: return None, None
    limpio = " ".join(nombre_completo.replace("\n"," ").replace("-"," ").split())
    partes = limpio.split()
    if len(partes)<2: return None, None
    return partes[0], partes[1]

# =========================
# EXTRAER INFORMACIÓN
# =========================
def extraer_info(pdf_bytes):
    texto = obtener_texto_con_ocr(pdf_bytes)
    base = detectar_base(texto)

    # Detectar curso RAMPA PAX/OT
    curso_detectado = None
    for c in ["SEGURIDAD EN RAMPA PAX", "SEGURIDAD EN RAMPA OT"]:
        if c in texto:
            curso_detectado = c
            break

    nombre_completo = detectar_nombre_con_flexibilidad(texto)
    if not nombre_completo:
        return None, None, None, None, None, "ERROR: Sin nombre"
    primer_nombre, primer_apellido = extraer_primer_nombre_apellido(nombre_completo)
    if not primer_nombre or not primer_apellido:
        return None, None, None, None, None, "ERROR: Nombre inválido"
    base_ab = base_abrev.get(base, "XXX")

    if curso_detectado:  # RAMPA PAX/OT
        nuevo_nombre = f"{base_ab} {curso_detectado} {primer_nombre} {primer_apellido}".upper() + ".pdf"
        tipo = ""  # no se usa tipo
        return base_ab, curso_detectado, tipo, f"{primer_nombre} {primer_apellido}", nuevo_nombre, "✅"
    
    # Cursos normales
    curso = detectar_curso(texto)
    tipo = detectar_tipo(texto)
    nuevo_nombre = f"{base_ab} {curso} {tipo} {primer_nombre} {primer_apellido}".upper() + ".pdf"
    return base_ab, curso, tipo, f"{primer_nombre} {primer_apellido}", nuevo_nombre, "✅"

# =========================
# SEPARAR PDF EN PÁGINAS
# =========================
def separar_paginas_pdf(pdf_bytes, nombre_origen):
    paginas = []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for i in range(len(doc)):
            nuevo_doc = fitz.open()
            nuevo_doc.insert_pdf(doc, from_page=i, to_page=i)
            buffer = BytesIO()
            nuevo_doc.save(buffer, garbage=4, deflate=True, clean=True, incremental=False)
            buffer.seek(0)
            paginas.append((f"{nombre_origen}_pag_{i+1}", buffer.read()))
            nuevo_doc.close()
        doc.close()
    except Exception as e:
        st.warning(f"Error separando '{nombre_origen}': {str(e)}")
    return paginas

# =========================
# EXTRAER PDFs DE ZIP O PDF
# =========================
def extraer_pdfs_de_archivos(uploaded_files):
    pdfs = []
    for uploaded in uploaded_files:
        contenido = uploaded.read()
        if uploaded.name.lower().endswith(".pdf"):
            nombre_base = uploaded.name.replace(".pdf","")
            pdfs.extend(separar_paginas_pdf(contenido, nombre_base))
        elif uploaded.name.lower().endswith(".zip"):
            try:
                with ZipFile(BytesIO(contenido)) as zipf:
                    for nombre_archivo in zipf.namelist():
                        if nombre_archivo.lower().endswith(".pdf"):
                            pdf_bytes = zipf.read(nombre_archivo)
                            nombre_base = nombre_archivo.replace(".pdf","")
                            pdfs.extend(separar_paginas_pdf(pdf_bytes, nombre_base))
            except Exception as e:
                st.warning(f"Error ZIP '{uploaded.name}': {str(e)}")
    return pdfs

# =========================
# CREAR ZIP ORGANIZADO
# =========================
def crear_zip_organizado(renombrados_info):
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer,"w") as zipf:
        for info in renombrados_info:
            nuevo_nombre, pdf_bytes, tipo, base = info["Nombre final"], info["Contenido"], info["Cargo"], info["Base"]
            # Carpeta según cargo
            if tipo.upper() == "OT":
                carpeta_tipo = "RAMPA"
            elif tipo.upper() == "SAP":
                carpeta_tipo = "PAX"
            else:
                carpeta_tipo = "OTROS"
            ruta_zip = f"{base}/{carpeta_tipo}/{nuevo_nombre}"
            zipf.writestr(ruta_zip, pdf_bytes)
    zip_buffer.seek(0)
    return zip_buffer

# =========================
# STREAMLIT UI
# =========================
st.markdown("<h1 style='text-align:center;color:#1f77b4;'>📜 RENOMBRADOR DE CERTIFICADOS</h1>", unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

with st.sidebar:
    st.header("ℹ️ Información")
    st.markdown("""
    ### Bases disponibles:
    ADZ, AXM, CLO, BAQ, BGA, SMR, CTG, PEI
    ### Tipos de cargo:
    OT: Operaciones Terrestres
    SAP: Servicio al Pasajero
    """)
    st.info(f"Fecha: {datetime.now().strftime('%d/%m/%Y')}")

uploaded_files = st.file_uploader("Sube tus PDFs o ZIP", accept_multiple_files=True, type=["pdf","zip"])

if uploaded_files:
    with st.spinner("Procesando archivos..."):
        all_pdfs = extraer_pdfs_de_archivos(uploaded_files)
    if not all_pdfs:
        st.error("No se encontraron PDFs.")
    else:
        log, renombrados_info, errores = [], [], 0
        progress = st.progress(0)
        status_text = st.empty()
        for i, (nombre_original, pdf_bytes) in enumerate(all_pdfs):
            progress.progress((i+1)/len(all_pdfs))
            status_text.text(f"Procesando {i+1}/{len(all_pdfs)}")
            base, curso, tipo, alumno, nuevo_nombre, estado = extraer_info(pdf_bytes)
            if estado.startswith("ERROR"):
                errores += 1
                log.append({"Página original": nombre_original,"Estado":estado})
                continue
            renombrados_info.append({"Nombre final":nuevo_nombre,"Contenido":pdf_bytes,"Cargo":tipo,"Base":base})
            log.append({"Página original":nombre_original,"Estado":estado,"Nombre final":nuevo_nombre,"Base":base,"Curso":curso,"Tipo":tipo,"Alumno":alumno})
        progress.empty(); status_text.empty()
        
        st.success(f"Procesados {len(all_pdfs)} páginas. Errores: {errores}")
        df_log = pd.DataFrame(log)
        st.dataframe(df_log, use_container_width=True, height=400)

        col1, col2 = st.columns(2)
        with col1:
            zip_buffer = crear_zip_organizado(renombrados_info)
            st.download_button("📦 Descargar ZIP", zip_buffer, file_name=f"certificados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip", mime="application/zip")
        with col2:
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                df_log.to_excel(writer,index=False,sheet_name="Reporte")
            excel_buffer.seek(0)
            st.download_button("📊 Descargar Excel", excel_buffer, file_name=f"reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Sube archivos PDF o ZIP para iniciar el procesamiento.")
