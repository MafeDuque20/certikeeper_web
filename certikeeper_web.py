import streamlit as st
import fitz
import re
import pandas as pd
from zipfile import ZipFile
from io import BytesIO

# Diccionario de bases
base_abrev = {
    "SAN ANDRES": "ADZ",
    "ARMENIA": "AXM",
    "CALI": "CLO",
    "BARRANQUILLA": "BAQ",
    "BUCARAMANGA": "BGA",
    "SANTA MARTA": "SMR",
    "CARTAGENA": "CTG"
}

# Cursos válidos
cursos_validos = {
    "SMS ESP": "SMS ESP",
    "SEGURIDAD EN RAMPA PAX": "SEGURIDAD EN RAMPA PAX",
    "SEGURIDAD EN RAMPA OT": "SEGURIDAD EN RAMPA",
    "FACTORES HUMANOS": "FACTORES HUMANOS",
    "ER 201": "ER 201",
    "EQUIPAJES": "EQUIPAJES",
    "DESPACHO CENTRALIZADO": "DESPACHO",
    "ATENCIÓN A PASAJEROS": "ATENCIÓN A PASAJEROS"
}

# Palabras inválidas
palabras_invalidas = {"CARGO", "NEO", "AERO", "AGENTE", "SUPERVISOR", "COORDINADOR", "OPERADOR", "A", "PDE", "NEL", "EEE"}

# Función para extraer texto del PDF
def obtener_texto_con_ocr(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    for page in doc:
        texto += page.get_text()
    doc.close()
    return texto.upper()

def detectar_curso(texto):
    for linea in texto.splitlines():
        for clave in cursos_validos:
            if clave in linea:
                return cursos_validos[clave]
    return "CURSO"

def detectar_base(texto):
    for base in base_abrev:
        if base in texto:
            return base
    return "XXX"

def detectar_tipo(texto):
    return "OT" if "AGENTE DE RAMPA" in texto else "SAP"

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

def extraer_primer_nombre_apellido(nombre_completo):
    palabras = nombre_completo.strip().split()
    palabras_limpias = [p for p in palabras if p.isalpha() and p not in palabras_invalidas]

    if len(palabras_limpias) < 2:
        return None, None

    if len(palabras_limpias) >= 4:
        return palabras_limpias[0], palabras_limpias[2]
    else:
        return palabras_limpias[0], palabras_limpias[1]

def extraer_info(pdf_bytes):
    texto = obtener_texto_con_ocr(pdf_bytes)
    base = detectar_base(texto)
    curso = detectar_curso(texto)
    tipo = detectar_tipo(texto)
    nombre_completo = detectar_nombre_con_flexibilidad(texto)

    if not nombre_completo:
        return None, None, None, None, None, "ERROR: Sin nombre"

    primer_nombre, primer_apellido = extraer_primer_nombre_apellido(nombre_completo)

    if not primer_nombre or not primer_apellido:
        return None, None, None, None, None, "ERROR: Nombre inválido"

    base_ab = base_abrev.get(base, "XXX")
    nuevo_nombre = f"{base_ab} {curso} {tipo} {primer_nombre} {primer_apellido}".upper() + ".pdf"

    return base_ab, curso, tipo, f"{primer_nombre} {primer_apellido}", nuevo_nombre, "✅"

def separar_paginas_pdf(pdf_bytes, nombre_base):
    """ Separa cada página de un PDF en archivos PDF individuales en memoria. Retorna lista de tuplas: (nombre_generado, bytes_de_pdf) """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pdfs_individuales = []
    for i, page in enumerate(doc, start=1):
        nuevo_pdf = fitz.open()  # Crear PDF vacío
        nuevo_pdf.insert_pdf(doc, from_page=i-1, to_page=i-1)
        buffer = BytesIO()
        nuevo_pdf.save(buffer)
        buffer.seek(0)
        nombre = f"{nombre_base}_pagina{i}.pdf"
        pdfs_individuales.append((nombre, buffer.read()))
        nuevo_pdf.close()
    doc.close()
    return pdfs_individuales

def extraer_pdfs_de_archivos(uploaded_files):
    """ 
    Extrae todos los PDFs de los archivos subidos.
    Si es PDF directo, lo separa por páginas si tiene más de 1.
    Si es ZIP, extrae todos los PDFs dentro del ZIP y los separa por páginas.
    Retorna una lista de tuplas (nombre_original, contenido_pdf)
    """
    pdfs_extraidos = []

    for uploaded in uploaded_files:
        contenido = uploaded.read()

        if uploaded.name.lower().endswith(".pdf"):
            # PDF individual directo, separar por páginas
            pdfs_extraidos.extend(separar_paginas_pdf(contenido, uploaded.name.replace(".pdf","")))

        elif uploaded.name.lower().endswith(".zip"):
            # Extraer PDFs del ZIP y separar páginas
            try:
                with ZipFile(BytesIO(contenido)) as zipf:
                    archivos_en_zip = zipf.namelist()
                    for nombre_archivo in archivos_en_zip:
                        if nombre_archivo.lower().endswith(".pdf"):
                            pdf_bytes = zipf.read(nombre_archivo)
                            pdfs_extraidos.extend(separar_paginas_pdf(pdf_bytes, nombre_archivo.replace(".pdf","")))
            except Exception as e:
                st.warning(f"⚠️ Error al procesar ZIP '{uploaded.name}': {str(e)}")

    return pdfs_extraidos


# INTERFAZ STREAMLIT
st.title("📋 CertiKeeper Web")
st.write("Sube tus archivos PDF o ZIP con certificados. Cada página será procesada como un PDF individual.")
st.write("Cada certificado será renombrado según su contenido.")

uploaded_files = st.file_uploader(
    "Sube tus archivos (PDF o ZIP)", accept_multiple_files=True, type=["pdf", "zip"]
)

if uploaded_files:
    st.info(f"📦 Procesando {len(uploaded_files)} archivo(s) subido(s)...")

    all_pdfs = extraer_pdfs_de_archivos(uploaded_files)

    if not all_pdfs:
        st.error("❌ No se encontraron archivos PDF para procesar.")
    else:
        st.success(f"✅ Se encontraron {len(all_pdfs)} página(s) PDF para procesar")

        # Mostrar los PDFs extraídos
        with st.expander("📄 Ver páginas PDF extraídas"):
            for i, (nombre, _) in enumerate(all_pdfs, 1):
                st.text(f"{i}. {nombre}")

        # Procesar cada página PDF individualmente
        log = []
        renombrados = []
        errores = 0

        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, (nombre_original, pdf_bytes) in enumerate(all_pdfs):
            progress = (idx + 1) / len(all_pdfs)
            progress_bar.progress(progress)
            status_text.text(f"Procesando {idx + 1}/{len(all_pdfs)}: {nombre_original}")

            base, curso, tipo, alumno, nuevo_nombre, estado = extraer_info(pdf_bytes)

            if estado.startswith("ERROR"):
                errores += 1
                log.append({
                    "Archivo original": nombre_original,
                    "Estado": estado,
                    "Nombre final": "N/A",
                    "Base": "N/A",
                    "Curso": "N/A",
                    "Tipo": "N/A",
                    "Alumno": "N/A"
                })
                continue

            renombrados.append((nuevo_nombre, pdf_bytes))

            log.append({
                "Archivo original": nombre_original,
                "Estado": estado,
                "Nombre final": nuevo_nombre,
                "Base": base,
                "Curso": curso,
                "Tipo": tipo,
                "Alumno": alumno
            })

        progress_bar.empty()
        status_text.empty()

        # Mostrar resultados
        st.write("---")
        st.subheader("📊 Resultados del procesamiento")

        col1, col2, col3 = st.columns(3)
        col1.metric("Total procesados", len(all_pdfs))
        col2.metric("Exitosos", len(renombrados))
        col3.metric("Con errores", errores)

        df_log = pd.DataFrame(log)
        st.dataframe(df_log, use_container_width=True)

        # Opciones de descarga
        st.write("---")
        st.subheader("📥 Descargas")

        col1, col2 = st.columns(2)

        # Descargar Excel
        with col1:
            excel_buffer = BytesIO()
            df_log.to_excel(excel_buffer, index=False, engine='openpyxl')
            excel_buffer.seek(0)
            st.download_button(
                "📊 Descargar log Excel",
                data=excel_buffer,
                file_name="log_certificados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Descargar ZIP
        with col2:
            if renombrados:
                zip_buffer = BytesIO()
                with ZipFile(zip_buffer, "w") as zipf:
                    for nombre, contenido in renombrados:
                        zipf.writestr(nombre, contenido)
                zip_buffer.seek(0)

                st.download_button(
                    "📦 Descargar ZIP con certificados renombrados",
                    data=zip_buffer,
                    file_name="Certificados_Renombrados.zip",
                    mime="application/zip"
                )
            else:
                st.warning("⚠️ No hay certificados exitosos para descargar")

        # Mostrar errores
        if errores > 0:
            with st.expander(f"⚠️ Ver {errores} error(es)"):
                df_errores = df_log[df_log['Estado'].str.startswith('ERROR')]
                st.dataframe(df_errores, use_container_width=True)
