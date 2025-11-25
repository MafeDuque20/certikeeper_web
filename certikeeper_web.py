# certikeeper_web.py
import streamlit as st
import fitz
import re
import pandas as pd
from zipfile import ZipFile
from io import BytesIO

# 🗂️ Diccionario de bases
base_abrev = {
    "SAN ANDRES": "ADZ", "ARMENIA": "AXM", "CALI": "CLO", "BARRANQUILLA": "BAQ",
    "BUCARAMANGA": "BGA", "SANTA MARTA": "SMR", "CARTAGENA": "CTG"
}

# 🎓 Cursos válidos
cursos_validos = {
    "SMS ESP": "SMS ESP", "SEGURIDAD EN RAMPA PAX": "SEGURIDAD EN RAMPA PAX",
    "SEGURIDAD EN RAMPA OT": "SEGURIDAD EN RAMPA", "FACTORES HUMANOS": "FACTORES HUMANOS",
    "ER 201": "ER 201", "EQUIPAJES": "EQUIPAJES", "DESPACHO CENTRALIZADO": "DESPACHO",
    "ATENCIÓN A PASAJEROS": "ATENCIÓN A PASAJEROS"
}

# ❌ Palabras inválidas
palabras_invalidas = {
    "CARGO", "NEO", "AERO", "AGENTE", "SUPERVISOR", "COORDINADOR", "OPERADOR", "A",
    "PDE", "NEL", "EEE"
}

# 📄 Función para extraer texto del PDF
def obtener_texto_con_ocr(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    for page in doc:
        texto += page.get_text()
    doc.close()
    return texto.upper()

# 🎯 Funciones de detección
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

# ✂️ Extraer primer nombre y primer apellido
def extraer_primer_nombre_apellido(nombre_completo):
    palabras = nombre_completo.strip().split()
    palabras_limpias = [p for p in palabras if p.isalpha() and p not in palabras_invalidas]
    if len(palabras_limpias) < 2:
        return None, None
    if len(palabras_limpias) >= 4:
        return palabras_limpias[0], palabras_limpias[2]
    else:
        return palabras_limpias[0], palabras_limpias[1]

# 🏷️ Extraer toda la información
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

# 🌐 Interfaz con Streamlit
st.title("CertiKeeper Web")
st.write("Sube tus archivos PDF o ZIP y el sistema renombrará cada certificado dentro de los ZIP individualmente.")

uploaded_files = st.file_uploader(
    "Sube tus archivos", accept_multiple_files=True, type=["pdf","zip"]
)

# Lista para todos los PDFs a procesar
all_pdfs = []

if uploaded_files:
    for uploaded in uploaded_files:
        nombre_archivo = uploaded.name.lower()
        contenido = uploaded.read()

        # Si es PDF individual
        if nombre_archivo.endswith(".pdf"):
            all_pdfs.append((uploaded.name, contenido))

        # Si es ZIP, extraer todos los PDFs dentro
        elif nombre_archivo.endswith(".zip"):
            with ZipFile(BytesIO(contenido)) as zipf:
                for f in zipf.namelist():
                    if f.lower().endswith(".pdf"):
                        pdf_bytes = zipf.read(f)
                        # Guardar como (nombre original dentro del zip, contenido)
                        all_pdfs.append((f, pdf_bytes))

# Procesar cada PDF individualmente
if all_pdfs:
    log = []
    renombrados = []

    for nombre_original, pdf_bytes in all_pdfs:
        base, curso, tipo, alumno, nuevo_nombre, estado = extraer_info(pdf_bytes)

        if estado.startswith("ERROR"):
            log.append({"Archivo original": nombre_original, "Estado": estado})
            continue

        renombrados.append((nuevo_nombre, pdf_bytes))
        log.append({
            "Archivo original": nombre_original,
            "Nombre final": nuevo_nombre,
            "Base": base,
            "Curso": curso,
            "Tipo": tipo,
            "Alumno": alumno,
            "Estado": estado
        })

    # Mostrar log
    df_log = pd.DataFrame(log)
    st.dataframe(df_log)

    # Descargar Excel log
    excel_buffer = BytesIO()
    df_log.to_excel(excel_buffer, index=False)
    st.download_button("📥 Descargar log Excel", excel_buffer, file_name="log_certificados.xlsx")

    # Crear ZIP final con todos los PDFs renombrados
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "w") as zipf:
        for nombre, contenido in renombrados:
            zipf.writestr(nombre, contenido)
    st.download_button("📥 Descargar ZIP certificados", zip_buffer, file_name="Certificados_Renombrados.zip")
