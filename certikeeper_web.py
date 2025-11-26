# certikeeper_web.py
import streamlit as st
import fitz
import os
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
    "CARGO", "NEO", "AERO", "AGENTE", "SUPERVISOR", "COORDINADOR", "OPERADOR",
    "A", "PDE", "NEL", "EEE"
}

# 📄 Extraer texto del PDF
def obtener_texto_con_ocr(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    for page in doc:
        texto += page.get_text()
    doc.close()
    return texto.upper()

# Detectores
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

# Extrae toda la info
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

    return base_ab, curso, tipo, f"{primer_nombre} {primer_apellido}", nuevo_nombre, "OK"

# ---------------------------
#          STREAMLIT
# ---------------------------
st.title("CertiKeeper Web")
st.write("Sube PDFs o ZIPs con certificados para renombrarlos automáticamente.")

uploaded_files = st.file_uploader(
    "Sube archivos PDF o ZIP", accept_multiple_files=True, type=["pdf", "zip"]
)

if uploaded_files:
    log = []
    certificados_por_base = {}  # <-- ← Agrupación por base

    for uploaded in uploaded_files:
        if uploaded.type == "application/zip":
            # Extraer PDFs dentro del ZIP
            with ZipFile(uploaded) as z:
                for nombre in z.namelist():
                    if nombre.lower().endswith(".pdf"):
                        pdf_bytes = z.read(nombre)

                        # Procesar PDF extraído
                        base, curso, tipo, alumno, nuevo_nombre, estado = extraer_info(pdf_bytes)

                        if estado != "OK":
                            log.append({"Archivo original": nombre, "Estado": estado})
                            continue

                        # Guardar en carpeta por base
                        certificados_por_base.setdefault(base, []).append((nuevo_nombre, pdf_bytes))

                        log.append({
                            "Archivo original": nombre,
                            "Nombre final": nuevo_nombre,
                            "Base": base,
                            "Curso": curso,
                            "Tipo": tipo,
                            "Alumno": alumno,
                            "Estado": estado
                        })

        else:  # PDF suelto
            pdf_bytes = uploaded.read()
            base, curso, tipo, alumno, nuevo_nombre, estado = extraer_info(pdf_bytes)

            if estado != "OK":
                log.append({"Archivo original": uploaded.name, "Estado": estado})
                continue

            certificados_por_base.setdefault(base, []).append((nuevo_nombre, pdf_bytes))

            log.append({
                "Archivo original": uploaded.name,
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
    st.download_button("📥 Descargar log Excel", excel_buffer.getvalue(),
                       file_name="log_certificados.xlsx")

    # Crear ZIP final con carpetas por base
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "w") as zipf:
        for base, archivos in certificados_por_base.items():
            carpeta = f"{base}/"
            for nombre, contenido in archivos:
                zipf.writestr(carpeta + nombre, contenido)

    st.download_button("📥 Descargar ZIP organizado por base",
                       zip_buffer.getvalue(),
                       file_name="Certificados_Por_Base.zip")
