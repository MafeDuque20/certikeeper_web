import streamlit as st
import fitz
import re
import pandas as pd
from zipfile import ZipFile
from io import BytesIO

# =========================
#     DICCIONARIOS BASE
# =========================

base_abrev = {
    "SAN ANDRES": "ADZ", "ARMENIA": "AXM", "CALI": "CLO",
    "BARRANQUILLA": "BAQ", "BUCARAMANGA": "BGA",
    "SANTA MARTA": "SMR", "CARTAGENA": "CTG"
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

palabras_invalidas = {"CARGO", "NEO", "AERO", "AGENTE", "SUPERVISOR",
                      "COORDINADOR", "OPERADOR", "A", "PDE", "NEL", "EEE"}

# =========================
#      OCR PDF
# =========================
def obtener_texto_con_ocr(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    for page in doc:
        texto += page.get_text()
    doc.close()
    return texto.upper()

# =========================
#   DETECCIÓN DE CURSO
# =========================
def detectar_curso(texto):
    for linea in texto.splitlines():
        for clave in cursos_validos:
            if clave in linea:
                return cursos_validos[clave]
    return "CURSO"

# =========================
#   DETECCIÓN DE BASE
# =========================
def detectar_base(texto):
    for base in base_abrev:
        if base in texto:
            return base
    return "XXX"

# =========================
#   DETECCIÓN DE CARGO (OT / SAP)
# =========================
def detectar_tipo(texto):
    texto = texto.upper()

    # Palabras clave fuertes OT
    claves_ot = [
        "OT", "OPERACIONES TERRESTRES", "RAMPA", "AGENTE DE RAMPA",
        "OPERADOR DE RAMPA", "OPERARIO", "OPERACIÓN TERRESTRE"
    ]

    # Palabras clave fuertes SAP
    claves_sap = [
        "SAP", "PAX", "PASAJEROS", "SERVICIO AL PASAJERO",
        "ATENCIÓN A PASAJEROS", "CHECK IN", "PASAJERO"
    ]

    # Detección OT
    for palabra in claves_ot:
        if palabra in texto:
            return "OT"

    # Detección SAP
    for palabra in claves_sap:
        if palabra in texto:
            return "SAP"

    # DEFAULT → SAP (porque SAP tiende a aparecer más en texto ambiguo)
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

def extraer_primer_nombre_apellido(nombre_completo):
    palabras = nombre_completo.strip().split()
    palabras_limpias = [p for p in palabras if p.isalpha() and p not in palabras_invalidas]
    if len(palabras_limpias) < 2:
        return None, None
    if len(palabras_limpias) >= 4:
        return palabras_limpias[0], palabras_limpias[2]
    else:
        return palabras_limpias[0], palabras_limpias[1]

# =========================
# EXTRAER INFORMACIÓN
# =========================
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

    # ⚠️ CURSO SEGURIDAD EN RAMPA → no agregar PAX u OT al nombre
    if "SEGURIDAD EN RAMPA" in curso:
        nuevo_nombre = f"{base_ab} {curso} {tipo} {primer_nombre} {primer_apellido}".upper() + ".pdf"
    else:
        nuevo_nombre = f"{base_ab} {curso} {tipo} {primer_nombre} {primer_apellido}".upper() + ".pdf"

    return base_ab, curso, tipo, f"{primer_nombre} {primer_apellido}", nuevo_nombre, "✅"

# =========================
# SEPARAR PDF EN PÁGINAS
# =========================
def separar_paginas_pdf(pdf_bytes, nombre_origen):
    paginas_individuales = []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for num_pagina in range(len(doc)):
            nuevo_doc = fitz.open()
            nuevo_doc.insert_pdf(doc, from_page=num_pagina, to_page=num_pagina)
            pdf_buffer = BytesIO()
            nuevo_doc.save(pdf_buffer, garbage=4, deflate=True, clean=True, incremental=False)
            pdf_buffer.seek(0)
            paginas_individuales.append((f"{nombre_origen}_pag_{num_pagina+1}", pdf_buffer.read()))
            nuevo_doc.close()
        doc.close()
    except Exception as e:
        st.warning(f"Error separando '{nombre_origen}': {str(e)}")
    return paginas_individuales

# =========================
# EXTRAER PDFs DE ZIP O PDF
# =========================
def extraer_pdfs_de_archivos(uploaded_files):
    pdfs_extraidos = []
    for uploaded in uploaded_files:
        contenido = uploaded.read()
        if uploaded.name.lower().endswith(".pdf"):
            nombre_base = uploaded.name.replace(".pdf", "")
            pdfs_extraidos.extend(separar_paginas_pdf(contenido, nombre_base))
        elif uploaded.name.lower().endswith(".zip"):
            try:
                with ZipFile(BytesIO(contenido)) as zipf:
                    for nombre_archivo in zipf.namelist():
                        if nombre_archivo.lower().endswith(".pdf"):
                            pdf_bytes = zipf.read(nombre_archivo)
                            nombre_base = nombre_archivo.replace(".pdf", "")
                            pdfs_extraidos.extend(separar_paginas_pdf(pdf_bytes, nombre_base))
            except Exception as e:
                st.warning(f"Error ZIP '{uploaded.name}': {str(e)}")
    return pdfs_extraidos

# =========================
#      STREAMLIT UI
# =========================
st.title("RENOMBRAR CERTIFICADOS")
st.write("Cada página del PDF se convertirá en un certificado individual.")

uploaded_files = st.file_uploader(
    "📂 Sube PDFs o ZIP con certificados",
    accept_multiple_files=True,
    type=["pdf", "zip"]
)

if uploaded_files:
    st.info(f"Procesando {len(uploaded_files)} archivo(s)...")
    all_pdfs = extraer_pdfs_de_archivos(uploaded_files)

    if not all_pdfs:
        st.error("❌ No se encontraron páginas PDF.")
    else:
        st.success(f"Se extrajeron {len(all_pdfs)} páginas.")

        log = []
        renombrados = []
        errores = 0

        progress = st.progress(0)

        for i, (nombre_original, pdf_bytes) in enumerate(all_pdfs):
            progress.progress((i+1)/len(all_pdfs))

            base, curso, tipo, alumno, nuevo_nombre, estado = extraer_info(pdf_bytes)

            if estado.startswith("ERROR"):
                errores += 1
                log.append({
                    "Página original": nombre_original,
                    "Estado": estado,
                    "Nombre final": "",
                    "Base": "",
                    "Curso": "",
                    "Tipo": "",
                    "Alumno": ""
                })
                continue

            renombrados.append((nuevo_nombre, pdf_bytes))

            log.append({
                "Página original": nombre_original,
                "Estado": estado,
                "Nombre final": nuevo_nombre,
                "Base": base,
                "Curso": curso,
                "Tipo": tipo,
                "Alumno": alumno
            })

        st.write("### 📊 Resultados")
        st.dataframe(pd.DataFrame(log))

        # DESCARGA ZIP PRINCIPAL
        if renombrados:
            zip_buffer = BytesIO()
            with ZipFile(zip_buffer, "w") as zipf:
                for nombre, contenido in renombrados:
                    zipf.writestr(nombre, contenido)
            zip_buffer.seek(0)
            st.download_button(
                "📦 Descargar certificados renombrados",
                data=zip_buffer,
                file_name="certificados.zip"
            )

        # AGRUPAR POR BASE + CARGO
        if renombrados:
            zip_bases = BytesIO()
            with ZipFile(zip_bases, "w") as z:
                for nombre, contenido in renombrados:
                    partes = nombre.split()

                    base = partes[0]
                    tipo = partes[2]  # OT / SAP
                    carpeta = "RAMPA" if tipo == "OT" else "PAX"

                    ruta = f"{base}/{carpeta}/{nombre}"
                    z.writestr(ruta, contenido)

            zip_bases.seek(0)
            st.download_button(
                "📂 Descargar ZIP organizado por BASE y CARGO",
                zip_bases,
                "Certificados_Por_Base_Cargo.zip"
            )
