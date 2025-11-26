import streamlit as st
import fitz
import re
import pandas as pd
from zipfile import ZipFile
from io import BytesIO

# Diccionario de bases
base_abrev = {
    "SAN ANDRES": "ADZ", "ARMENIA": "AXM", "CALI": "CLO", "BARRANQUILLA": "BAQ",
    "BUCARAMANGA": "BGA", "SANTA MARTA": "SMR", "CARTAGENA": "CTG"
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
    "ATENCIÓN A PASAJEROS": "ATENCIÓN A PASAJEROS",
    "BRS": "BRS",
    "MODELO DE EXPERIENCIA": "MODELO DE EXPERIENCIA",
    "PROCESOS PARA LA ATENCION DE AERONAVE": "PROCESOS PARA LA ATENCION DE AERONAVE"
}

# Palabras inválidas
palabras_invalidas = {"CARGO", "NEO", "AERO", "AGENTE", "SUPERVISOR",
                      "COORDINADOR", "OPERADOR", "A", "PDE", "NEL", "EEE"}

# === OCR PDF ===
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

# === SEPARAR PDF (COMPATIBLE WINDOWS PREVIEW) ===
def separar_paginas_pdf(pdf_bytes, nombre_origen):
    paginas_individuales = []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for num_pagina in range(len(doc)):
            nuevo_doc = fitz.open()
            nuevo_doc.insert_pdf(doc, from_page=num_pagina, to_page=num_pagina)
            pdf_buffer = BytesIO()
            nuevo_doc.save(pdf_buffer, garbage=4, deflate=True, clean=True, incremental=False, ascii=False)
            pdf_buffer.seek(0)
            paginas_individuales.append((f"{nombre_origen}_pag_{num_pagina+1}", pdf_buffer.read()))
            nuevo_doc.close()
        doc.close()
    except Exception as e:
        st.warning(f"Error al separar páginas de '{nombre_origen}': {str(e)}")
    return paginas_individuales

# === Extraer PDFs desde PDF o ZIP ===
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
                st.warning(f"Error al procesar ZIP '{uploaded.name}': {str(e)}")
    return pdfs_extraidos

# === INTERFAZ STREAMLIT ===
st.title("CERTIFICADOS")
st.write("Sube tus archivos PDF o ZIP con certificados.")
st.write("**Cada página de cada PDF se convertirá en un certificado individual** y será renombrado según su contenido.")

uploaded_files = st.file_uploader(
    "Sube tus archivos (PDF o ZIP)",
    accept_multiple_files=True,
    type=["pdf", "zip"],
    help="Los PDFs se separarán página por página automáticamente"
)

if uploaded_files:
    st.info(f"📦 Procesando {len(uploaded_files)} archivo(s) subido(s)...")
    all_pdfs = extraer_pdfs_de_archivos(uploaded_files)

    if not all_pdfs:
        st.error("❌ No se encontraron páginas PDF para procesar.")
    else:
        st.success(f"✅ Se extrajeron {len(all_pdfs)} página(s) individual(es)")

        with st.expander("📄 Ver páginas PDF extraídas"):
            for i, (nombre, _) in enumerate(all_pdfs, 1):
                st.text(f"{i}. {nombre}")

        log = []
        renombrados = []
        errores = 0
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, (nombre_original, pdf_bytes) in enumerate(all_pdfs):
            progress = (idx + 1) / len(all_pdfs)
            progress_bar.progress(progress)
            status_text.text(f"Procesando {idx+1}/{len(all_pdfs)}: {nombre_original}")

            base, curso, tipo, alumno, nuevo_nombre, estado = extraer_info(pdf_bytes)
            if estado.startswith("ERROR"):
                errores += 1
                log.append({
                    "Página original": nombre_original,
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
                "Página original": nombre_original,
                "Estado": estado,
                "Nombre final": nuevo_nombre,
                "Base": base,
                "Curso": curso,
                "Tipo": tipo,
                "Alumno": alumno
            })

        progress_bar.empty()
        status_text.empty()

        st.write("---")
        st.subheader("📊 Resultados del procesamiento")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total páginas", len(all_pdfs))
        col2.metric("Exitosos", len(renombrados))
        col3.metric("Con errores", errores)
        df_log = pd.DataFrame(log)
        st.dataframe(df_log, use_container_width=True)

        st.write("---")
        st.subheader("📥 Descargas")
        col1, col2 = st.columns(2)

        with col1:
            excel_buffer = BytesIO()
            df_log.to_excel(excel_buffer, index=False, engine='openpyxl')
            excel_buffer.seek(0)
            st.download_button(
                label="📊 Descargar log Excel",
                data=excel_buffer,
                file_name="log_certificados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col2:
            if renombrados:
                zip_buffer = BytesIO()
                with ZipFile(zip_buffer, "w") as zipf:
                    for nombre, contenido in renombrados:
                        zipf.writestr(nombre, contenido)
                zip_buffer.seek(0)
                st.download_button(
                    label="📦 Descargar ZIP con certificados renombrados",
                    data=zip_buffer,
                    file_name="Certificados_Renombrados.zip",
                    mime="application/zip"
                )
            else:
                st.warning("⚠️ No hay certificados exitosos para descargar")

        # === ZIP por bases, cargos y repetidos ===
        st.write("---")
        st.subheader("📂 Descargar ZIP por bases, cargos y repetidos")
        if renombrados:
            zip_bases_buffer = BytesIO()
            with ZipFile(zip_bases_buffer, "w") as zipf:
                certificados_guardados = {}
                for nombre, contenido in renombrados:
                    partes = nombre.split()
                    base_actual = partes[0].upper()
                    tipo_cargo = partes[2].upper()
                    alumno = " ".join(partes[3:5]).upper()
                    curso = " ".join(partes[1:2]).upper()

                    carpeta_cargo = "RAMPA" if tipo_cargo == "OT" else "PAX"
                    clave = (base_actual, curso, alumno)

                    if clave not in certificados_guardados:
                        certificados_guardados[clave] = 0
                    certificados_guardados[clave] += 1

                    if certificados_guardados[clave] == 1:
                        ruta = f"{base_actual}/{carpeta_cargo}/{nombre}"
                    else:
                        ruta = f"REPETIDOS/{base_actual}/{carpeta_cargo}/{nombre}"

                    zipf.writestr(ruta, contenido)

            zip_bases_buffer.seek(0)
            st.download_button(
                label="📂 Descargar ZIP por bases, cargos y repetidos",
                data=zip_bases_buffer,
                file_name="Certificados_Por_Base_y_Cargo.zip",
                mime="application/zip"
            )
        else:
            st.warning("⚠️ No hay certificados para agrupar por bases y cargos")

        # Mostrar errores
        if errores > 0:
            with st.expander(f"⚠️ Ver {errores} error(es)"):
                df_errores = df_log[df_log['Estado'].str.startswith('ERROR')]
                st.dataframe(df_errores, use_container_width=True)
