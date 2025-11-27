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

palabras_invalidas = {"CARGO", "NEO", "AERO", "AGENTE", "SUPERVISOR", "COORDINADOR", "OPERADOR", "A", "PDE", "NEL", "EEE"}

# =========================
# OCR PDF
# =========================
def obtener_texto_con_ocr(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = ""
    for page in doc:
        texto += page.get_text()
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
# DETECCIÓN DE CARGO (OT / SAP)
# =========================
def detectar_tipo(texto):
    texto = texto.upper()
    
    claves_ot = [
        "OT", "OPERACIONES TERRESTRES",
        "AGENTE DE RAMPA", "OPERADOR DE RAMPA",
        "OPERARIO", "OPERACIÓN TERRESTRE"
    ]
    
    claves_sap = [
        "SAP", "PAX", "PASAJEROS",
        "SERVICIO AL PASAJERO", "ATENCIÓN A PASAJEROS",
        "CHECK IN", "PASAJERO"
    ]
    
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
# SOLO PRIMER NOMBRE + PRIMER APELLIDO
# =========================
def extraer_primer_nombre_apellido(nombre_completo):
    if not nombre_completo:
        return None, None

    limpio = nombre_completo.replace("\n", " ").replace("-", " ")
    limpio = " ".join(limpio.split())

    partes = limpio.split()
    if len(partes) < 2:
        return None, None

    primer_nombre = partes[0]
    apellido = partes[1]
    if len(partes) >= 3 and len(partes[1]) <= 3:
        apellido = partes[1] + partes[2]
    apellido = apellido.replace(" ", "")

    return primer_nombre, apellido

# =========================
# EXTRAER INFORMACIÓN (CON RAMPA CORRECTO)
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

    if curso_detectado:
        tipo = ""  # No se usa tipo en este curso
        curso_final = curso_detectado

        nombre_completo = detectar_nombre_con_flexibilidad(texto)
        if not nombre_completo:
            return None, None, None, None, None, "ERROR: Sin nombre"

        primer_nombre, primer_apellido = extraer_primer_nombre_apellido(nombre_completo)
        if not primer_nombre or not primer_apellido:
            return None, None, None, None, None, "ERROR: Nombre inválido"

        base_ab = base_abrev.get(base, "XXX")
        nuevo_nombre = f"{base_ab} {curso_final} {primer_nombre} {primer_apellido}".upper() + ".pdf"

        return base_ab, curso_final, tipo, f"{primer_nombre} {primer_apellido}", nuevo_nombre, "✅"

    # Cursos normales
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
# STREAMLIT UI
# =========================
st.markdown("""
    <h1 style='text-align: center; color: #1f77b4;'>
        📜 RENOMBRADOR DE CERTIFICADOS
    </h1>
    <p style='text-align: center; color: #666; font-size: 1.1em;'>
        Sistema automático de procesamiento y renombrado de certificados
    </p>
    <hr style='margin: 20px 0;'>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("ℹ️ Información")
    st.markdown("""
    ### 📋 Bases disponibles:
    - ADZ (San Andrés)
    - AXM (Armenia)
    - CLO (Cali)
    - BAQ (Barranquilla)
    - BGA (Bucaramanga)
    - SMR (Santa Marta)
    - CTG (Cartagena)
    - PEI (Pereira)
    
    ### 🎓 Tipos de cargo:
    - **OT**: Operaciones Terrestres
    - **SAP**: Servicio al Pasajero
    
    ### ✨ Características:
    - Procesa PDFs individuales
    - Procesa ZIP
    - Separa páginas
    - Detecta nombre/base/curso
    - Genera Excel
    """)
    st.markdown("---")
    st.info(f"**Fecha:** {datetime.now().strftime('%d/%m/%Y')}")

col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("### 📂 Carga de archivos")
    st.write("Sube uno o varios archivos PDF o ZIP.")

with col2:
    st.markdown("### 📊 Estadísticas")
    stat_placeholder = st.empty()

uploaded_files = st.file_uploader(
    "Arrastra o selecciona tus archivos",
    accept_multiple_files=True,
    type=["pdf", "zip"]
)

if uploaded_files:
    with st.spinner(f"🔄 Procesando {len(uploaded_files)} archivo(s)..."):
        all_pdfs = extraer_pdfs_de_archivos(uploaded_files)
    
    if not all_pdfs:
        st.error("❌ No se encontraron páginas PDF.")
    else:
        st.success(f"✅ Se extrajeron {len(all_pdfs)} páginas.")
        
        log = []
        renombrados = []
        errores = 0
        
        progress = st.progress(0)
        status_text = st.empty()
        
        for i, (nombre_original, pdf_bytes) in enumerate(all_pdfs):
            progress.progress((i+1)/len(all_pdfs))
            status_text.text(f"Procesando página {i+1}/{len(all_pdfs)}")
            
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
        
        status_text.empty()
        progress.empty()
        
        st.markdown("---")
        st.markdown("### 📈 Resumen del Procesamiento")

        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)

        with col_stat1:
            st.metric("Total Procesados", len(all_pdfs))
        with col_stat2:
            st.metric("Exitosos", len(renombrados))
        with col_stat3:
            st.metric("Errores", errores)
        with col_stat4:
            st.metric("Estado", "Sin errores" if errores == 0 else f"{errores} error(es)")

        st.markdown("---")
        st.markdown("### 📊 Resultados Detallados")

        df_log = pd.DataFrame(log)

        st.dataframe(df_log, use_container_width=True, height=400)

        st.markdown("---")
        st.markdown("### 📥 Descargas")

        col_d1, col_d2 = st.columns(2)

        with col_d1:
            if renombrados:
                zip_buffer = BytesIO()
                with ZipFile(zip_buffer, "w") as zipf:
                    for nombre, contenido in renombrados:
                        zipf.writestr(nombre, contenido)
                zip_buffer.seek(0)

                st.download_button(
                    "📦 Descargar ZIP",
                    zip_buffer,
                    file_name=f"certificados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )

        with col_d2:
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_log.to_excel(writer, index=False, sheet_name='Reporte')
            excel_buffer.seek(0)

            st.download_button(
                "📊 Descargar Excel",
                excel_buffer,
                file_name=f"reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

else:
    st.markdown("""
        <div style='text-align: center; padding: 40px; background-color: #f5f5f5; border-radius: 10px; margin-top: 20px;'>
            <h3 style='color: #666;'>👆 Por favor, carga uno o más archivos para iniciar el procesamiento.</h3>
            <p style='color: #777;'>Acepta PDF individuales o ZIP con múltiples certificados.</p>
        </div>
    """, unsafe_allow_html=True)
