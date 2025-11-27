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
    "CARTAGENA": "CTG"
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
    
    # Palabras clave fuertes OT
    claves_ot = [
        "OT", "OPERACIONES TERRESTRES", "RAMPA",
        "AGENTE DE RAMPA", "OPERADOR DE RAMPA",
        "OPERARIO", "OPERACIÓN TERRESTRE"
    ]
    
    # Palabras clave fuertes SAP
    claves_sap = [
        "SAP", "PAX", "PASAJEROS",
        "SERVICIO AL PASAJERO", "ATENCIÓN A PASAJEROS",
        "CHECK IN", "PASAJERO"
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
    
    # Crear nombre del archivo SIEMPRE con el tipo
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

# Header con estilo
st.markdown("""
    <h1 style='text-align: center; color: #1f77b4;'>
        📜 RENOMBRADOR DE CERTIFICADOS
    </h1>
    <p style='text-align: center; color: #666; font-size: 1.1em;'>
        Sistema automático de procesamiento y renombrado de certificados
    </p>
    <hr style='margin: 20px 0;'>
""", unsafe_allow_html=True)

# Información útil en sidebar
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
    
    ### 🎓 Tipos de cargo:
    - **OT**: Operaciones Terrestres
    - **SAP**: Servicio al Pasajero
    
    ### ✨ Características:
    - ✅ Procesa PDFs individuales
    - ✅ Procesa archivos ZIP
    - ✅ Separa páginas automáticamente
    - ✅ Detecta nombre, base y curso
    - ✅ Genera reporte Excel
    """)
    
    st.markdown("---")
    st.info(f"**Fecha:** {datetime.now().strftime('%d/%m/%Y')}")

# Área principal
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("### 📂 Carga de archivos")
    st.write("Sube uno o varios archivos PDF, o un archivo ZIP que contenga PDFs.")

with col2:
    st.markdown("### 📊 Estadísticas")
    stat_placeholder = st.empty()

uploaded_files = st.file_uploader(
    "Arrastra o selecciona tus archivos",
    accept_multiple_files=True,
    type=["pdf", "zip"],
    help="Puedes subir múltiples PDFs o archivos ZIP"
)

if uploaded_files:
    with st.spinner(f"🔄 Procesando {len(uploaded_files)} archivo(s)..."):
        all_pdfs = extraer_pdfs_de_archivos(uploaded_files)
    
    if not all_pdfs:
        st.error("❌ No se encontraron páginas PDF en los archivos cargados.")
    else:
        st.success(f"✅ Se extrajeron **{len(all_pdfs)}** páginas correctamente.")
        
        log = []
        renombrados = []
        errores = 0
        
        progress = st.progress(0)
        status_text = st.empty()
        
        for i, (nombre_original, pdf_bytes) in enumerate(all_pdfs):
            progress.progress((i+1)/len(all_pdfs))
            status_text.text(f"Procesando: {i+1}/{len(all_pdfs)}")
            
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
        
        # ============================================
        # ESTADÍSTICAS MEJORADAS
        # ============================================
        st.markdown("---")
        st.markdown("### 📈 Resumen del Procesamiento")
        
        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
        
        with col_stat1:
            st.markdown(f"""
                <div style='text-align: center; padding: 20px; background-color: #e3f2fd; border-radius: 10px;'>
                    <h2 style='color: #1976d2; margin: 0;'>{len(all_pdfs)}</h2>
                    <p style='color: #424242; margin: 5px 0 0 0;'>Total Procesados</p>
                </div>
            """, unsafe_allow_html=True)
        
        with col_stat2:
            porcentaje_exito = (len(renombrados)/len(all_pdfs)*100) if len(all_pdfs) > 0 else 0
            st.markdown(f"""
                <div style='text-align: center; padding: 20px; background-color: #e8f5e9; border-radius: 10px;'>
                    <h2 style='color: #388e3c; margin: 0;'>{len(renombrados)}</h2>
                    <p style='color: #424242; margin: 5px 0 0 0;'>✅ Exitosos ({porcentaje_exito:.1f}%)</p>
                </div>
            """, unsafe_allow_html=True)
        
        with col_stat3:
            st.markdown(f"""
                <div style='text-align: center; padding: 20px; background-color: #ffebee; border-radius: 10px;'>
                    <h2 style='color: #d32f2f; margin: 0;'>{errores}</h2>
                    <p style='color: #424242; margin: 5px 0 0 0;'>❌ Errores</p>
                </div>
            """, unsafe_allow_html=True)
        
        with col_stat4:
            estado_general = "✅ SIN ERRORES" if errores == 0 else f"⚠️ {errores} ERROR(ES)"
            color_estado = "#388e3c" if errores == 0 else "#f57c00"
            bg_color = "#e8f5e9" if errores == 0 else "#fff3e0"
            st.markdown(f"""
                <div style='text-align: center; padding: 20px; background-color: {bg_color}; border-radius: 10px;'>
                    <h2 style='color: {color_estado}; margin: 0; font-size: 1.3em;'>{estado_general}</h2>
                    <p style='color: #424242; margin: 5px 0 0 0;'>Estado General</p>
                </div>
            """, unsafe_allow_html=True)
        
        # ============================================
        # TABLA DE RESULTADOS CON FILTROS
        # ============================================
        st.markdown("---")
        st.markdown("### 📊 Resultados Detallados")
        
        # Crear DataFrame
        df_log = pd.DataFrame(log)
        
        # Filtros
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            filtro_estado = st.selectbox("🔍 Filtrar por estado:", ["Todos", "✅", "ERROR"])
        with col_f2:
            bases_unicas = ["Todas"] + sorted([b for b in df_log["Base"].unique() if b])
            filtro_base = st.selectbox("🌍 Filtrar por base:", bases_unicas)
        with col_f3:
            tipos_unicos = ["Todos"] + sorted([t for t in df_log["Tipo"].unique() if t])
            filtro_tipo = st.selectbox("👥 Filtrar por tipo:", tipos_unicos)
        
        # Aplicar filtros
        df_filtrado = df_log.copy()
        if filtro_estado != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Estado"] == filtro_estado]
        if filtro_base != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Base"] == filtro_base]
        if filtro_tipo != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Tipo"] == filtro_tipo]
        
        st.dataframe(df_filtrado, use_container_width=True, height=400)
        
        # ============================================
        # BOTONES DE DESCARGA MEJORADOS
        # ============================================
        st.markdown("---")
        st.markdown("### 📥 Descargas Disponibles")
        
        col_d1, col_d2 = st.columns(2)
        
        with col_d1:
            st.markdown("#### 📦 Certificados Renombrados")
            st.write("Descarga todos los certificados procesados en un solo archivo ZIP.")
            
            if renombrados:
                zip_buffer = BytesIO()
                with ZipFile(zip_buffer, "w") as zipf:
                    for nombre, contenido in renombrados:
                        zipf.writestr(nombre, contenido)
                zip_buffer.seek(0)
                
                st.download_button(
                    label="📦 Descargar ZIP de Certificados",
                    data=zip_buffer,
                    file_name=f"certificados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            else:
                st.warning("⚠️ No hay certificados exitosos para descargar")
        
        with col_d2:
            st.markdown("#### 📊 Reporte en Excel")
            st.write("Descarga el reporte completo con todos los detalles del procesamiento.")
            
            # Crear Excel con formato mejorado
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_log.to_excel(writer, index=False, sheet_name='Reporte Completo')
                
                # Ajustar ancho de columnas
                worksheet = writer.sheets['Reporte Completo']
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            excel_buffer.seek(0)
            
            st.download_button(
                label="📊 Descargar Reporte Excel",
                data=excel_buffer,
                file_name=f"reporte_certificados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        # ============================================
        # HERRAMIENTA DE BÚSQUEDA
        # ============================================
        st.markdown("---")
        st.markdown("### 🔍 Buscar Certificado Específico")
        
        col_search1, col_search2 = st.columns([3, 1])
        
        with col_search1:
            busqueda = st.text_input(
                "Ingresa nombre o apellido del alumno:",
                placeholder="Ejemplo: Juan, Perez, etc.",
                help="La búsqueda no distingue mayúsculas y minúsculas"
            )
        
        with col_search2:
            st.write("")
            st.write("")
            buscar_btn = st.button("🔍 Buscar", use_container_width=True)
        
        if busqueda or buscar_btn:
            resultados = df_log[
                df_log["Alumno"].str.contains(busqueda.upper(), na=False) |
                df_log["Nombre final"].str.contains(busqueda.upper(), na=False)
            ]
            
            if len(resultados) > 0:
                st.success(f"✅ Se encontraron **{len(resultados)}** resultado(s)")
                st.dataframe(resultados, use_container_width=True)
            else:
                st.warning("⚠️ No se encontraron resultados para tu búsqueda")

else:
    # Mensaje cuando no hay archivos
    st.markdown("""
        <div style='text-align: center; padding: 40px; background-color: #f5f5f5; border-radius: 10px; margin-top: 20px;'>
            <h3 style='color: #666;'>👆 Por favor, carga uno o más archivos para comenzar</h3>
            <p style='color: #888;'>Formatos aceptados: PDF, ZIP</p>
        </div>
    """, unsafe_allow_html=True)
