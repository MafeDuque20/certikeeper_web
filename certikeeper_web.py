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
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================
# ESTILOS CSS PERSONALIZADOS
# =========================
st.markdown("""
    <style>
    /* Tema principal */
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
    }
    
    /* Tarjetas con glassmorphism */
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    div[data-testid="stFileUploader"] {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 15px;
        padding: 2rem;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.3);
    }
    
    /* Botones mejorados */
    .stDownloadButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        font-size: 1rem;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stDownloadButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.6);
    }
    
    /* Sidebar mejorado */
    section[data-testid="stSidebar"] {
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(10px);
        border-right: 1px solid rgba(255, 255, 255, 0.3);
    }
    
    /* Títulos */
    h1 {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-weight: 800;
        text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    /* DataFrame mejorado */
    .stDataFrame {
        background: white;
        border-radius: 10px;
        padding: 1rem;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
    }
    
    /* Métricas */
    div[data-testid="stMetricValue"] {
        font-size: 2rem;
        font-weight: 700;
        color: #667eea;
    }
    
    /* Info boxes */
    .stAlert {
        border-radius: 10px;
        border-left: 4px solid #667eea;
        background: rgba(255, 255, 255, 0.9);
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 10px;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background: rgba(255, 255, 255, 0.9);
        border-radius: 10px;
        padding: 5px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 600;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

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
# CREAR ZIP ORGANIZADO (con RAMPA PAX/OT e INSTRUCTORES)
# =========================
def crear_zip_organizado(renombrados_info):
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer,"w") as zipf:
        for info in renombrados_info:
            nuevo_nombre = info["Nombre final"]
            pdf_bytes = info["Contenido"]
            tipo = info["Cargo"].upper() if info["Cargo"] else ""
            base = info["Base"]

            # Carpetas especiales
            if "INSTRUCTOR" in tipo:
                carpeta_tipo = "INSTRUCTORES"
            elif "SEGURIDAD EN RAMPA PAX" in nuevo_nombre:
                carpeta_tipo = "PAX"
            elif "SEGURIDAD EN RAMPA OT" in nuevo_nombre:
                carpeta_tipo = "RAMPA"
            elif tipo == "OT":
                carpeta_tipo = "RAMPA"
            elif tipo == "SAP":
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
st.markdown("<h1 style='text-align:center;'>📜 RENOMBRADOR DE CERTIFICADOS</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;color:white;font-size:1.2rem;'>Sistema automatizado de procesamiento y organización de certificados</p>", unsafe_allow_html=True)
st.markdown("<hr style='border: 2px solid white; margin: 2rem 0;'>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🎯 Panel de Control")
    st.markdown("---")
    
    with st.expander("📍 Bases Disponibles", expanded=True):
        cols = st.columns(2)
        bases_list = list(base_abrev.items())
        for idx, (ciudad, codigo) in enumerate(bases_list):
            with cols[idx % 2]:
                st.markdown(f"**{codigo}** - {ciudad}")
    
    with st.expander("👥 Tipos de Cargo", expanded=True):
        st.markdown("""
        - **OT**: Operaciones Terrestres
        - **SAP**: Servicio al Pasajero
        - **INSTRUCTOR**: Personal docente
        """)
    
    with st.expander("📚 Cursos Disponibles", expanded=False):
        for curso in cursos_validos.values():
            st.markdown(f"• {curso}")
    
    st.markdown("---")
    st.info(f"📅 **Fecha:** {datetime.now().strftime('%d/%m/%Y')}\n\n⏰ **Hora:** {datetime.now().strftime('%H:%M:%S')}")
    
    st.markdown("---")
    st.markdown("### 💡 Ayuda")
    st.markdown("""
    **Formatos aceptados:**
    - PDF individual
    - Múltiples PDFs
    - Archivos ZIP
    
    **Proceso:**
    1. Sube tus archivos
    2. Espera el procesamiento
    3. Descarga resultados
    """)

# Zona de carga de archivos
st.markdown("<div style='background: rgba(255,255,255,0.95); padding: 2rem; border-radius: 15px; box-shadow: 0 8px 32px rgba(0,0,0,0.1);'>", unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "📤 Arrastra o selecciona tus archivos",
    accept_multiple_files=True,
    type=["pdf","zip"],
    help="Puedes subir archivos PDF individuales o comprimidos en ZIP"
)
st.markdown("</div>", unsafe_allow_html=True)

if uploaded_files:
    st.markdown("<br>", unsafe_allow_html=True)
    
    with st.spinner("🔄 Procesando archivos..."):
        all_pdfs = extraer_pdfs_de_archivos(uploaded_files)
    
    if not all_pdfs:
        st.error("❌ No se encontraron PDFs válidos en los archivos subidos.")
    else:
        log, renombrados_info, errores = [], [], 0
        
        # Barra de progreso mejorada
        progress_container = st.container()
        with progress_container:
            progress_bar = st.progress(0)
            status_text = st.empty()
        
        for i, (nombre_original, pdf_bytes) in enumerate(all_pdfs):
            progress_bar.progress((i+1)/len(all_pdfs))
            status_text.markdown(f"**Procesando:** {i+1}/{len(all_pdfs)} - `{nombre_original}`")
            
            base, curso, tipo, alumno, nuevo_nombre, estado = extraer_info(pdf_bytes)
            
            if estado.startswith("ERROR"):
                errores += 1
                log.append({"Página original": nombre_original,"Estado":estado})
                continue
            
            renombrados_info.append({"Nombre final":nuevo_nombre,"Contenido":pdf_bytes,"Cargo":tipo,"Base":base})
            log.append({"Página original":nombre_original,"Estado":estado,"Nombre final":nuevo_nombre,"Base":base,"Curso":curso,"Tipo":tipo,"Alumno":alumno})
        
        progress_bar.empty()
        status_text.empty()
        
        # Métricas con diseño mejorado
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("📄 Total Procesados", len(all_pdfs))
        with col2:
            st.metric("✅ Exitosos", len(renombrados_info))
        with col3:
            st.metric("❌ Errores", errores)
        with col4:
            tasa_exito = (len(renombrados_info)/len(all_pdfs)*100) if len(all_pdfs) > 0 else 0
            st.metric("📊 Tasa de Éxito", f"{tasa_exito:.1f}%")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Tabs para organizar información
        tab1, tab2, tab3 = st.tabs(["📋 Registro Completo", "📊 Estadísticas", "🔍 Filtros"])
        
        with tab1:
            st.markdown("### Detalle del Procesamiento")
            df_log = pd.DataFrame(log)
            st.dataframe(df_log, use_container_width=True, height=400)
        
        with tab2:
            if renombrados_info:
                st.markdown("### Distribución por Base")
                df_stats = pd.DataFrame(renombrados_info)
                base_counts = df_stats['Base'].value_counts()
                col_a, col_b = st.columns(2)
                with col_a:
                    st.bar_chart(base_counts)
                with col_b:
                    for base, count in base_counts.items():
                        st.metric(f"Base {base}", count)
        
        with tab3:
            st.markdown("### Filtrar Resultados")
            df_log = pd.DataFrame(log)
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                if 'Base' in df_log.columns:
                    bases_filter = st.multiselect("Filtrar por Base", df_log['Base'].unique())
                    if bases_filter:
                        df_log = df_log[df_log['Base'].isin(bases_filter)]
            
            with col_f2:
                if 'Curso' in df_log.columns:
                    cursos_filter = st.multiselect("Filtrar por Curso", df_log['Curso'].unique())
                    if cursos_filter:
                        df_log = df_log[df_log['Curso'].isin(cursos_filter)]
            
            st.dataframe(df_log, use_container_width=True, height=300)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Botones de descarga mejorados
        st.markdown("### 📥 Descargar Resultados")
        col1, col2 = st.columns(2)
        
        with col1:
            zip_buffer = crear_zip_organizado(renombrados_info)
            st.download_button(
                "📦 Descargar ZIP Organizado",
                zip_buffer,
                file_name=f"certificados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                use_container_width=True
            )
        
        with col2:
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                pd.DataFrame(log).to_excel(writer,index=False,sheet_name="Reporte")
            excel_buffer.seek(0)
            st.download_button(
                "📊 Descargar Reporte Excel",
                excel_buffer,
                file_name=f"reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

else:
    # Estado inicial con instrucciones
    st.markdown("""
    <div style='background: rgba(255,255,255,0.9); padding: 3rem; border-radius: 15px; text-align: center; margin-top: 2rem;'>
        <h2 style='color: #667eea;'>👋 ¡Bienvenido!</h2>
        <p style='font-size: 1.2rem; color: #666;'>
            Comienza subiendo tus archivos PDF o ZIP en el área superior
        </p>
        <p style='color: #888;'>
            El sistema procesará automáticamente todos los certificados y los organizará por base y tipo de cargo
        </p>
    </div>
    """, unsafe_allow_html=True)
