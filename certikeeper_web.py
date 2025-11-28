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
    page_title="CertiKeeper Web",
    page_icon="🔐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================
# ESTILOS CSS PERSONALIZADOS
# =========================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    .main {
        background: #0a0a0a;
        padding: 2rem;
    }
    
    .stApp {
        background: #0a0a0a;
    }
    
    /* File Uploader */
    div[data-testid="stFileUploader"] {
        background: #1a1a1a;
        border-radius: 12px;
        padding: 2rem;
        border: 2px solid #2a2a2a;
        transition: all 0.3s ease;
    }
    
    div[data-testid="stFileUploader"]:hover {
        border-color: #3a3a3a;
        background: #1f1f1f;
    }
    
    /* Botones */
    .stDownloadButton button {
        background: linear-gradient(135deg, #1a1a1a 0%, #2d2d2d 100%);
        color: white;
        border: 1px solid #3a3a3a;
        border-radius: 8px;
        padding: 0.65rem 1.5rem;
        font-weight: 600;
        font-size: 0.95rem;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .stDownloadButton button:hover {
        background: linear-gradient(135deg, #2d2d2d 0%, #3d3d3d 100%);
        border-color: #4a4a4a;
        transform: translateY(-1px);
    }
    
    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #121212;
        border-right: 1px solid #2a2a2a;
    }
    
    section[data-testid="stSidebar"] > div {
        background: #121212;
    }
    
    /* Títulos */
    h1 {
        color: white;
        font-weight: 800;
        letter-spacing: -0.02em;
    }
    
    h2, h3 {
        color: #e0e0e0;
        font-weight: 700;
    }
    
    /* DataFrame */
    .stDataFrame {
        background: #1a1a1a;
        border-radius: 8px;
        border: 1px solid #2a2a2a;
    }
    
    /* Métricas */
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-weight: 700;
        color: white;
    }
    
    div[data-testid="stMetricLabel"] {
        color: #888;
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, #1a1a1a 0%, #3a3a3a 100%);
        border-radius: 4px;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: transparent;
        border-bottom: 1px solid #2a2a2a;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 6px 6px 0 0;
        padding: 10px 20px;
        font-weight: 600;
        color: #888;
        background: transparent;
        border: none;
    }
    
    .stTabs [aria-selected="true"] {
        background: #1a1a1a;
        color: white;
        border-bottom: 2px solid white;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background: #1a1a1a;
        border-radius: 8px;
        color: white;
        font-weight: 600;
    }
    
    .streamlit-expanderContent {
        background: #151515;
        border: 1px solid #2a2a2a;
        border-top: none;
    }
    
    /* Info boxes */
    .stAlert {
        background: #1a1a1a;
        border: 1px solid #2a2a2a;
        border-radius: 8px;
        color: #f0f0f0;
    }
    
    /* Text inputs */
    .stTextInput input {
        background: #1a1a1a;
        border: 1px solid #2a2a2a;
        border-radius: 6px;
        color: white;
    }
    
    .stTextInput input:focus {
        border-color: #3a3a3a;
        box-shadow: 0 0 0 1px #3a3a3a;
    }
    
    /* Multiselect */
    .stMultiSelect {
        background: #1a1a1a;
    }
    
    /* Spinner */
    .stSpinner > div {
        border-top-color: white !important;
    }
    
    /* Custom card */
    .metric-card {
        background: #1a1a1a;
        border: 1px solid #2a2a2a;
        border-radius: 8px;
        padding: 1.5rem;
        text-align: center;
    }
    
    p, li, span {
        color: #f5f5f5;
    }
    
    /* Mejorar legibilidad de textos */
    .stMarkdown, .stText {
        color: #ffffff;
    }
    
    label {
        color: #ffffff !important;
    }
    
    /* Multiselect text */
    div[data-baseweb="select"] span {
        color: #ffffff !important;
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
# FUNCIONES DE PROCESAMIENTO
# =========================
def obtener_texto_con_ocr(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texto = "".join([page.get_text() for page in doc])
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
    if not nombre_completo: return None, None
    limpio = " ".join(nombre_completo.replace("\n"," ").replace("-"," ").split())
    partes = limpio.split()
    if len(partes)<2: return None, None
    return partes[0], partes[1]

def extraer_info(pdf_bytes):
    texto = obtener_texto_con_ocr(pdf_bytes)
    base = detectar_base(texto)

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

    if curso_detectado:
        nuevo_nombre = f"{base_ab} {curso_detectado} {primer_nombre} {primer_apellido}".upper() + ".pdf"
        tipo = ""
        return base_ab, curso_detectado, tipo, f"{primer_nombre} {primer_apellido}", nuevo_nombre, "✅"
    
    curso = detectar_curso(texto)
    tipo = detectar_tipo(texto)
    nuevo_nombre = f"{base_ab} {curso} {tipo} {primer_nombre} {primer_apellido}".upper() + ".pdf"
    return base_ab, curso, tipo, f"{primer_nombre} {primer_apellido}", nuevo_nombre, "✅"

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
        st.warning(f"Error: {nombre_origen}")
    return paginas

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
                st.warning(f"Error ZIP: {uploaded.name}")
    return pdfs

def crear_zip_organizado(renombrados_info):
    zip_buffer = BytesIO()
    certificados_vistos = {}  # Dict para rastrear certificados únicos: clave -> índice en renombrados_info
    
    with ZipFile(zip_buffer,"w") as zipf:
        for idx, info in enumerate(renombrados_info):
            nuevo_nombre = info["Nombre final"]
            pdf_bytes = info["Contenido"]
            tipo = info["Cargo"].upper() if info["Cargo"] else ""
            base = info["Base"]
            alumno = info.get("Alumno", "")
            
            # Crear clave única para detectar duplicados (alumno + tipo + base)
            # Limpiar espacios extras y normalizar
            alumno_limpio = " ".join(alumno.strip().split()).upper()
            tipo_limpio = tipo.strip().upper()
            base_limpio = base.strip().upper()
            clave_unica = f"{alumno_limpio}|{tipo_limpio}|{base_limpio}"
            
            # Verificar si es duplicado
            es_duplicado = False
            if clave_unica in certificados_vistos:
                es_duplicado = True
                # Marcar el actual como duplicado
            else:
                # Guardar el índice del primer certificado con esta clave
                certificados_vistos[clave_unica] = idx

            # Determinar carpeta de destino
            if es_duplicado:
                carpeta_tipo = "Repetidos"
            elif "INSTRUCTOR" in tipo:
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
st.markdown("<h1 style='text-align:center; margin-bottom: 0.5rem;'>🔐 hB</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;color:#888;font-size:0.95rem; margin-bottom: 2rem;'>Sistema interno de gestión de certificados</p>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Sistema")
    st.markdown("---")
    
    with st.expander("📍 Bases", expanded=False):
        for ciudad, codigo in base_abrev.items():
            st.markdown(f"`{codigo}` {ciudad}")
    
    with st.expander("👥 Cargos", expanded=False):
        st.markdown("• **OT** - Operaciones\n• **SAP** - Pasajeros\n• **INSTRUCTOR** - Docente")
    
    st.markdown("---")
    st.markdown(f"**{datetime.now().strftime('%d/%m/%Y')}** · {datetime.now().strftime('%H:%M')}")

# Zona de carga
st.markdown("<div style='background: #1a1a1a; padding: 2rem; border-radius: 12px; border: 2px solid #2a2a2a;'>", unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "Cargar archivos",
    accept_multiple_files=True,
    type=["pdf","zip"],
    help="PDF o ZIP"
)
st.markdown("</div>", unsafe_allow_html=True)

if uploaded_files:
    st.markdown("<br>", unsafe_allow_html=True)
    
    with st.spinner("Procesando..."):
        all_pdfs = extraer_pdfs_de_archivos(uploaded_files)
    
    if not all_pdfs:
        st.error("No se encontraron PDFs válidos")
    else:
        log, renombrados_info, errores = [], [], 0
        
        progress_container = st.container()
        with progress_container:
            progress_bar = st.progress(0)
            status_text = st.empty()
        
        for i, (nombre_original, pdf_bytes) in enumerate(all_pdfs):
            progress_bar.progress((i+1)/len(all_pdfs))
            status_text.markdown(f"**{i+1}/{len(all_pdfs)}** `{nombre_original}`")
            
            base, curso, tipo, alumno, nuevo_nombre, estado = extraer_info(pdf_bytes)
            
            if estado.startswith("ERROR"):
                errores += 1
                log.append({"ID": len(log)+1, "Página original": nombre_original,"Estado":estado, "Nombre final": "", "Base": "", "Curso": "", "Tipo": "", "Alumno": ""})
                continue
            
            renombrados_info.append({"Nombre final":nuevo_nombre,"Contenido":pdf_bytes,"Cargo":tipo,"Base":base,"Alumno":alumno})
            log.append({"ID": len(log)+1, "Página original":nombre_original,"Estado":estado,"Nombre final":nuevo_nombre,"Base":base,"Curso":curso,"Tipo":tipo,"Alumno":alumno})
        
        progress_bar.empty()
        status_text.empty()
        
        # Métricas
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("TOTAL", len(all_pdfs))
        with col2:
            st.metric("EXITOSOS", len(renombrados_info))
        with col3:
            st.metric("ERRORES", errores)
        with col4:
            tasa_exito = (len(renombrados_info)/len(all_pdfs)*100) if len(all_pdfs) > 0 else 0
            st.metric("TASA ÉXITO", f"{tasa_exito:.1f}%")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Tabs
        tab1, tab2, tab3 = st.tabs(["📋 Registros", "📊 Estadísticas", "✏️ Editor"])
        
        with tab1:
            df_log = pd.DataFrame(log)
            st.dataframe(df_log, use_container_width=True, height=400, hide_index=True)
        
        with tab2:
            if renombrados_info:
                df_stats = pd.DataFrame(renombrados_info)
                base_counts = df_stats['Base'].value_counts()
                col_a, col_b = st.columns([2, 1])
                with col_a:
                    st.bar_chart(base_counts)
                with col_b:
                    for base, count in base_counts.items():
                        st.metric(f"{base}", count)
        
        with tab3:
            st.markdown("### Filtrar y Editar")
            df_edit = pd.DataFrame(log)
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                if 'Base' in df_edit.columns:
                    bases_filter = st.multiselect("Base", df_edit['Base'].unique())
                    if bases_filter:
                        df_edit = df_edit[df_edit['Base'].isin(bases_filter)]
            
            with col_f2:
                if 'Curso' in df_edit.columns:
                    cursos_filter = st.multiselect("Curso", df_edit['Curso'].unique())
                    if cursos_filter:
                        df_edit = df_edit[df_edit['Curso'].isin(cursos_filter)]
            
            st.markdown("#### Editar Nombres")
            st.info("💡 Haz clic en una celda de 'Nombre final' para editarla")
            
            # Editor de datos
            edited_df = st.data_editor(
                df_edit,
                use_container_width=True,
                height=350,
                hide_index=True,
                column_config={
                    "ID": st.column_config.NumberColumn("ID", disabled=True, width="small"),
                    "Página original": st.column_config.TextColumn("Original", disabled=True),
                    "Estado": st.column_config.TextColumn("Estado", disabled=True, width="small"),
                    "Nombre final": st.column_config.TextColumn("Nombre Final", width="large"),
                    "Base": st.column_config.TextColumn("Base", disabled=True, width="small"),
                    "Curso": st.column_config.TextColumn("Curso", disabled=True),
                    "Tipo": st.column_config.TextColumn("Tipo", disabled=True, width="small"),
                    "Alumno": st.column_config.TextColumn("Alumno", disabled=True)
                },
                disabled=["ID", "Página original", "Estado", "Base", "Curso", "Tipo", "Alumno"]
            )
            
            # Actualizar renombrados_info con los cambios
            if not edited_df.equals(df_edit):
                st.success("✅ Cambios detectados")
                
                # Actualizar los nombres en renombrados_info
                for idx, row in edited_df.iterrows():
                    if row["Estado"] == "✅":
                        # Buscar el índice correspondiente en renombrados_info
                        original_name = row["Nombre final"]
                        for i, info in enumerate(renombrados_info):
                            # Comparar por alumno y base para identificar el registro correcto
                            if (info["Base"] == row["Base"] and 
                                row["Alumno"] in info["Nombre final"]):
                                renombrados_info[i]["Nombre final"] = row["Nombre final"]
                                break
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Botones de descarga
        col1, col2 = st.columns(2)
        
        with col1:
            zip_buffer = crear_zip_organizado(renombrados_info)
            st.download_button(
                "📦 Descargar ZIP",
                zip_buffer,
                file_name=f"certificados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                use_container_width=True
            )
        
        with col2:
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                # Usar edited_df si existe, sino df_log
                final_df = edited_df if 'edited_df' in locals() else pd.DataFrame(log)
                final_df.to_excel(writer,index=False,sheet_name="Reporte")
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
    <div style='background: #1a1a1a; padding: 3rem; border-radius: 12px; text-align: center; margin-top: 2rem; border: 1px solid #2a2a2a;'>
        <h2 style='color: white; margin-bottom: 1rem;'>Comienza aquí</h2>
        <p style='font-size: 1rem; color: #888;'>
            Sube archivos PDF o ZIP para procesarlos automáticamente
        </p>
    </div>
    """, unsafe_allow_html=True)
