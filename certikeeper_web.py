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

# Estilos CSS personalizados
st.markdown("""
<style>
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    h1 {
        background: linear-gradient(120deg, #1f77b4, #667eea);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        margin: 10px 0;
    }
    .stDownloadButton button {
        background: linear-gradient(120deg, #667eea, #764ba2);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 12px 24px;
        font-weight: 600;
        transition: all 0.3s;
    }
    .stDownloadButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(0,0,0,0.2);
    }
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
        color: white;
    }
    div[data-testid="stMetricValue"] {
        font-size: 28px;
        font-weight: 700;
        color: #1f77b4;
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
            if clave en linea:
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
st.markdown("<p style='text-align:center; color:#666; font-size:18px;'>Sistema inteligente de procesamiento y organización de certificados</p>", unsafe_allow_html=True)
st.markdown("<hr style='margin: 30px 0;'>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🎯 Panel de Control")
    st.markdown("---")
    
    st.markdown("#### 📍 Bases disponibles:")
    bases_col1, bases_col2 = st.columns(2)
    with bases_col1:
        st.markdown("🔹 ADZ  \n🔹 AXM  \n🔹 CLO  \n🔹 BAQ")
    with bases_col2:
        st.markdown("🔹 BGA  \n🔹 SMR  \n🔹 CTG  \n🔹 PEI")
    
    st.markdown("---")
    st.markdown("#### 👥 Tipos de cargo:")
    st.markdown("**OT** → Operaciones Terrestres  \n**SAP** → Servicio al Pasajero")
    
    st.markdown("---")
    st.info(f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    
    st.markdown("---")
    st.markdown("#### 💡 Consejos")
    st.markdown("• Puedes subir múltiples archivos  \n• Soporta PDF y ZIP  \n• Usa los filtros para análisis")

uploaded_files = st.file_uploader("📤 Sube tus PDFs o ZIP", accept_multiple_files=True, type=["pdf","zip"], help="Arrastra y suelta archivos aquí o haz clic para seleccionar")

if uploaded_files:
    with st.spinner("🔄 Procesando archivos..."):
        all_pdfs = extraer_pdfs_de_archivos(uploaded_files)
    
    if not all_pdfs:
        st.error("❌ No se encontraron PDFs.")
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
        
        progress.empty()
        status_text.empty()
        
        # Métricas destacadas
        st.markdown("### 📊 Resumen del Procesamiento")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        
        exitosos = len(all_pdfs) - errores
        with col_m1:
            st.metric("Total Procesados", len(all_pdfs), delta=None)
        with col_m2:
            st.metric("Exitosos", exitosos, delta=f"{(exitosos/len(all_pdfs)*100):.0f}%" if len(all_pdfs) > 0 else "0%")
        with col_m3:
            st.metric("Con Errores", errores, delta=f"-{(errores/len(all_pdfs)*100):.0f}%" if errores > 0 else "0%", delta_color="inverse")
        with col_m4:
            bases_unicas = len(set([x["Base"] for x in log if "Base" in x]))
            st.metric("Bases Detectadas", bases_unicas)
        
        st.markdown("---")
        
        # Panel de filtros
        st.markdown("### 🔍 Filtros y Análisis")
        
        df_log = pd.DataFrame(log)
        
        if not df_log.empty and "Base" in df_log.columns:
            col_f1, col_f2, col_f3, col_f4 = st.columns(4)
            
            with col_f1:
                bases_disponibles = ["Todas"] + sorted(df_log["Base"].dropna().unique().tolist())
                filtro_base = st.selectbox("📍 Filtrar por Base", bases_disponibles)
            
            with col_f2:
                if "Curso" in df_log.columns:
                    cursos_disponibles = ["Todos"] + sorted(df_log["Curso"].dropna().unique().tolist())
                    filtro_curso = st.selectbox("📚 Filtrar por Curso", cursos_disponibles)
                else:
                    filtro_curso = "Todos"
            
            with col_f3:
                if "Tipo" in df_log.columns:
                    tipos_disponibles = ["Todos"] + sorted(df_log["Tipo"].dropna().unique().tolist())
                    filtro_tipo = st.selectbox("👤 Filtrar por Tipo", tipos_disponibles)
                else:
                    filtro_tipo = "Todos"
            
            with col_f4:
                estados_disponibles = ["Todos"] + sorted(df_log["Estado"].dropna().unique().tolist())
                filtro_estado = st.selectbox("✅ Filtrar por Estado", estados_disponibles)
            
            # Aplicar filtros
            df_filtrado = df_log.copy()
            
            if filtro_base != "Todas":
                df_filtrado = df_filtrado[df_filtrado["Base"] == filtro_base]
            
            if filtro_curso != "Todos" and "Curso" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Curso"] == filtro_curso]
            
            if filtro_tipo != "Todos" and "Tipo" in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado["Tipo"] == filtro_tipo]
            
            if filtro_estado != "Todos":
                df_filtrado = df_filtrado[df_filtrado["Estado"] == filtro_estado]
            
            # Buscador de texto
            st.markdown("#### 🔎 Búsqueda en resultados")
            busqueda = st.text_input("Buscar por nombre, archivo o cualquier campo:", placeholder="Escribe para buscar...")
            
            if busqueda:
                mask = df_filtrado.astype(str).apply(lambda x: x.str.contains(busqueda, case=False, na=False)).any(axis=1)
                df_filtrado = df_filtrado[mask]
            
            st.markdown(f"**Mostrando {len(df_filtrado)} de {len(df_log)} registros**")
        else:
            df_filtrado = df_log
        
        # Tabla de resultados
        st.markdown("### 📋 Resultados Detallados")
        st.dataframe(
            df_filtrado, 
            use_container_width=True, 
            height=400,
            column_config={
                "Estado": st.column_config.TextColumn("Estado", width="small"),
                "Base": st.column_config.TextColumn("Base", width="small"),
            }
        )
        
        # Estadísticas adicionales
        if not df_filtrado.empty and "Base" in df_filtrado.columns:
            st.markdown("### 📈 Estadísticas")
            
            col_s1, col_s2 = st.columns(2)
            
            with col_s1:
                st.markdown("#### Por Base")
                if "Base" in df_filtrado.columns:
                    conteo_bases = df_filtrado["Base"].value_counts()
                    st.bar_chart(conteo_bases)
            
            with col_s2:
                st.markdown("#### Por Curso")
                if "Curso" in df_filtrado.columns:
                    conteo_cursos = df_filtrado["Curso"].value_counts()
                    st.bar_chart(conteo_cursos)
        
        st.markdown("---")
        
        # Botones de descarga
        st.markdown("### 💾 Descargar Resultados")
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
                df_log.to_excel(writer, index=False, sheet_name="Reporte")
            excel_buffer.seek(0)
            st.download_button(
                "📊 Descargar Reporte Excel", 
                excel_buffer, 
                file_name=f"reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

else:
    st.markdown("""
    <div style='text-align: center; padding: 50px; background: white; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
        <h3 style='color: #666;'>👆 Comienza subiendo tus archivos</h3>
        <p style='color: #999;'>Arrastra PDFs o archivos ZIP para procesarlos automáticamente</p>
    </div>
    """, unsafe_allow_html=True)
