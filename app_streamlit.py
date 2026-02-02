#!/usr/bin/env python3
"""
Aplicación Streamlit para el Actualizador de Maestra de Rutas
============================================================
Interfaz de usuario moderna con confirmación de matches relativos.
"""

import streamlit as st
import pandas as pd
import tempfile
from pathlib import Path
from datetime import datetime

# Importar funciones del actualizador
from actualizador_maestra_rutas import (
    cargar_datos,
    matching_exacto,
    matching_relativo,
    aplicar_cambios,
    calcular_kpis,
    generar_reporte,
    guardar_maestra_actualizada
)
from generador_pdf import generar_pdf_comparacion

# Configuración de la página
st.set_page_config(
    page_title="Actualizador Maestra de Rutas - Castaño",
    page_icon="🥐",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS personalizados - Inspirado en Castaño
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap');
    
    /* Variables de color Castaño */
    :root {
        --castano-gold: #FDBA24;
        --castano-brown: #3D0C11;
        --castano-red: #D71934;
        --castano-beige: #F2ECE1;
        --castano-white: #FFFFFF;
        --castano-text: #906F6F;
    }
    
    /* Fondo general */
    .stApp {
        background: var(--castano-beige) !important;
        font-family: 'Poppins', sans-serif !important;
    }
    
    /* Header con logo simulado */
    .header-bar {
        background: var(--castano-gold);
        color: var(--castano-brown);
        padding: 0.5rem;
        text-align: center;
        font-weight: 600;
        font-size: 0.9rem;
        border-radius: 0 0 16px 16px;
        margin-bottom: 1rem;
    }
    
    /* Título principal */
    .main-title {
        text-align: center;
        font-size: 2.2rem;
        font-weight: 700;
        color: var(--castano-brown) !important;
        font-family: 'Poppins', sans-serif !important;
        margin-bottom: 0.3rem;
    }
    
    .subtitle {
        text-align: center;
        color: var(--castano-text);
        font-size: 1rem;
        margin-bottom: 2rem;
        font-family: 'Poppins', sans-serif !important;
    }
    
    /* Cards y contenedores */
    .stFileUploader > div {
        background: var(--castano-white) !important;
        border-radius: 16px !important;
        border: 2px solid var(--castano-gold) !important;
    }
    
    /* Botones primarios */
    .stButton > button {
        background: var(--castano-red) !important;
        color: white !important;
        border: none !important;
        border-radius: 24px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 600 !important;
        font-family: 'Poppins', sans-serif !important;
        transition: all 0.3s ease !important;
    }
    
    .stButton > button:hover {
        background: var(--castano-brown) !important;
        transform: translateY(-2px) !important;
    }
    
    /* Métricas/KPIs */
    .stMetric {
        background: var(--castano-white) !important;
        border-radius: 16px !important;
        padding: 1rem !important;
        border: 1px solid rgba(61, 12, 17, 0.1) !important;
    }
    
    .stMetric label {
        color: var(--castano-text) !important;
        font-family: 'Poppins', sans-serif !important;
    }
    
    .stMetric [data-testid="stMetricValue"] {
        color: var(--castano-brown) !important;
        font-family: 'Poppins', sans-serif !important;
        font-weight: 700 !important;
    }
    
    /* Botones de descarga */
    .stDownloadButton > button {
        background: var(--castano-gold) !important;
        color: var(--castano-brown) !important;
        border: none !important;
        border-radius: 24px !important;
        font-weight: 600 !important;
        font-family: 'Poppins', sans-serif !important;
    }
    
    .stDownloadButton > button:hover {
        background: var(--castano-brown) !important;
        color: white !important;
    }
    
    /* Mensajes de info */
    .stAlert {
        border-radius: 16px !important;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        color: var(--castano-text);
        font-size: 0.8rem;
        margin-top: 3rem;
        padding: 1rem;
        font-family: 'Poppins', sans-serif !important;
    }
    
    /* Divisores */
    hr {
        border-color: var(--castano-gold) !important;
        opacity: 0.3 !important;
    }
    
    /* Tabla de confirmación */
    .confirmation-table {
        background: var(--castano-white);
        border-radius: 8px;
        padding: 0.5rem;
        margin: 0.3rem 0;
        border-left: 4px solid var(--castano-gold);
    }
</style>
""", unsafe_allow_html=True)

# Header con estilo Castaño
st.markdown('<div class="header-bar">🥐 Sistema de Planificación de Rutas - Castaño</div>', unsafe_allow_html=True)

# Título principal
st.markdown('<h1 class="main-title">Actualizador Maestra de Rutas</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Sistema de sincronización y comparación de planillas</p>', unsafe_allow_html=True)

# Inicializar session state
if 'step' not in st.session_state:
    st.session_state.step = 'upload'  # upload -> confirm -> results
if 'matches_pendientes' not in st.session_state:
    st.session_state.matches_pendientes = None
if 'matches_confirmados' not in st.session_state:
    st.session_state.matches_confirmados = []
if 'datos_temp' not in st.session_state:
    st.session_state.datos_temp = None
if 'resultados' not in st.session_state:
    st.session_state.resultados = None

# ==================== PASO 1: CARGA DE ARCHIVOS ====================
if st.session_state.step == 'upload':
    st.markdown("### 📁 Cargar Archivos")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Maestra de Rutas**")
        maestra_file = st.file_uploader(
            "Arrastra o selecciona el archivo",
            type=['xlsx', 'xls'],
            key="maestra",
            help="Archivo Excel con la Maestra de Rutas original"
        )
    
    with col2:
        st.markdown("**Archivo Compilado**")
        compilado_file = st.file_uploader(
            "Arrastra o selecciona el archivo",
            type=['xlsx', 'xls'],
            key="compilado",
            help="Archivo Excel con los datos del Compilado"
        )
    
    st.markdown("---")
    
    if maestra_file and compilado_file:
        if st.button("🚀 Analizar Coincidencias", type="primary", use_container_width=True):
            with st.spinner("Analizando archivos..."):
                try:
                    # Crear directorio temporal para archivos
                    temp_dir = tempfile.mkdtemp()
                    temp_path = Path(temp_dir)
                    
                    # Guardar archivos temporalmente
                    maestra_path = temp_path / "maestra_temp.xlsx"
                    compilado_path = temp_path / "compilado_temp.xlsx"
                    
                    with open(maestra_path, 'wb') as f:
                        f.write(maestra_file.getvalue())
                    with open(compilado_path, 'wb') as f:
                        f.write(compilado_file.getvalue())
                    
                    # Cargar datos
                    df_maestra, df_compilado = cargar_datos(str(maestra_path), str(compilado_path))
                    
                    # Guardar copia original para comparación PDF
                    df_maestra_original = df_maestra.copy()
                    
                    # Matching exacto
                    coincidencias_exactas, sin_coincidencia = matching_exacto(df_maestra, df_compilado)
                    
                    # Matching relativo
                    coincidencias_relativas, ambiguos, sin_match = matching_relativo(df_maestra, sin_coincidencia)
                    
                    # Guardar datos en session state
                    st.session_state.datos_temp = {
                        'temp_path': temp_path,
                        'df_maestra': df_maestra,
                        'df_maestra_original': df_maestra_original,
                        'df_compilado': df_compilado,
                        'coincidencias_exactas': coincidencias_exactas,
                        'coincidencias_relativas': coincidencias_relativas,
                        'ambiguos': ambiguos,
                        'sin_match': sin_match
                    }
                    
                    # Si hay matches relativos, ir a paso de confirmación
                    if coincidencias_relativas:
                        st.session_state.matches_pendientes = coincidencias_relativas
                        st.session_state.matches_confirmados = [True] * len(coincidencias_relativas)
                        st.session_state.step = 'confirm'
                    else:
                        # Si no hay matches relativos, ir directo a generar
                        st.session_state.step = 'generate'
                    
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ Error durante el análisis: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
    else:
        st.info("📤 Carga ambos archivos para habilitar el procesamiento")

# ==================== PASO 2: CONFIRMACIÓN DE MATCHES RELATIVOS ====================
elif st.session_state.step == 'confirm':
    datos = st.session_state.datos_temp
    matches = st.session_state.matches_pendientes
    
    st.markdown("### 🔍 Confirmar Coincidencias Relativas")
    st.markdown(f"Se encontraron **{len(matches)}** coincidencias relativas que requieren tu confirmación.")
    st.markdown("Estas coincidencias fueron encontradas porque el **código de centro** coincide, pero el formato difiere.")
    
    st.markdown("---")
    
    # Mostrar resumen de exactas
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Match Exacto", len(datos['coincidencias_exactas']))
    with col2:
        st.metric("Match Relativo (pendiente)", len(matches))
    with col3:
        st.metric("Sin Coincidencia", len(datos['sin_match']))
    
    st.markdown("---")
    st.markdown("#### Selecciona las coincidencias que deseas aplicar:")
    
    # Crear lista de checkboxes para cada match
    df_maestra = datos['df_maestra']
    
    for i, match in enumerate(matches):
        comp_row = match['compilado_row']
        maestra_row = df_maestra.loc[match['maestra_idx']].to_dict()
        
        # Container para cada match
        with st.container():
            # Checkbox en línea con el título
            col_check, col_info = st.columns([1, 11])
            
            with col_check:
                st.session_state.matches_confirmados[i] = st.checkbox(
                    "",  # Sin label, solo el checkbox
                    value=st.session_state.matches_confirmados[i],
                    key=f"confirm_{i}",
                    label_visibility="collapsed"
                )
            
            with col_info:
                tipo = match.get('tipo_match', 'RELATIVO')
                confianza = match.get('confianza', 0.7) * 100
                st.markdown(f"**Centro {match['center_code']}** - {comp_row.get('center_desc', 'N/A')} | `{tipo}` ({confianza:.0f}%)")
            
            # Expander con detalles (opcional)
            with st.expander("Ver detalles", expanded=False):
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**📋 Compilado (datos nuevos):**")
                    st.markdown(f"- Región: `{comp_row.get('region_desc', 'N/A')}`")
                    st.markdown(f"- Cliente: `{comp_row.get('customer_desc', 'N/A')}`")
                    st.markdown(f"- Formato: `{comp_row.get('formato', 'N/A')}`")
                    st.markdown(f"- Centro: `{comp_row.get('center_desc', 'N/A')}`")
                    st.markdown(f"- Usuario: `{comp_row.get('usuario', 'N/A')}`")
                
                with col2:
                    st.markdown("**📊 Maestra (datos actuales):**")
                    st.markdown(f"- Región: `{maestra_row.get('region_desc', 'N/A')}`")
                    st.markdown(f"- Cliente: `{maestra_row.get('customer_desc', 'N/A')}`")
                    st.markdown(f"- Formato: `{maestra_row.get('formato', 'N/A')}`")
                    st.markdown(f"- Centro: `{maestra_row.get('center_desc', 'N/A')}`")
                    st.markdown(f"- Usuario: `{maestra_row.get('usuario', 'N/A')}`")
    
    st.markdown("---")
    
    # Botones de acción
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("⬅️ Volver", use_container_width=True):
            st.session_state.step = 'upload'
            st.session_state.matches_pendientes = None
            st.session_state.datos_temp = None
            st.rerun()
    
    with col2:
        confirmados_count = sum(st.session_state.matches_confirmados)
        st.markdown(f"**{confirmados_count} de {len(matches)}** seleccionados")
    
    with col3:
        if st.button("✅ Confirmar y Generar Archivos", type="primary", use_container_width=True):
            st.session_state.step = 'generate'
            st.rerun()

# ==================== PASO 3: GENERAR ARCHIVOS ====================
elif st.session_state.step == 'generate':
    with st.spinner("Generando archivos finales..."):
        try:
            datos = st.session_state.datos_temp
            temp_path = datos['temp_path']
            df_maestra = datos['df_maestra']
            df_maestra_original = datos['df_maestra_original']
            coincidencias_exactas = datos['coincidencias_exactas']
            
            # Filtrar matches relativos confirmados
            matches_relativos_confirmados = []
            if st.session_state.matches_pendientes:
                for i, match in enumerate(st.session_state.matches_pendientes):
                    if st.session_state.matches_confirmados[i]:
                        matches_relativos_confirmados.append(match)
            
            # Combinar coincidencias
            todas_coincidencias = coincidencias_exactas + matches_relativos_confirmados
            
            # Aplicar cambios
            df_actualizado, log_cambios = aplicar_cambios(df_maestra, todas_coincidencias)
            
            # Calcular KPIs
            kpis = calcular_kpis(
                len(datos['df_compilado']),
                len(coincidencias_exactas),
                len(matches_relativos_confirmados),
                len(todas_coincidencias)
            )
            
            # Generar timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Guardar maestra actualizada
            maestra_out_path = temp_path / f"Maestra_ACTUALIZADA_{timestamp}.xlsx"
            guardar_maestra_actualizada(df_actualizado, str(maestra_out_path))
            
            # Generar reporte Excel
            reporte_path = temp_path / f"Reporte_{timestamp}.xlsx"
            generar_reporte(
                kpis,
                log_cambios,
                datos['ambiguos'],
                datos['sin_match'],
                str(reporte_path)
            )
            
            # Generar PDF
            pdf_path = temp_path / f"Comparacion_{timestamp}.pdf"
            generar_pdf_comparacion(
                todas_coincidencias,
                df_maestra_original,
                str(pdf_path),
                "Reporte de Comparación Maestra vs Compilado",
                coincidencias_exactas=coincidencias_exactas,
                coincidencias_relativas=matches_relativos_confirmados,
                sin_match=datos['sin_match']
            )
            
            # Leer archivos para descarga
            with open(maestra_out_path, 'rb') as f:
                maestra_bytes = f.read()
            with open(reporte_path, 'rb') as f:
                reporte_bytes = f.read()
            with open(pdf_path, 'rb') as f:
                pdf_bytes = f.read()
            
            # Guardar resultados
            st.session_state.resultados = {
                'kpis': kpis,
                'match_exacto': len(coincidencias_exactas),
                'match_relativo': len(matches_relativos_confirmados),
                'total_cambios': len(log_cambios),
                'total_compilado': len(datos['df_compilado']),
                'maestra_bytes': maestra_bytes,
                'reporte_bytes': reporte_bytes,
                'pdf_bytes': pdf_bytes,
                'maestra_filename': f"Maestra_ACTUALIZADA_{timestamp}.xlsx",
                'reporte_filename': f"Reporte_{timestamp}.xlsx",
                'pdf_filename': f"Comparacion_{timestamp}.pdf"
            }
            
            st.session_state.step = 'results'
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ Error durante la generación: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
            if st.button("⬅️ Volver"):
                st.session_state.step = 'upload'
                st.rerun()

# ==================== PASO 4: MOSTRAR RESULTADOS ====================
elif st.session_state.step == 'results':
    res = st.session_state.resultados
    
    st.success("✅ Procesamiento completado exitosamente")
    
    st.markdown("---")
    st.markdown("### 📈 Resultados")
    
    # KPIs en columnas
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        match_exacto_pct = (res['match_exacto'] / res['total_compilado'] * 100) if res['total_compilado'] > 0 else 0
        st.metric(
            label="Match Exacto",
            value=f"{match_exacto_pct:.1f}%",
            help="Porcentaje de filas con coincidencia exacta"
        )
    
    with col2:
        match_relativo_pct = (res['match_relativo'] / res['total_compilado'] * 100) if res['total_compilado'] > 0 else 0
        st.metric(
            label="Match Relativo",
            value=f"{match_relativo_pct:.1f}%",
            help="Porcentaje de filas con coincidencia heurística"
        )
    
    with col3:
        filas_actualizadas = res['match_exacto'] + res['match_relativo']
        filas_pct = (filas_actualizadas / res['total_compilado'] * 100) if res['total_compilado'] > 0 else 0
        st.metric(
            label="Filas Actualizadas",
            value=f"{filas_pct:.1f}%",
            help="Porcentaje de filas del Compilado procesadas"
        )
    
    with col4:
        st.metric(
            label="Total Cambios",
            value=res['total_cambios'],
            help="Número total de campos modificados"
        )
    
    # Sección de descargas
    st.markdown("---")
    st.markdown("### 📥 Descargar Archivos")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.download_button(
            label="📊 Maestra Actualizada",
            data=res['maestra_bytes'],
            file_name=res['maestra_filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.caption("Excel con cambios aplicados")
    
    with col2:
        st.download_button(
            label="📈 Reporte KPIs",
            data=res['reporte_bytes'],
            file_name=res['reporte_filename'],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.caption("Detalle de cambios")
    
    with col3:
        st.download_button(
            label="📄 Comparación PDF",
            data=res['pdf_bytes'],
            file_name=res['pdf_filename'],
            mime="application/pdf",
            use_container_width=True
        )
        st.caption("Visual de diferencias")
    
    # Botón para nuevo procesamiento
    st.markdown("---")
    if st.button("🔄 Nuevo Procesamiento", use_container_width=True):
        st.session_state.step = 'upload'
        st.session_state.matches_pendientes = None
        st.session_state.matches_confirmados = []
        st.session_state.datos_temp = None
        st.session_state.resultados = None
        st.rerun()

# Footer
st.markdown('<p class="footer">🥐 Castaño - Sistema de Planificación de Rutas | Tienda Perfecta 2026</p>', unsafe_allow_html=True)
