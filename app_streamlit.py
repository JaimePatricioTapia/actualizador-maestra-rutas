#!/usr/bin/env python3
"""
Actualizador Maestra de Rutas - Streamlit App
==============================================
Aplicación web para actualizar la Maestra de Rutas con datos del Compilado.
"""

import streamlit as st
import pandas as pd
import tempfile
import os
from pathlib import Path
from datetime import datetime

# Importar módulos locales
from actualizador_maestra_rutas import (
    cargar_datos, matching_exacto, matching_relativo,
    aplicar_cambios, guardar_maestra_actualizada, generar_reporte, calcular_kpis
)
from generador_pdf import generar_pdf_comparacion

# Configuración de la página
st.set_page_config(
    page_title="Actualizador Maestra de Rutas",
    page_icon="📊",
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
</style>
""", unsafe_allow_html=True)

# Header con estilo Castaño
st.markdown('<div class="header-bar">🥐 Sistema de Planificación de Rutas - Castaño</div>', unsafe_allow_html=True)

# Título principal
st.markdown('<h1 class="main-title">Actualizador Maestra de Rutas</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Sistema de sincronización y comparación de planillas</p>', unsafe_allow_html=True)

# Inicializar session state
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'resultados' not in st.session_state:
    st.session_state.resultados = None

# Sección de carga de archivos
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

# Botón de procesamiento
st.markdown("---")

if maestra_file and compilado_file:
    if st.button("🚀 Procesar Archivos", type="primary", use_container_width=True):
        with st.spinner("Procesando archivos..."):
            try:
                # Crear directorio temporal para archivos
                with tempfile.TemporaryDirectory() as temp_dir:
                    temp_path = Path(temp_dir)
                    
                    # Guardar archivos temporalmente
                    maestra_path = temp_path / "maestra_temp.xlsx"
                    compilado_path = temp_path / "compilado_temp.xlsx"
                    
                    with open(maestra_path, 'wb') as f:
                        f.write(maestra_file.getvalue())
                    with open(compilado_path, 'wb') as f:
                        f.write(compilado_file.getvalue())
                    
                    # Cargar datos (ambos archivos)
                    df_maestra, df_compilado = cargar_datos(str(maestra_path), str(compilado_path))
                    
                    # Guardar copia original para comparación PDF
                    df_maestra_original = df_maestra.copy()
                    
                    # Matching exacto
                    coincidencias_exactas, sin_coincidencia = matching_exacto(df_maestra, df_compilado)
                    
                    # Matching relativo
                    coincidencias_relativas, ambiguos, sin_match = matching_relativo(df_maestra, sin_coincidencia)
                    
                    # Combinar coincidencias
                    todas_coincidencias = coincidencias_exactas + coincidencias_relativas
                    
                    # Aplicar cambios
                    df_actualizado, log_cambios = aplicar_cambios(df_maestra, todas_coincidencias)
                    
                    # Calcular KPIs
                    kpis = calcular_kpis(
                        len(df_compilado),
                        len(coincidencias_exactas),
                        len(coincidencias_relativas),
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
                        ambiguos,
                        sin_match,
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
                        coincidencias_relativas=coincidencias_relativas,
                        sin_match=sin_match
                    )
                    
                    # Leer archivos para descarga
                    with open(maestra_out_path, 'rb') as f:
                        maestra_bytes = f.read()
                    with open(reporte_path, 'rb') as f:
                        reporte_bytes = f.read()
                    with open(pdf_path, 'rb') as f:
                        pdf_bytes = f.read()
                    
                    # Guardar resultados en session state
                    st.session_state.processed = True
                    st.session_state.resultados = {
                        'kpis': kpis,
                        'match_exacto': len(coincidencias_exactas),
                        'match_relativo': len(coincidencias_relativas),
                        'total_cambios': len(log_cambios),
                        'total_compilado': len(df_compilado),
                        'maestra_bytes': maestra_bytes,
                        'reporte_bytes': reporte_bytes,
                        'pdf_bytes': pdf_bytes,
                        'maestra_filename': f"Maestra_ACTUALIZADA_{timestamp}.xlsx",
                        'reporte_filename': f"Reporte_{timestamp}.xlsx",
                        'pdf_filename': f"Comparacion_{timestamp}.pdf"
                    }
                    
                    st.success("✅ Procesamiento completado exitosamente")
                    st.rerun()
                    
            except Exception as e:
                st.error(f"❌ Error durante el procesamiento: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
else:
    st.info("📤 Carga ambos archivos para habilitar el procesamiento")

# Mostrar resultados si hay datos procesados
if st.session_state.processed and st.session_state.resultados:
    res = st.session_state.resultados
    
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
        st.session_state.processed = False
        st.session_state.resultados = None
        st.rerun()

# Footer
st.markdown('<p class="footer">🥐 Castaño - Sistema de Planificación de Rutas | Tienda Perfecta 2026</p>', unsafe_allow_html=True)
