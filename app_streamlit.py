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

# Estilos CSS personalizados
st.markdown("""
<style>
    /* Tema oscuro mejorado */
    .stApp {
        background: linear-gradient(135deg, #0E1117 0%, #1a1f2e 100%);
    }
    
    /* Título principal */
    .main-title {
        text-align: center;
        font-size: 2.5rem;
        font-weight: 700;
        background: linear-gradient(90deg, #00D9FF, #00FF88);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
    }
    
    .subtitle {
        text-align: center;
        color: #8892A0;
        font-size: 1rem;
        margin-bottom: 2rem;
    }
    
    /* Cards de KPIs */
    .kpi-container {
        display: flex;
        justify-content: center;
        gap: 1.5rem;
        flex-wrap: wrap;
        margin: 2rem 0;
    }
    
    .kpi-card {
        background: linear-gradient(145deg, #1E2329, #2A3140);
        border-radius: 16px;
        padding: 1.5rem 2rem;
        text-align: center;
        border: 1px solid #3A4150;
        min-width: 180px;
    }
    
    .kpi-value {
        font-size: 2.5rem;
        font-weight: 700;
        color: #00D9FF;
    }
    
    .kpi-label {
        font-size: 0.9rem;
        color: #8892A0;
        margin-top: 0.5rem;
    }
    
    /* Botones de descarga */
    .download-section {
        background: linear-gradient(145deg, #1E2329, #2A3140);
        border-radius: 16px;
        padding: 1.5rem;
        margin-top: 2rem;
        border: 1px solid #3A4150;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        color: #5A6270;
        font-size: 0.8rem;
        margin-top: 3rem;
        padding: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Título principal
st.markdown('<h1 class="main-title">📊 Actualizador Maestra de Rutas</h1>', unsafe_allow_html=True)
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
                        "Reporte de Comparación Maestra vs Compilado"
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
st.markdown('<p class="footer">Sistema de Actualización de Maestra de Rutas v1.0 | 2026</p>', unsafe_allow_html=True)
