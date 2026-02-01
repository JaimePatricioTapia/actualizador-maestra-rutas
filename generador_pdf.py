#!/usr/bin/env python3
"""
Generador de PDF para Comparación de Maestra vs Compilado
==========================================================
Genera un PDF mostrando las filas comparadas con diferencias resaltadas.
"""

import pandas as pd
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Tuple

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak


# Colores para el reporte
COLOR_MODIFICADO = colors.HexColor('#C8E6C9')   # Verde claro para modificado
COLOR_DIFERENCIA = colors.HexColor('#FFCDD2')   # Rojo claro para diferencias
COLOR_MAESTRA = colors.HexColor('#E3F2FD')      # Azul muy claro para identificar Maestra
COLOR_COMPILADO = colors.HexColor('#FFF3E0')    # Naranja muy claro para Compilado
COLOR_HEADER = colors.HexColor('#1976D2')       # Azul para encabezados
COLOR_TEXTO_HEADER = colors.white
COLOR_SIN_MATCH = colors.HexColor('#FFF9C4')    # Amarillo claro para sin match

# Columnas del reporte
COLUMNAS_REPORTE = [
    'origen', 'region_desc', 'customer_desc', 'formato', 'center_code', 
    'center_desc', 'rol', 'usuario', 'lunes', 'martes', 'miercoles', 
    'jueves', 'viernes', 'sabado'
]

COLUMNAS_CORTAS = [
    'Origen', 'Región', 'Cliente', 'Formato', 'Código', 'Centro', 'Rol', 'Usuario',
    'L', 'M', 'X', 'J', 'V', 'S'
]

# Campos comparables (donde buscar diferencias para resaltar en rojo)
CAMPOS_COMPARABLES = [
    'region_desc', 'customer_desc', 'formato', 'center_code', 'center_desc', 
    'usuario', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado'
]


def normalizar_valor(valor) -> str:
    """Normaliza un valor para comparación."""
    if pd.isna(valor) or valor is None:
        return ''
    return str(valor).strip().lower()


def encontrar_diferencias(row_maestra: Dict, row_compilado: Dict) -> List[str]:
    """
    Encuentra los campos que tienen diferencias entre Maestra y Compilado.
    Retorna lista de nombres de campos con diferencias.
    """
    diferencias = []
    
    for campo in CAMPOS_COMPARABLES:
        val_maestra = normalizar_valor(row_maestra.get(campo, ''))
        val_compilado = normalizar_valor(row_compilado.get(campo, ''))
        
        if val_maestra != val_compilado:
            diferencias.append(campo)
    
    return diferencias


def preparar_fila_tabla(row: Dict, origen: str, campos: List[str]) -> List[str]:
    """Prepara una fila para la tabla del PDF."""
    fila = [origen]
    
    for campo in campos[1:]:  # Saltamos 'origen' que ya lo agregamos
        campo_lower = campo.lower()
        if campo_lower in row:
            val = row[campo_lower]
            if pd.isna(val) or val is None:
                fila.append('')
            else:
                fila.append(str(val)[:50])  # Limitar longitud
        else:
            fila.append('')
    
    return fila


def generar_pdf_comparacion(coincidencias: List[Dict], 
                            df_maestra: pd.DataFrame,
                            ruta_salida: str,
                            titulo: str = "Reporte de Comparación",
                            coincidencias_exactas: List[Dict] = None,
                            coincidencias_relativas: List[Dict] = None,
                            sin_match: List[Dict] = None) -> str:
    """
    Genera un PDF con la comparación visual de todas las filas.
    
    Args:
        coincidencias: Lista de todas las coincidencias (compatibilidad)
        df_maestra: DataFrame de la Maestra original (antes de cambios)
        ruta_salida: Ruta donde guardar el PDF
        titulo: Título del reporte
        coincidencias_exactas: Lista de coincidencias exactas (opcional)
        coincidencias_relativas: Lista de coincidencias relativas (opcional)
        sin_match: Lista de filas sin coincidencia (opcional)
    
    Returns:
        str: Ruta del archivo generado
    """
    print(f"\n📄 Generando PDF de comparación: {ruta_salida}")
    
    # Si no se pasan separadas, usar todas las coincidencias como exactas
    if coincidencias_exactas is None and coincidencias_relativas is None:
        coincidencias_exactas = coincidencias
        coincidencias_relativas = []
    
    if sin_match is None:
        sin_match = []
    
    # Crear documento
    doc = SimpleDocTemplate(
        ruta_salida,
        pagesize=landscape(A4),
        rightMargin=5*mm,
        leftMargin=5*mm,
        topMargin=10*mm,
        bottomMargin=10*mm
    )
    
    # Estilos
    styles = getSampleStyleSheet()
    titulo_style = ParagraphStyle(
        'TituloReporte',
        parent=styles['Heading1'],
        fontSize=16,
        alignment=TA_CENTER,
        spaceAfter=10*mm
    )
    
    subtitulo_style = ParagraphStyle(
        'Subtitulo',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_CENTER,
        textColor=colors.gray,
        spaceAfter=5*mm
    )
    
    seccion_style = ParagraphStyle(
        'SeccionTitulo',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=COLOR_HEADER,
        spaceBefore=10*mm,
        spaceAfter=5*mm
    )
    
    # Elementos del documento
    elementos = []
    
    # Título
    total_coincidencias = len(coincidencias_exactas) + len(coincidencias_relativas)
    elementos.append(Paragraph(titulo, titulo_style))
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    elementos.append(Paragraph(f"Generado: {fecha} | Total comparaciones: {total_coincidencias} | Sin match: {len(sin_match)}", subtitulo_style))
    elementos.append(Spacer(1, 5*mm))
    
    # Leyenda
    leyenda_data = [
        ['Leyenda:', '', ''],
        ['', 'MODIFICADO', 'Valor que se aplica (cambio detectado)'],
        ['', 'MAESTRA', 'Fila original de Maestra de Rutas'],
        ['', 'COMPILADO', 'Fila del archivo Compilado']
    ]
    leyenda_table = Table(leyenda_data, colWidths=[60, 80, 220])
    leyenda_table.setStyle(TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BACKGROUND', (1, 1), (1, 1), COLOR_MODIFICADO),
        ('BACKGROUND', (1, 2), (1, 2), COLOR_MAESTRA),
        ('BACKGROUND', (1, 3), (1, 3), COLOR_COMPILADO),
        ('ALIGN', (0, 0), (0, 0), 'RIGHT'),
        ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
        ('BOX', (1, 1), (1, 1), 0.5, colors.black),
        ('BOX', (1, 2), (1, 2), 0.5, colors.black),
        ('BOX', (1, 3), (1, 3), 0.5, colors.black),
    ]))
    elementos.append(leyenda_table)
    elementos.append(Spacer(1, 10*mm))
    
    # ==================== SECCIÓN 1: MATCH EXACTO ====================
    if coincidencias_exactas:
        elementos.append(Paragraph(f"📌 COINCIDENCIAS EXACTAS ({len(coincidencias_exactas)})", seccion_style))
        agregar_seccion_coincidencias(elementos, coincidencias_exactas, df_maestra)
    
    # ==================== SECCIÓN 2: MATCH RELATIVO ====================
    if coincidencias_relativas:
        elementos.append(PageBreak())
        elementos.append(Paragraph(f"🔍 COINCIDENCIAS RELATIVAS ({len(coincidencias_relativas)})", seccion_style))
        agregar_seccion_coincidencias(elementos, coincidencias_relativas, df_maestra)
    
    # ==================== SECCIÓN 3: SIN COINCIDENCIA ====================
    if sin_match:
        elementos.append(PageBreak())
        elementos.append(Paragraph(f"⚠️ SIN COINCIDENCIA ({len(sin_match)})", seccion_style))
        agregar_seccion_sin_match(elementos, sin_match)
    
    # Construir PDF
    doc.build(elementos)
    print(f"   ✅ PDF generado exitosamente")
    
    return ruta_salida


def agregar_seccion_coincidencias(elementos: List, coincidencias: List[Dict], df_maestra: pd.DataFrame):
    """Agrega una sección de coincidencias al PDF."""
    comparaciones_por_pagina = 5
    contador = 0
    datos_tabla = [COLUMNAS_CORTAS]
    
    for match in coincidencias:
        maestra_idx = match['maestra_idx']
        comp_row = match['compilado_row']
        
        maestra_row = df_maestra.loc[maestra_idx].to_dict()
        
        fila_modificado = preparar_fila_tabla(comp_row, 'MODIFICADO', COLUMNAS_REPORTE)
        fila_maestra = preparar_fila_tabla(maestra_row, 'MAESTRA', COLUMNAS_REPORTE)
        fila_compilado = preparar_fila_tabla(comp_row, 'COMPILADO', COLUMNAS_REPORTE)
        
        datos_tabla.append(fila_modificado)
        datos_tabla.append(fila_maestra)
        datos_tabla.append(fila_compilado)
        
        contador += 1
        
        if contador % comparaciones_por_pagina == 0 and contador < len(coincidencias):
            tabla = crear_tabla_comparacion(datos_tabla, coincidencias[:contador], df_maestra, comparaciones_por_pagina)
            elementos.append(tabla)
            elementos.append(PageBreak())
            datos_tabla = [COLUMNAS_CORTAS]
    
    if len(datos_tabla) > 1:
        tabla = crear_tabla_comparacion(datos_tabla, coincidencias, df_maestra, comparaciones_por_pagina)
        elementos.append(tabla)


def agregar_seccion_sin_match(elementos: List, sin_match: List[Dict]):
    """Agrega la sección de filas sin coincidencia al PDF."""
    # Columnas para sin match
    columnas_sin_match = ['Región', 'Cliente', 'Formato', 'Código', 'Centro', 'Usuario']
    
    datos_tabla = [columnas_sin_match]
    
    for item in sin_match:
        row = item.get('compilado_row', item)
        fila = [
            str(row.get('region_desc', ''))[:30],
            str(row.get('customer_desc', ''))[:30],
            str(row.get('formato', ''))[:20],
            str(row.get('center_code', ''))[:15],
            str(row.get('center_desc', ''))[:30],
            str(row.get('usuario', ''))[:30]
        ]
        datos_tabla.append(fila)
    
    # Crear tabla
    ancho_disponible = 287 * mm
    col_widths = [ancho_disponible/6] * 6
    
    tabla = Table(datos_tabla, colWidths=col_widths, repeatRows=1)
    tabla.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), COLOR_HEADER),
        ('TEXTCOLOR', (0, 0), (-1, 0), COLOR_TEXTO_HEADER),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 1), (-1, -1), COLOR_SIN_MATCH),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))
    elementos.append(tabla)


def crear_tabla_comparacion(datos: List[List[str]], 
                            coincidencias: List[Dict],
                            df_maestra: pd.DataFrame,
                            comparaciones_por_pagina: int = 5) -> Table:
    """Crea una tabla con estilos para la comparación (3 filas por coincidencia)."""
    
    # Ancho total disponible en landscape A4 (297mm) menos márgenes (10mm total)
    ancho_disponible = 287 * mm
    
    # Proporciones relativas de cada columna (14 columnas total)
    # Origen, Región, Cliente, Formato, Código, Centro, Rol, Usuario, L, M, X, J, V, S
    proporciones = [6, 7, 7, 7, 5, 12, 6, 12, 3, 3, 3, 3, 3, 3]
    total_proporciones = sum(proporciones)
    
    # Calcular anchos proporcionales para usar todo el ancho
    col_widths = [(p / total_proporciones) * ancho_disponible for p in proporciones]
    
    tabla = Table(datos, colWidths=col_widths, repeatRows=1)
    
    # Estilos base
    estilos = [
        # Encabezado
        ('BACKGROUND', (0, 0), (-1, 0), COLOR_HEADER),
        ('TEXTCOLOR', (0, 0), (-1, 0), COLOR_TEXTO_HEADER),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 7),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        
        # Cuerpo
        ('FONTSIZE', (0, 1), (-1, -1), 6),
        ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('WORDWRAP', (0, 0), (-1, -1), True),
        
        # Bordes
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        
        # Padding
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]
    
    # Aplicar colores por fila (grupos de 3: MODIFICADO, MAESTRA, COMPILADO)
    fila_idx = 1  # Empezamos después del encabezado
    match_idx = 0
    
    while fila_idx < len(datos) and match_idx < len(coincidencias):
        match = coincidencias[match_idx]
        maestra_row = df_maestra.loc[match['maestra_idx']].to_dict()
        comp_row = match['compilado_row']
        diferencias = encontrar_diferencias(maestra_row, comp_row)
        
        # Fila MODIFICADO (verde claro - muestra lo que se aplicará)
        estilos.append(('BACKGROUND', (0, fila_idx), (-1, fila_idx), COLOR_MODIFICADO))
        estilos.append(('FONTNAME', (0, fila_idx), (0, fila_idx), 'Helvetica-Bold'))
        
        # Fila MAESTRA (azul claro - original)
        estilos.append(('BACKGROUND', (0, fila_idx + 1), (-1, fila_idx + 1), COLOR_MAESTRA))
        
        # Fila COMPILADO (naranja claro - fuente)
        estilos.append(('BACKGROUND', (0, fila_idx + 2), (-1, fila_idx + 2), COLOR_COMPILADO))
        
        # Línea separadora gruesa después de cada grupo de 3
        estilos.append(('LINEBELOW', (0, fila_idx + 2), (-1, fila_idx + 2), 2, colors.black))
        
        fila_idx += 3  # Avanzar 3 filas
        match_idx += 1
    
    tabla.setStyle(TableStyle(estilos))
    return tabla


if __name__ == "__main__":
    # Test básico
    print("Módulo de generación de PDF cargado correctamente")
