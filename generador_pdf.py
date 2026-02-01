#!/usr/bin/env python3
"""
Generador de PDF con comparación visual Maestra vs Compilado
============================================================
Genera un PDF mostrando las filas comparadas con diferencias resaltadas.
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from typing import List, Dict, Tuple
import pandas as pd
from datetime import datetime


# Colores para el reporte
COLOR_MODIFICADO = colors.HexColor('#C8E6C9')   # Verde claro para modificado
COLOR_DIFERENCIA = colors.HexColor('#FFCDD2')  # Rojo claro para diferencias
COLOR_MAESTRA = colors.HexColor('#E3F2FD')     # Azul muy claro para identificar Maestra
COLOR_COMPILADO = colors.HexColor('#FFF3E0')   # Naranja muy claro para Compilado
COLOR_HEADER = colors.HexColor('#1976D2')      # Azul para encabezados
COLOR_TEXTO_HEADER = colors.white

# Columnas a mostrar en el reporte
COLUMNAS_REPORTE = [
    'origen', 'region_desc', 'customer_desc', 'formato', 'center_code', 
    'center_desc', 'rol', 'usuario', 'lunes', 'martes', 'miercoles', 
    'jueves', 'viernes', 'sabado'
]

COLUMNAS_CORTAS = [
    'Origen', 'Región', 'Cliente', 'Formato', 'Código', 
    'Centro', 'Rol', 'Usuario', 'L', 'M', 'X', 
    'J', 'V', 'S'
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
    val_str = str(valor).strip().upper()
    
    # Valores que significan "vacío" o "no marcado"
    if val_str in ['0', '0.0', 'NAN', 'NONE', 'N', 'NO', '']:
        return ''
    
    # Valores que significan "marcado" - normalizar todos a 'X'
    if val_str in ['X', '1', '1.0', 'SI', 'S', 'Y', 'YES', 'TRUE']:
        return 'X'
    
    return val_str


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
                            titulo: str = "Reporte de Comparación") -> str:
    """
    Genera un PDF con la comparación visual de todas las filas.
    
    Args:
        coincidencias: Lista de diccionarios con las coincidencias encontradas
        df_maestra: DataFrame de la Maestra original (antes de cambios)
        ruta_salida: Ruta donde guardar el PDF
        titulo: Título del reporte
    
    Returns:
        str: Ruta del archivo generado
    """
    print(f"\n📄 Generando PDF de comparación: {ruta_salida}")
    
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
    
    # Elementos del documento
    elementos = []
    
    # Título
    elementos.append(Paragraph(titulo, titulo_style))
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    elementos.append(Paragraph(f"Generado: {fecha} | Total comparaciones: {len(coincidencias)}", subtitulo_style))
    elementos.append(Spacer(1, 5*mm))
    
    # Leyenda
    leyenda_data = [
        ['Leyenda:', '', ''],
        ['', 'MODIFICADO', 'Valor que se aplica (cambio detectado)'],
        ['', 'MAESTRA', 'Fila original de Maestra de Rutas'],
        ['', 'COMPILADO', 'Fila del archivo Compilado'],
        ['', '█ Rojo', 'Celda con diferencia detectada']
    ]
    leyenda_table = Table(leyenda_data, colWidths=[60, 80, 220])
    leyenda_table.setStyle(TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BACKGROUND', (1, 1), (1, 1), COLOR_MODIFICADO),
        ('BACKGROUND', (1, 2), (1, 2), COLOR_MAESTRA),
        ('BACKGROUND', (1, 3), (1, 3), COLOR_COMPILADO),
        ('BACKGROUND', (1, 4), (1, 4), COLOR_DIFERENCIA),
        ('ALIGN', (0, 0), (0, 0), 'RIGHT'),
        ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
        ('BOX', (1, 1), (1, 1), 0.5, colors.black),
        ('BOX', (1, 2), (1, 2), 0.5, colors.black),
        ('BOX', (1, 3), (1, 3), 0.5, colors.black),
        ('BOX', (1, 4), (1, 4), 0.5, colors.black),
    ]))
    elementos.append(leyenda_table)
    elementos.append(Spacer(1, 10*mm))
    
    # Procesar cada coincidencia
    comparaciones_por_pagina = 5  # Grupos de 3 filas por página (reducido por 3 filas)
    contador = 0
    datos_tabla = []
    
    # Encabezados
    datos_tabla.append(COLUMNAS_CORTAS)
    
    for match in coincidencias:
        maestra_idx = match['maestra_idx']
        comp_row = match['compilado_row']
        center_code = match['center_code']
        
        # Obtener fila de la maestra
        maestra_row = df_maestra.loc[maestra_idx].to_dict()
        
        # Encontrar diferencias
        diferencias = encontrar_diferencias(maestra_row, comp_row)
        
        # Crear fila MODIFICADO (muestra los valores nuevos que se aplican)
        fila_modificado = preparar_fila_tabla(comp_row, 'MODIFICADO', COLUMNAS_REPORTE)
        fila_maestra = preparar_fila_tabla(maestra_row, 'MAESTRA', COLUMNAS_REPORTE)
        fila_compilado = preparar_fila_tabla(comp_row, 'COMPILADO', COLUMNAS_REPORTE)
        
        # Orden: MODIFICADO (verde), MAESTRA (azul), COMPILADO (naranja)
        datos_tabla.append(fila_modificado)
        datos_tabla.append(fila_maestra)
        datos_tabla.append(fila_compilado)
        
        contador += 1
        
        # Agregar separador cada grupo de 3 filas
        if contador % comparaciones_por_pagina == 0 and contador < len(coincidencias):
            # Crear tabla actual y agregar salto de página
            tabla = crear_tabla_comparacion(datos_tabla, coincidencias[:contador], df_maestra, comparaciones_por_pagina)
            elementos.append(tabla)
            elementos.append(PageBreak())
            datos_tabla = [COLUMNAS_CORTAS]  # Reiniciar con encabezados
    
    # Agregar última tabla si hay datos
    if len(datos_tabla) > 1:
        tabla = crear_tabla_comparacion(datos_tabla, coincidencias, df_maestra, comparaciones_por_pagina)
        elementos.append(tabla)
    
    # Construir PDF
    doc.build(elementos)
    print(f"   ✅ PDF generado exitosamente")
    
    return ruta_salida


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
        
        # Resaltar diferencias en la fila MODIFICADO
        for campo in diferencias:
            if campo in COLUMNAS_REPORTE:
                col_idx = COLUMNAS_REPORTE.index(campo)
                # Resaltar solo en MODIFICADO con rojo
                estilos.append(('BACKGROUND', (col_idx, fila_idx), (col_idx, fila_idx), COLOR_DIFERENCIA))
                estilos.append(('FONTNAME', (col_idx, fila_idx), (col_idx, fila_idx), 'Helvetica-Bold'))
        
        # Línea separadora gruesa después de cada grupo de 3
        estilos.append(('LINEBELOW', (0, fila_idx + 2), (-1, fila_idx + 2), 2, colors.black))
        
        fila_idx += 3  # Avanzar 3 filas
        match_idx += 1
    
    tabla.setStyle(TableStyle(estilos))
    return tabla


if __name__ == "__main__":
    # Test básico
    print("Módulo de generación de PDF cargado correctamente")
