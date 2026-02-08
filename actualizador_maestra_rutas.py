#!/usr/bin/env python3
"""
Sistema de Actualización de Maestra de Rutas
=============================================
Actualiza la planilla Maestra de Rutas basándose en el archivo Compilado,
aplicando lógica de coincidencias exactas y relativas.

Autor: Sistema Automatizado
Fecha: 2026-01-31
Versión: 2.1.0 - Ciclo 2 con normalización center_code (2026-02-02)
"""

import pandas as pd
import numpy as np
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Tuple, Dict, List, Optional
import warnings
warnings.filterwarnings('ignore')


# =============================================================================
# CONFIGURACIÓN
# =============================================================================

# Columnas modificables en la Maestra de Rutas
DIAS_MODIFICABLES = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado']
CAMPO_USUARIO = 'usuario'
CAMPO_ROL = 'rol'
CAMPO_CENTER_CODE = 'center_code'
ROL_MODIFICABLE = 'Supervisor'

# Dominio para emails
DOMINIO_EMAIL = 'castano.cl'


# =============================================================================
# FUNCIONES DE NORMALIZACIÓN
# =============================================================================

def normalizar_texto(texto: str) -> str:
    """
    Normaliza texto: corrige encoding, minúsculas, sin acentos, espacios simples.
    Maneja casos de encoding corrupto (UTF-8 interpretado como Latin-1).
    """
    if pd.isna(texto):
        return ''
    texto = str(texto).strip()
    
    # Intentar corregir encoding doble (UTF-8 leído como Latin-1)
    # Esto arregla casos como "VIÃ'A" → "VIÑA"
    try:
        texto_bytes = texto.encode('latin-1')
        texto = texto_bytes.decode('utf-8')
    except (UnicodeDecodeError, UnicodeEncodeError):
        pass  # Si falla, mantener el texto original
    
    texto = texto.lower()
    # Remover acentos
    texto = unicodedata.normalize('NFD', texto)
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    # Espacios simples
    texto = re.sub(r'\s+', ' ', texto)
    return texto


def normalizar_dia(valor) -> str:
    """
    Normaliza el valor de un día:
    - X/x → 'X'
    - 0/vacío/NaN → '' (vacío)
    """
    if pd.isna(valor):
        return ''
    valor_str = str(valor).strip().upper()
    if valor_str == 'X':
        return 'X'
    return ''


def normalizar_usuario(usuario: str) -> str:
    """
    Convierte el usuario al formato email:
    - Si ya es email → retorna tal cual (en minúsculas)
    - Si es 'Nombre Apellido' → 'nombre.apellido@castano.cl'
    """
    if pd.isna(usuario) or str(usuario).strip() == '':
        return ''
    
    usuario = str(usuario).strip()
    
    # Si ya es email
    if '@' in usuario:
        return usuario.lower()
    
    # Convertir "Nombre Apellido" a email
    partes = usuario.split()
    if len(partes) >= 2:
        nombre = normalizar_texto(partes[0])
        apellido = normalizar_texto(partes[-1])
        return f"{nombre}.{apellido}@{DOMINIO_EMAIL}"
    elif len(partes) == 1:
        return f"{normalizar_texto(partes[0])}@{DOMINIO_EMAIL}"
    
    return usuario.lower()


def extraer_digitos(center_code: str) -> str:
    """Extrae solo los dígitos de un center_code."""
    if pd.isna(center_code):
        return ''
    return re.sub(r'\D', '', str(center_code))


def normalizar_familia(customer_desc: str, formato: str) -> str:
    """
    Crea una 'familia' normalizada combinando customer_desc + formato.
    Ejemplo: 'CENCOSUD' + 'Jumbo' → 'cencosud_jumbo'
    """
    customer = normalizar_texto(customer_desc)
    fmt = normalizar_texto(formato)
    
    # Mapeos conocidos para variaciones
    mapeos_formato = {
        'santa isabel': 'si',
        'express de lider': 'express',
        'hiper lider': 'hiper',
        'mayorista 10': 'm10',
        'super 10': 's10',
    }
    
    fmt = mapeos_formato.get(fmt, fmt.replace(' ', '_'))
    return f"{customer}_{fmt}"


# =============================================================================
# FUNCIONES DE CARGA DE DATOS
# =============================================================================

def cargar_datos(ruta_maestra: str, ruta_compilado: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Carga ambos archivos Excel.
    
    Returns:
        Tuple[DataFrame, DataFrame]: (maestra, compilado)
    """
    print(f"📂 Cargando Maestra de Rutas: {ruta_maestra}")
    df_maestra = pd.read_excel(ruta_maestra)
    print(f"   → {len(df_maestra):,} filas cargadas")
    
    print(f"📂 Cargando Compilado: {ruta_compilado}")
    df_compilado = pd.read_excel(ruta_compilado)
    print(f"   → {len(df_compilado):,} filas cargadas")
    
    # Normalizar nombres de columnas del compilado a minúsculas
    df_compilado.columns = [col.lower() for col in df_compilado.columns]
    
    return df_maestra, df_compilado


# =============================================================================
# ALGORITMOS DE MATCHING
# =============================================================================

def matching_exacto(df_maestra: pd.DataFrame, df_compilado: pd.DataFrame) -> Tuple[List[Dict], List[Dict]]:
    """
    Busca coincidencias exactas usando dos criterios:
    
    CRITERIO 1: center_desc coincide (con o sin otros campos)
    CRITERIO 2: center_code + ≥2 campos coinciden (cliente, formato)
    
    Lógica:
    - Si center_desc coincide → EXACTO (95-100%)
    - Si center_code + cliente + formato coinciden → EXACTO (90-100%)
    - Si no → sin coincidencia exacta (va a matching relativo)
    
    Returns:
        Tuple[List[Dict], List[Dict]]: (coincidencias, sin_coincidencia)
    """
    print("\n🔍 Ejecutando matching exacto (doble criterio)...")
    
    coincidencias = []
    sin_coincidencia = []
    procesados = set()  # Para no procesar dos veces
    
    # Filtrar supervisores en maestra
    maestra_supervisores = df_maestra[df_maestra[CAMPO_ROL] == ROL_MODIFICABLE].copy()
    
    # ========== ÍNDICE 1: Por center_desc ==========
    maestra_por_center_desc = {}
    for idx, row in maestra_supervisores.iterrows():
        center_desc = normalizar_texto(row.get('center_desc', ''))
        if center_desc and center_desc not in maestra_por_center_desc:
            maestra_por_center_desc[center_desc] = {
                'idx': idx,
                'center_code': str(row.get(CAMPO_CENTER_CODE, '')).strip(),
                'cliente': normalizar_texto(row.get('customer_desc', '')),
                'formato': normalizar_texto(row.get('formato', '')),
                'row': row
            }
    
    # ========== ÍNDICE 2: Por center_code ==========
    maestra_por_center_code = {}
    for idx, row in maestra_supervisores.iterrows():
        center_code = str(row.get(CAMPO_CENTER_CODE, '')).strip()
        if center_code and center_code not in maestra_por_center_code:
            maestra_por_center_code[center_code] = {
                'idx': idx,
                'center_desc': normalizar_texto(row.get('center_desc', '')),
                'cliente': normalizar_texto(row.get('customer_desc', '')),
                'formato': normalizar_texto(row.get('formato', '')),
                'row': row
            }
    
    print(f"   📊 Índice center_desc: {len(maestra_por_center_desc)} únicos")
    print(f"   📊 Índice center_code: {len(maestra_por_center_code)} únicos")
    
    # ========== BUSCAR COINCIDENCIAS ==========
    for comp_idx, comp_row in df_compilado.iterrows():
        center_desc_comp = normalizar_texto(comp_row.get('center_desc', ''))
        center_code_comp = str(comp_row.get(CAMPO_CENTER_CODE, '')).strip() if CAMPO_CENTER_CODE in comp_row else ''
        cliente_comp = normalizar_texto(comp_row.get('customer_desc', ''))
        formato_comp = normalizar_texto(comp_row.get('formato', ''))
        
        match_encontrado = False
        
        # CRITERIO 1: Match por center_desc
        if center_desc_comp in maestra_por_center_desc:
            match_data = maestra_por_center_desc[center_desc_comp]
            
            # Contar campos adicionales que coinciden
            campos = 1  # center_desc ya coincide
            if center_code_comp and center_code_comp == match_data['center_code']:
                campos += 1
            if formato_comp and formato_comp == match_data['formato']:
                campos += 1
            if cliente_comp and cliente_comp == match_data['cliente']:
                campos += 1
            
            confianza = min(1.0, 0.90 + (campos * 0.025))  # 90% base + 2.5% por campo
            
            coincidencias.append({
                'compilado_idx': comp_idx,
                'maestra_idx': match_data['idx'],
                'center_code': center_code_comp or match_data['center_code'],
                'tipo_match': 'EXACTO',
                'compilado_row': comp_row.to_dict(),
                'confianza': confianza,
                'criterio': 'center_desc',
                'campos_match': campos
            })
            procesados.add(comp_idx)
            match_encontrado = True
        
        # CRITERIO 2: Match por center_code + campos adicionales
        if not match_encontrado and center_code_comp in maestra_por_center_code:
            match_data = maestra_por_center_code[center_code_comp]
            
            # Contar campos que coinciden (excluyendo center_code que ya coincide)
            campos = 1  # center_code ya coincide
            if formato_comp and formato_comp == match_data['formato']:
                campos += 1
            if cliente_comp and cliente_comp == match_data['cliente']:
                campos += 1
            
            # Requiere al menos 2 campos coincidentes (center_code + otro)
            if campos >= 2:
                confianza = min(1.0, 0.85 + (campos * 0.05))  # 85% base + 5% por campo
                
                coincidencias.append({
                    'compilado_idx': comp_idx,
                    'maestra_idx': match_data['idx'],
                    'center_code': center_code_comp,
                    'tipo_match': 'EXACTO',
                    'compilado_row': comp_row.to_dict(),
                    'confianza': confianza,
                    'criterio': 'center_code+campos',
                    'campos_match': campos
                })
                procesados.add(comp_idx)
                match_encontrado = True
        
        # Sin coincidencia exacta
        if not match_encontrado:
            sin_coincidencia.append({
                'compilado_idx': comp_idx,
                'center_code': center_code_comp,
                'compilado_row': comp_row.to_dict()
            })
    
    print(f"   ✅ Coincidencias exactas: {len(coincidencias)}")
    print(f"   ⚠️  Sin coincidencia exacta: {len(sin_coincidencia)}")
    
    return coincidencias, sin_coincidencia


def extraer_palabras_clave(texto: str) -> set:
    """Extrae palabras clave de un texto (para comparar center_desc)."""
    if pd.isna(texto):
        return set()
    texto_norm = normalizar_texto(texto)
    # Remover palabras comunes/ruido
    palabras_ruido = {'de', 'la', 'el', 'los', 'las', 'y', 'en', 'del'}
    palabras = set(texto_norm.split()) - palabras_ruido
    return palabras


def matching_relativo(df_maestra: pd.DataFrame, sin_coincidencia: List[Dict]) -> Tuple[List[Dict], List[Dict], List[Dict]]:
    """
    Para filas sin match exacto, aplica heurísticas de matching relativo.
    
    Lógica mejorada:
    CICLO 1: Match por (region, familia, dígitos)
    CICLO 2: Match por center_code exacto (cuando formato difiere)
    
    Returns:
        Tuple[List[Dict], List[Dict], List[Dict]]: (coincidencias, ambiguos, sin_match)
    """
    if not sin_coincidencia:
        return [], [], []
    
    print("\n🤖 Ejecutando matching relativo (heurístico mejorado)...")
    
    coincidencias = []
    ambiguos = []
    sin_match_ciclo1 = []
    sin_match_final = []
    
    # Filtrar supervisores en maestra
    maestra_supervisores = df_maestra[df_maestra[CAMPO_ROL] == ROL_MODIFICABLE].copy()
    
    # ==================== CICLO 1: Match por (region, familia, dígitos) ====================
    print("   📌 Ciclo 1: Buscando por región + familia + dígitos...")
    
    # Crear índice por (region, familia, dígitos) en maestra
    maestra_por_clave = {}
    for idx, row in maestra_supervisores.iterrows():
        region = normalizar_texto(row.get('region_desc', ''))
        familia = normalizar_familia(
            row.get('customer_desc', ''), 
            row.get('formato', '')
        )
        digitos = extraer_digitos(str(row[CAMPO_CENTER_CODE]))
        center_desc_palabras = extraer_palabras_clave(row.get('center_desc', ''))
        key = (region, familia, digitos)
        
        if key not in maestra_por_clave:
            maestra_por_clave[key] = {
                'idx': idx,
                'center_desc_palabras': center_desc_palabras
            }
    
    # Buscar cada fila sin coincidencia (Ciclo 1)
    for item in sin_coincidencia:
        comp_row = item['compilado_row']
        
        region = normalizar_texto(comp_row.get('region_desc', ''))
        familia = normalizar_familia(
            comp_row.get('customer_desc', ''), 
            comp_row.get('formato', '')
        )
        digitos = extraer_digitos(item['center_code'])
        center_desc_comp = extraer_palabras_clave(comp_row.get('center_desc', ''))
        key = (region, familia, digitos)
        
        if key in maestra_por_clave:
            maestra_info = maestra_por_clave[key]
            maestra_idx = maestra_info['idx']
            
            # Verificar similitud de center_desc (al menos 1 palabra en común)
            palabras_comunes = center_desc_comp & maestra_info['center_desc_palabras']
            
            if len(palabras_comunes) > 0 or len(center_desc_comp) == 0:
                # Match encontrado con verificación de center_desc
                confianza = 0.90 if len(palabras_comunes) > 0 else 0.75
                coincidencias.append({
                    'compilado_idx': item['compilado_idx'],
                    'maestra_idx': maestra_idx,
                    'center_code': item['center_code'],
                    'tipo_match': 'RELATIVO',
                    'compilado_row': comp_row,
                    'confianza': confianza,
                    'region': region,
                    'familia': familia,
                    'digitos': digitos,
                    'palabras_comunes': list(palabras_comunes)
                })
            else:
                # center_desc no coincide - caso ambiguo
                ambiguos.append({
                    'compilado_idx': item['compilado_idx'],
                    'center_code': item['center_code'],
                    'compilado_row': comp_row,
                    'familia': familia,
                    'digitos': digitos,
                    'motivo': 'center_desc no coincide'
                })
        else:
            # Sin candidatos en Ciclo 1 - pasar a Ciclo 2
            sin_match_ciclo1.append(item)
    
    print(f"      Ciclo 1 - Coincidencias: {len(coincidencias)}")
    print(f"      Ciclo 1 - Pendientes para Ciclo 2: {len(sin_match_ciclo1)}")
    
    # ==================== CICLO 2: Match por center_code exacto ====================
    if sin_match_ciclo1:
        print("   🔍 Ciclo 2: Buscando por center_code (ignora diferencia de formato)...")
        
        # Crear índice por center_code NORMALIZADO en maestra (solo dígitos)
        maestra_por_center = {}
        for idx, row in maestra_supervisores.iterrows():
            center_code_raw = str(row[CAMPO_CENTER_CODE]).strip()
            center_code_norm = extraer_digitos(center_code_raw)  # Normalizar a solo dígitos
            center_desc_palabras = extraer_palabras_clave(row.get('center_desc', ''))
            
            if center_code_norm not in maestra_por_center:
                maestra_por_center[center_code_norm] = []
            
            maestra_por_center[center_code_norm].append({
                'idx': idx,
                'row': row,
                'center_code_original': center_code_raw,
                'center_desc_palabras': center_desc_palabras
            })
        
        print(f"      DEBUG: Índice maestra tiene {len(maestra_por_center)} center_codes únicos")
        
        for item in sin_match_ciclo1:
            comp_row = item['compilado_row']
            center_code_raw = str(item['center_code']).strip()
            center_code_norm = extraer_digitos(center_code_raw)  # Normalizar a solo dígitos
            center_desc_comp = extraer_palabras_clave(comp_row.get('center_desc', ''))
            
            print(f"      DEBUG: Buscando center_code normalizado: '{center_code_norm}' (original: '{center_code_raw}')")
            
            # Buscar por center_code normalizado
            if center_code_norm in maestra_por_center:
                candidatos = maestra_por_center[center_code_norm]
                
                # Buscar el candidato con mayor similitud en center_desc
                mejor_match = None
                mejor_score = -1  # Cambiado a -1 para aceptar score 0
                
                for candidato in candidatos:
                    palabras_comunes = center_desc_comp & candidato['center_desc_palabras']
                    score = len(palabras_comunes)
                    
                    if score > mejor_score:
                        mejor_score = score
                        mejor_match = candidato
                
                # Si no hay ningún candidato con palabras en común, tomar el primero
                if mejor_match is None and candidatos:
                    mejor_match = candidatos[0]
                    mejor_score = 0
                
                if mejor_match is not None:
                    # Match por center_code encontrado
                    confianza = 0.70 if mejor_score > 0 else 0.50  # Menor confianza sin palabras comunes
                    coincidencias.append({
                        'compilado_idx': item['compilado_idx'],
                        'maestra_idx': mejor_match['idx'],
                        'center_code': center_code_raw,
                        'tipo_match': 'RELATIVO_CENTERCODE',
                        'compilado_row': comp_row,
                        'confianza': confianza,
                        'region': normalizar_texto(comp_row.get('region_desc', '')),
                        'familia': 'FORMATO_DIFERENTE',
                        'digitos': center_code_norm,
                        'palabras_comunes': list(center_desc_comp & mejor_match['center_desc_palabras']) if mejor_match else []
                    })
                else:
                    # center_code existe pero no hay match de palabras
                    sin_match_final.append({
                        'compilado_idx': item['compilado_idx'],
                        'center_code': center_code_raw,
                        'compilado_row': comp_row,
                        'region': normalizar_texto(comp_row.get('region_desc', '')),
                        'familia': normalizar_familia(comp_row.get('customer_desc', ''), comp_row.get('formato', '')),
                        'digitos': center_code_norm
                    })
            else:
                # center_code no existe en maestra
                sin_match_final.append({
                    'compilado_idx': item['compilado_idx'],
                    'center_code': center_code_raw,
                    'compilado_row': comp_row,
                    'region': normalizar_texto(comp_row.get('region_desc', '')),
                    'familia': normalizar_familia(comp_row.get('customer_desc', ''), comp_row.get('formato', '')),
                    'digitos': center_code_norm
                })
        
        print(f"      Ciclo 2 - Coincidencias adicionales: {len(coincidencias) - len([c for c in coincidencias if c['tipo_match'] == 'RELATIVO'])}")
    
    print(f"   ✅ Coincidencias relativas totales: {len(coincidencias)}")
    print(f"   ⚠️  Casos ambiguos: {len(ambiguos)}")
    print(f"   ❌ Sin coincidencia: {len(sin_match_final)}")
    
    return coincidencias, ambiguos, sin_match_final


# =============================================================================
# APLICACIÓN DE CAMBIOS
# =============================================================================

def aplicar_cambios(df_maestra: pd.DataFrame, coincidencias: List[Dict]) -> Tuple[pd.DataFrame, List[Dict]]:
    """
    Aplica los cambios del compilado a la maestra.
    
    Returns:
        Tuple[DataFrame, List[Dict]]: (maestra_actualizada, log_cambios)
    """
    print("\n📝 Aplicando cambios a la Maestra de Rutas...")
    
    df_actualizado = df_maestra.copy()
    log_cambios = []
    
    for match in coincidencias:
        maestra_idx = match['maestra_idx']
        comp_row = match['compilado_row']
        center_code = match['center_code']
        tipo_match = match['tipo_match']
        
        # Actualizar usuario
        nuevo_usuario = normalizar_usuario(comp_row.get(CAMPO_USUARIO, ''))
        if nuevo_usuario:
            valor_anterior = df_actualizado.at[maestra_idx, CAMPO_USUARIO]
            if pd.isna(valor_anterior):
                valor_anterior = ''
            
            if str(valor_anterior).lower() != nuevo_usuario:
                df_actualizado.at[maestra_idx, CAMPO_USUARIO] = nuevo_usuario
                log_cambios.append({
                    'center_code': center_code,
                    'campo': CAMPO_USUARIO,
                    'valor_anterior': valor_anterior,
                    'valor_nuevo': nuevo_usuario,
                    'tipo_match': tipo_match
                })
        
        # Actualizar días
        for dia in DIAS_MODIFICABLES:
            if dia in comp_row:
                nuevo_valor = normalizar_dia(comp_row[dia])
                valor_anterior = df_actualizado.at[maestra_idx, dia]
                
                if pd.isna(valor_anterior):
                    valor_anterior_str = ''
                else:
                    valor_anterior_str = str(valor_anterior).strip()
                
                # Solo registrar si hay cambio real
                if valor_anterior_str != nuevo_valor:
                    df_actualizado.at[maestra_idx, dia] = nuevo_valor if nuevo_valor else np.nan
                    log_cambios.append({
                        'center_code': center_code,
                        'campo': dia,
                        'valor_anterior': valor_anterior_str if valor_anterior_str else '(vacío)',
                        'valor_nuevo': nuevo_valor if nuevo_valor else '(vacío)',
                        'tipo_match': tipo_match
                    })
    
    print(f"   ✅ Total cambios aplicados: {len(log_cambios)}")
    
    return df_actualizado, log_cambios


# =============================================================================
# GENERACIÓN DE KPIS Y REPORTES
# =============================================================================

def calcular_kpis(total_compilado: int, 
                  coincidencias_exactas: int,
                  coincidencias_relativas: int,
                  filas_actualizadas: int) -> Dict:
    """Calcula los KPIs del proceso."""
    
    total_coincidencias = coincidencias_exactas + coincidencias_relativas
    
    kpis = {
        'total_filas_compilado': total_compilado,
        'coincidencias_exactas': coincidencias_exactas,
        'coincidencias_relativas': coincidencias_relativas,
        'total_coincidencias': total_coincidencias,
        'filas_actualizadas': filas_actualizadas,
        'pct_matching_exacto': (coincidencias_exactas / total_compilado * 100) if total_compilado > 0 else 0,
        'pct_matching_relativo': (coincidencias_relativas / total_compilado * 100) if total_compilado > 0 else 0,
        'pct_filas_actualizadas': (filas_actualizadas / total_compilado * 100) if total_compilado > 0 else 0,
        'pct_total_matching': (total_coincidencias / total_compilado * 100) if total_compilado > 0 else 0
    }
    
    return kpis


def generar_reporte(kpis: Dict, 
                    log_cambios: List[Dict],
                    ambiguos: List[Dict],
                    sin_match: List[Dict],
                    ruta_salida: str):
    """Genera el reporte de auditoría en Excel."""
    
    print(f"\n📊 Generando reporte: {ruta_salida}")
    
    with pd.ExcelWriter(ruta_salida, engine='xlsxwriter') as writer:
        # Hoja 1: KPIs
        df_kpis = pd.DataFrame([
            {'Métrica': 'Total filas en Compilado', 'Valor': kpis['total_filas_compilado']},
            {'Métrica': 'Coincidencias exactas', 'Valor': kpis['coincidencias_exactas']},
            {'Métrica': 'Coincidencias relativas (ML)', 'Valor': kpis['coincidencias_relativas']},
            {'Métrica': 'Total coincidencias', 'Valor': kpis['total_coincidencias']},
            {'Métrica': 'Filas con cambios aplicados', 'Valor': kpis['filas_actualizadas']},
            {'Métrica': '% Matching exacto', 'Valor': f"{kpis['pct_matching_exacto']:.2f}%"},
            {'Métrica': '% Matching relativo', 'Valor': f"{kpis['pct_matching_relativo']:.2f}%"},
            {'Métrica': '% Total matching', 'Valor': f"{kpis['pct_total_matching']:.2f}%"},
            {'Métrica': '% Filas actualizadas', 'Valor': f"{kpis['pct_filas_actualizadas']:.2f}%"},
        ])
        df_kpis.to_excel(writer, sheet_name='KPIs', index=False)
        
        # Hoja 2: Detalle de cambios
        if log_cambios:
            df_cambios = pd.DataFrame(log_cambios)
            df_cambios.to_excel(writer, sheet_name='Detalle Cambios', index=False)
        
        # Hoja 3: Casos ambiguos
        if ambiguos:
            df_ambiguos = pd.DataFrame([{
                'center_code': a['center_code'],
                'familia': a.get('familia', ''),
                'digitos': a.get('digitos', ''),
                'num_candidatos': len(a.get('candidatos', []))
            } for a in ambiguos])
            df_ambiguos.to_excel(writer, sheet_name='Casos Ambiguos', index=False)
        
        # Hoja 4: Sin coincidencia
        if sin_match:
            df_sin_match = pd.DataFrame([{
                'center_code': s['center_code'],
                'familia': s.get('familia', ''),
                'digitos': s.get('digitos', '')
            } for s in sin_match])
            df_sin_match.to_excel(writer, sheet_name='Sin Coincidencia', index=False)
    
    print("   ✅ Reporte generado exitosamente")


def guardar_maestra_actualizada(df_maestra: pd.DataFrame, ruta_salida: str):
    """Guarda la maestra actualizada en Excel con el nombre de hoja original."""
    print(f"\n💾 Guardando Maestra actualizada: {ruta_salida}")
    
    # Usar ExcelWriter para especificar el nombre de la hoja
    with pd.ExcelWriter(ruta_salida, engine='xlsxwriter') as writer:
        df_maestra.to_excel(writer, sheet_name='Maestra_de_Rutas', index=False)
    
    print("   ✅ Archivo guardado exitosamente")


# =============================================================================
# FUNCIÓN PRINCIPAL
# =============================================================================

def main():
    """Función principal del sistema de actualización."""
    
    print("=" * 60)
    print("SISTEMA DE ACTUALIZACIÓN DE MAESTRA DE RUTAS")
    print("=" * 60)
    
    # Configurar rutas
    directorio = Path(__file__).parent
    ruta_maestra = directorio / "Maestra_de_rutas_Castaño.xlsx"
    ruta_compilado = directorio / "compilado Alvaro Sauterer.xlsx"
    
    # Fecha para archivos de salida
    fecha = datetime.now().strftime("%Y-%m-%d_%H%M")
    ruta_maestra_actualizada = directorio / f"Maestra_de_rutas_ACTUALIZADA_{fecha}.xlsx"
    ruta_reporte = directorio / f"Reporte_Actualizacion_{fecha}.xlsx"
    
    # 1. Cargar datos
    df_maestra, df_compilado = cargar_datos(str(ruta_maestra), str(ruta_compilado))
    total_compilado = len(df_compilado)
    
    # 2. Matching exacto
    coincidencias_exactas, sin_coincidencia = matching_exacto(df_maestra, df_compilado)
    
    # 3. Matching relativo (para filas sin coincidencia exacta)
    coincidencias_relativas, ambiguos, sin_match = matching_relativo(df_maestra, sin_coincidencia)
    
    # 4. Combinar todas las coincidencias
    todas_coincidencias = coincidencias_exactas + coincidencias_relativas
    
    # 5. Aplicar cambios
    df_maestra_actualizada, log_cambios = aplicar_cambios(df_maestra, todas_coincidencias)
    
    # Contar filas únicas actualizadas
    filas_unicas = len(set(c['center_code'] for c in todas_coincidencias))
    
    # 6. Calcular KPIs
    kpis = calcular_kpis(
        total_compilado=total_compilado,
        coincidencias_exactas=len(coincidencias_exactas),
        coincidencias_relativas=len(coincidencias_relativas),
        filas_actualizadas=filas_unicas
    )
    
    # 7. Mostrar resumen
    print("\n" + "=" * 60)
    print("RESUMEN DE RESULTADOS")
    print("=" * 60)
    print(f"📊 Total filas en Compilado: {kpis['total_filas_compilado']}")
    print(f"✅ Matching exacto: {kpis['coincidencias_exactas']} ({kpis['pct_matching_exacto']:.2f}%)")
    print(f"🤖 Matching relativo: {kpis['coincidencias_relativas']} ({kpis['pct_matching_relativo']:.2f}%)")
    print(f"📈 Total matching: {kpis['total_coincidencias']} ({kpis['pct_total_matching']:.2f}%)")
    print(f"📝 Filas actualizadas: {kpis['filas_actualizadas']} ({kpis['pct_filas_actualizadas']:.2f}%)")
    print(f"🔧 Total cambios realizados: {len(log_cambios)}")
    
    if ambiguos:
        print(f"⚠️  Casos ambiguos (revisar): {len(ambiguos)}")
    if sin_match:
        print(f"❌ Sin coincidencia: {len(sin_match)}")
    
    # 8. Guardar archivos
    guardar_maestra_actualizada(df_maestra_actualizada, str(ruta_maestra_actualizada))
    generar_reporte(kpis, log_cambios, ambiguos, sin_match, str(ruta_reporte))
    
    # 9. Generar PDF de comparación
    ruta_pdf = directorio / f"Comparacion_Visual_{fecha}.pdf"
    try:
        from generador_pdf import generar_pdf_comparacion
        generar_pdf_comparacion(
            todas_coincidencias,
            df_maestra,  # Original antes de cambios
            str(ruta_pdf),
            titulo="Comparación Maestra vs Compilado"
        )
        pdf_generado = True
    except Exception as e:
        print(f"⚠️  No se pudo generar PDF: {e}")
        pdf_generado = False
    
    print("\n" + "=" * 60)
    print("✅ PROCESO COMPLETADO EXITOSAMENTE")
    print("=" * 60)
    print(f"\n📁 Archivos generados:")
    print(f"   • {ruta_maestra_actualizada.name}")
    print(f"   • {ruta_reporte.name}")
    if pdf_generado:
        print(f"   • {ruta_pdf.name}")
    
    return df_maestra_actualizada, kpis, log_cambios


if __name__ == "__main__":
    main()
