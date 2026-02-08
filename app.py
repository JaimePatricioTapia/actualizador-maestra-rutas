#!/usr/bin/env python3
"""
Interfaz Web para Sistema de Actualización de Maestra de Rutas
==============================================================
Aplicación Flask con interfaz moderna para cargar archivos y procesar.
"""

import os
import sys
from pathlib import Path
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from werkzeug.utils import secure_filename

# Importar módulos del sistema
from actualizador_maestra_rutas import (
    cargar_datos, matching_exacto, matching_relativo, 
    aplicar_cambios, calcular_kpis, generar_reporte,
    guardar_maestra_actualizada
)
from generador_pdf import generar_pdf_comparacion

# Configuración
app = Flask(__name__)
app.secret_key = 'maestra_rutas_secret_key_2024'

# Directorio para archivos subidos y generados
UPLOAD_FOLDER = Path(__file__).parent / 'uploads'
OUTPUT_FOLDER = Path(__file__).parent / 'output'
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)
app.config['OUTPUT_FOLDER'] = str(OUTPUT_FOLDER)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# Template HTML integrado
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Actualizador Maestra de Rutas</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
            min-height: 100vh;
            color: #fff;
        }
        
        .container {
            max-width: 1000px;
            margin: 0 auto;
            padding: 40px 20px;
        }
        
        header {
            text-align: center;
            margin-bottom: 40px;
        }
        
        h1 {
            font-size: 2.5rem;
            font-weight: 700;
            background: linear-gradient(90deg, #00d4ff, #7b2cbf);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            margin-bottom: 10px;
        }
        
        .subtitle {
            color: #8892b0;
            font-size: 1.1rem;
        }
        
        .card {
            background: rgba(255, 255, 255, 0.05);
            backdrop-filter: blur(10px);
            border-radius: 20px;
            padding: 30px;
            margin-bottom: 30px;
            border: 1px solid rgba(255, 255, 255, 0.1);
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
        }
        
        .card h2 {
            font-size: 1.3rem;
            margin-bottom: 20px;
            color: #00d4ff;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .upload-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
        }
        
        @media (max-width: 768px) {
            .upload-grid {
                grid-template-columns: 1fr;
            }
        }
        
        .upload-zone {
            border: 2px dashed rgba(0, 212, 255, 0.3);
            border-radius: 15px;
            padding: 40px 20px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            position: relative;
        }
        
        .upload-zone:hover {
            border-color: #00d4ff;
            background: rgba(0, 212, 255, 0.05);
        }
        
        .upload-zone.dragover {
            border-color: #7b2cbf;
            background: rgba(123, 44, 191, 0.1);
            transform: scale(1.02);
        }
        
        .upload-zone.has-file {
            border-color: #00ff88;
            background: rgba(0, 255, 136, 0.05);
        }
        
        .upload-zone input[type="file"] {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            opacity: 0;
            cursor: pointer;
        }
        
        .upload-icon {
            font-size: 3rem;
            margin-bottom: 15px;
        }
        
        .upload-label {
            font-size: 1rem;
            font-weight: 500;
            margin-bottom: 5px;
        }
        
        .upload-hint {
            font-size: 0.85rem;
            color: #8892b0;
        }
        
        .file-name {
            margin-top: 10px;
            padding: 8px 15px;
            background: rgba(0, 255, 136, 0.1);
            border-radius: 8px;
            font-size: 0.9rem;
            color: #00ff88;
            display: none;
        }
        
        .upload-zone.has-file .file-name {
            display: inline-block;
        }
        
        .btn {
            padding: 15px 40px;
            font-size: 1.1rem;
            font-weight: 600;
            border: none;
            border-radius: 10px;
            cursor: pointer;
            transition: all 0.3s ease;
            display: inline-flex;
            align-items: center;
            gap: 10px;
        }
        
        .btn-primary {
            background: linear-gradient(90deg, #00d4ff, #7b2cbf);
            color: #fff;
            width: 100%;
            justify-content: center;
        }
        
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 30px rgba(0, 212, 255, 0.3);
        }
        
        .btn-primary:disabled {
            opacity: 0.5;
            cursor: not-allowed;
            transform: none;
        }
        
        .btn-secondary {
            background: rgba(255, 255, 255, 0.1);
            color: #fff;
            border: 1px solid rgba(255, 255, 255, 0.2);
        }
        
        .btn-secondary:hover {
            background: rgba(255, 255, 255, 0.2);
        }
        
        .results {
            display: none;
        }
        
        .results.show {
            display: block;
        }
        
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 15px;
            margin-bottom: 25px;
        }
        
        @media (max-width: 768px) {
            .kpi-grid {
                grid-template-columns: repeat(2, 1fr);
            }
        }
        
        .kpi-card {
            background: rgba(255, 255, 255, 0.05);
            border-radius: 12px;
            padding: 20px;
            text-align: center;
        }
        
        .kpi-value {
            font-size: 2rem;
            font-weight: 700;
            background: linear-gradient(90deg, #00ff88, #00d4ff);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        
        .kpi-label {
            font-size: 0.85rem;
            color: #8892b0;
            margin-top: 5px;
        }
        
        .download-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 15px;
        }
        
        @media (max-width: 768px) {
            .download-grid {
                grid-template-columns: 1fr;
            }
        }
        
        .download-btn {
            padding: 15px 20px;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid rgba(255, 255, 255, 0.1);
            border-radius: 10px;
            color: #fff;
            text-decoration: none;
            display: flex;
            align-items: center;
            gap: 10px;
            transition: all 0.3s ease;
        }
        
        .download-btn:hover {
            background: rgba(0, 212, 255, 0.1);
            border-color: #00d4ff;
        }
        
        .download-icon {
            font-size: 1.5rem;
        }
        
        .download-info {
            flex: 1;
        }
        
        .download-title {
            font-weight: 600;
            margin-bottom: 3px;
        }
        
        .download-desc {
            font-size: 0.8rem;
            color: #8892b0;
        }
        
        .loading {
            display: none;
            text-align: center;
            padding: 40px;
        }
        
        .loading.show {
            display: block;
        }
        
        .spinner {
            width: 50px;
            height: 50px;
            border: 3px solid rgba(255, 255, 255, 0.1);
            border-top-color: #00d4ff;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }
        
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        
        .alert {
            padding: 15px 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .alert-error {
            background: rgba(255, 82, 82, 0.1);
            border: 1px solid rgba(255, 82, 82, 0.3);
            color: #ff5252;
        }
        
        .alert-success {
            background: rgba(0, 255, 136, 0.1);
            border: 1px solid rgba(0, 255, 136, 0.3);
            color: #00ff88;
        }
        
        footer {
            text-align: center;
            margin-top: 40px;
            color: #8892b0;
            font-size: 0.9rem;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📊 Actualizador Maestra de Rutas</h1>
            <p class="subtitle">Sistema de sincronización y comparación de planillas</p>
        </header>
        
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="alert alert-{{ category }}">
                        {{ message }}
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}
        
        <form id="uploadForm" method="POST" action="/procesar" enctype="multipart/form-data">
            <div class="card">
                <h2>📁 Cargar Archivos</h2>
                <div class="upload-grid">
                    <div class="upload-zone" id="maestraZone">
                        <input type="file" name="maestra" id="maestraInput" accept=".xlsx,.xls" required>
                        <div class="upload-icon">📋</div>
                        <div class="upload-label">Maestra de Rutas</div>
                        <div class="upload-hint">Arrastra o haz clic para seleccionar</div>
                        <div class="file-name" id="maestraFileName"></div>
                    </div>
                    <div class="upload-zone" id="compiladoZone">
                        <input type="file" name="compilado" id="compiladoInput" accept=".xlsx,.xls" required>
                        <div class="upload-icon">📝</div>
                        <div class="upload-label">Archivo Compilado</div>
                        <div class="upload-hint">Arrastra o haz clic para seleccionar</div>
                        <div class="file-name" id="compiladoFileName"></div>
                    </div>
                </div>
            </div>
            
            <div class="card">
                <button type="submit" class="btn btn-primary" id="processBtn" disabled>
                    <span>🚀</span>
                    <span>Procesar Archivos</span>
                </button>
            </div>
        </form>
        
        <div class="loading" id="loadingSection">
            <div class="spinner"></div>
            <p>Procesando archivos... Por favor espera.</p>
        </div>
        
        {% if resultados %}
        <div class="results show">
            <div class="card">
                <h2>📈 Resultados</h2>
                <div class="kpi-grid">
                    <div class="kpi-card">
                        <div class="kpi-value">{{ resultados.pct_matching_exacto }}%</div>
                        <div class="kpi-label">Match Exacto</div>
                    </div>
                    <div class="kpi-card">
                        <div class="kpi-value">{{ resultados.pct_matching_relativo }}%</div>
                        <div class="kpi-label">Match Relativo</div>
                    </div>
                    <div class="kpi-card">
                        <div class="kpi-value">{{ resultados.pct_filas_actualizadas }}%</div>
                        <div class="kpi-label">Filas Actualizadas</div>
                    </div>
                    <div class="kpi-card">
                        <div class="kpi-value">{{ resultados.total_cambios }}</div>
                        <div class="kpi-label">Total Cambios</div>
                    </div>
                </div>
                
                <h2>📥 Descargar Archivos</h2>
                <div class="download-grid">
                    <a href="/download/{{ resultados.archivo_maestra }}" class="download-btn" download="{{ resultados.archivo_maestra }}">
                        <span class="download-icon">📊</span>
                        <div class="download-info">
                            <div class="download-title">Maestra Actualizada</div>
                            <div class="download-desc">Excel con cambios aplicados</div>
                        </div>
                    </a>
                    <a href="/download/{{ resultados.archivo_reporte }}" class="download-btn" download="{{ resultados.archivo_reporte }}">
                        <span class="download-icon">📈</span>
                        <div class="download-info">
                            <div class="download-title">Reporte KPIs</div>
                            <div class="download-desc">Detalle de cambios</div>
                        </div>
                    </a>
                    <a href="/download/{{ resultados.archivo_pdf }}" class="download-btn" download="{{ resultados.archivo_pdf }}">
                        <span class="download-icon">📄</span>
                        <div class="download-info">
                            <div class="download-title">Comparación PDF</div>
                            <div class="download-desc">Visual de diferencias</div>
                        </div>
                    </a>
                </div>
            </div>
        </div>
        {% endif %}
        
        <footer>
            <p>Sistema de Actualización de Maestra de Rutas v1.0 | {{ current_year }}</p>
        </footer>
    </div>
    
    <script>
        // File upload handling
        const maestraInput = document.getElementById('maestraInput');
        const compiladoInput = document.getElementById('compiladoInput');
        const maestraZone = document.getElementById('maestraZone');
        const compiladoZone = document.getElementById('compiladoZone');
        const processBtn = document.getElementById('processBtn');
        const uploadForm = document.getElementById('uploadForm');
        const loadingSection = document.getElementById('loadingSection');
        
        function updateFileName(input, zone, fileNameEl) {
            if (input.files.length > 0) {
                document.getElementById(fileNameEl).textContent = input.files[0].name;
                zone.classList.add('has-file');
            } else {
                zone.classList.remove('has-file');
            }
            checkFilesSelected();
        }
        
        function checkFilesSelected() {
            if (maestraInput.files.length > 0 && compiladoInput.files.length > 0) {
                processBtn.disabled = false;
            } else {
                processBtn.disabled = true;
            }
        }
        
        maestraInput.addEventListener('change', () => updateFileName(maestraInput, maestraZone, 'maestraFileName'));
        compiladoInput.addEventListener('change', () => updateFileName(compiladoInput, compiladoZone, 'compiladoFileName'));
        
        // Drag and drop
        ['maestraZone', 'compiladoZone'].forEach(id => {
            const zone = document.getElementById(id);
            
            zone.addEventListener('dragover', (e) => {
                e.preventDefault();
                zone.classList.add('dragover');
            });
            
            zone.addEventListener('dragleave', () => {
                zone.classList.remove('dragover');
            });
            
            zone.addEventListener('drop', (e) => {
                e.preventDefault();
                zone.classList.remove('dragover');
            });
        });
        
        // Form submission
        uploadForm.addEventListener('submit', () => {
            loadingSection.classList.add('show');
            processBtn.disabled = true;
        });
    </script>
</body>
</html>
"""


@app.route('/')
def index():
    return render_template_string_custom(HTML_TEMPLATE, 
                                  resultados=None, 
                                  current_year=datetime.now().year)


def render_template_string_custom(template, **context):
    """Renderiza el template HTML con contexto usando Flask's render_template_string."""
    from flask import render_template_string as flask_render
    return flask_render(template, **context)


@app.route('/procesar', methods=['POST'])
def procesar():
    try:
        # Verificar archivos
        if 'maestra' not in request.files or 'compilado' not in request.files:
            flash('Debes seleccionar ambos archivos', 'error')
            return redirect(url_for('index'))
        
        maestra_file = request.files['maestra']
        compilado_file = request.files['compilado']
        
        if maestra_file.filename == '' or compilado_file.filename == '':
            flash('Debes seleccionar ambos archivos', 'error')
            return redirect(url_for('index'))
        
        if not (allowed_file(maestra_file.filename) and allowed_file(compilado_file.filename)):
            flash('Solo se permiten archivos Excel (.xlsx, .xls)', 'error')
            return redirect(url_for('index'))
        
        # Guardar archivos
        fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        maestra_filename = secure_filename(f"maestra_{fecha}.xlsx")
        compilado_filename = secure_filename(f"compilado_{fecha}.xlsx")
        
        maestra_path = UPLOAD_FOLDER / maestra_filename
        compilado_path = UPLOAD_FOLDER / compilado_filename
        
        maestra_file.save(str(maestra_path))
        compilado_file.save(str(compilado_path))
        
        # Procesar archivos
        print(f"\n{'='*60}")
        print("PROCESANDO DESDE INTERFAZ WEB")
        print(f"{'='*60}")
        
        # 1. Cargar datos
        df_maestra, df_compilado = cargar_datos(str(maestra_path), str(compilado_path))
        df_maestra_original = df_maestra.copy()  # Guardar original para PDF
        total_compilado = len(df_compilado)
        
        # 2. Matching exacto
        coincidencias_exactas, sin_coincidencia = matching_exacto(df_maestra, df_compilado)
        
        # 3. Matching relativo
        coincidencias_relativas, ambiguos, sin_match = matching_relativo(df_maestra, sin_coincidencia)
        
        # 4. Combinar coincidencias
        todas_coincidencias = coincidencias_exactas + coincidencias_relativas
        
        # 5. Aplicar cambios
        df_maestra_actualizada, log_cambios = aplicar_cambios(df_maestra, todas_coincidencias)
        
        # 6. Calcular KPIs
        filas_unicas = len(set(c['center_code'] for c in todas_coincidencias))
        kpis = calcular_kpis(
            total_compilado=total_compilado,
            coincidencias_exactas=len(coincidencias_exactas),
            coincidencias_relativas=len(coincidencias_relativas),
            filas_actualizadas=filas_unicas
        )
        
        # 7. Generar archivos de salida
        archivo_maestra = f"Maestra_ACTUALIZADA_{fecha}.xlsx"
        archivo_reporte = f"Reporte_{fecha}.xlsx"
        archivo_pdf = f"Comparacion_{fecha}.pdf"
        
        ruta_maestra_out = OUTPUT_FOLDER / archivo_maestra
        ruta_reporte_out = OUTPUT_FOLDER / archivo_reporte
        ruta_pdf_out = OUTPUT_FOLDER / archivo_pdf
        
        # Guardar maestra actualizada
        guardar_maestra_actualizada(df_maestra_actualizada, str(ruta_maestra_out))
        
        # Generar reporte Excel
        generar_reporte(kpis, log_cambios, ambiguos, sin_match, str(ruta_reporte_out))
        
        # Generar PDF comparativo
        generar_pdf_comparacion(
            todas_coincidencias,
            df_maestra_original,
            str(ruta_pdf_out),
            titulo="Comparación Maestra vs Compilado"
        )
        
        # Preparar resultados para mostrar
        resultados = {
            'pct_matching_exacto': f"{kpis['pct_matching_exacto']:.1f}",
            'pct_matching_relativo': f"{kpis['pct_matching_relativo']:.1f}",
            'pct_filas_actualizadas': f"{kpis['pct_filas_actualizadas']:.1f}",
            'total_cambios': len(log_cambios),
            'archivo_maestra': archivo_maestra,
            'archivo_reporte': archivo_reporte,
            'archivo_pdf': archivo_pdf
        }
        
        flash('Procesamiento completado exitosamente', 'success')
        return render_template_string_custom(HTML_TEMPLATE, 
                                      resultados=resultados, 
                                      current_year=datetime.now().year)
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        flash(f'Error al procesar: {str(e)}', 'error')
        return redirect(url_for('index'))


@app.route('/download/<filename>')
def download(filename):
    """Descargar archivo generado."""
    from flask import make_response
    file_path = OUTPUT_FOLDER / filename
    if file_path.exists():
        response = make_response(send_file(str(file_path), as_attachment=True))
        # Headers anti-caché para forzar descarga
        response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'
        return response
    else:
        flash('Archivo no encontrado', 'error')
        return redirect(url_for('index'))


if __name__ == '__main__':
    print("\n" + "=" * 60)
    print("🚀 INICIANDO INTERFAZ WEB")
    print("=" * 60)
    print("\n📌 Abre tu navegador en: http://localhost:5001")
    print("   Presiona Ctrl+C para detener el servidor\n")
    
    app.run(debug=True, host='0.0.0.0', port=5001)
