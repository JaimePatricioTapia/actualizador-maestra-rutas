# Actualizador Maestra de Rutas 📊

Sistema de sincronización y comparación de planillas para actualizar la Maestra de Rutas.

## Características

- 📁 Carga de archivos Excel (Maestra y Compilado)
- 🔍 Matching exacto y relativo de registros
- 📈 Visualización de KPIs en tiempo real
- 📄 Generación de PDF comparativo
- 📊 Reportes Excel detallados

## Cómo usar

1. Visita la aplicación en Streamlit Cloud
2. Carga el archivo **Maestra de Rutas** (Excel)
3. Carga el archivo **Compilado** (Excel)
4. Haz clic en **Procesar Archivos**
5. Descarga los resultados:
   - Maestra Actualizada (Excel)
   - Reporte de KPIs (Excel)
   - Comparación Visual (PDF)

## Instalación local

```bash
pip install -r requirements.txt
streamlit run app_streamlit.py
```

## Tecnologías

- Python 3.9+
- Streamlit
- Pandas
- ReportLab (PDF)

---
Desarrollado para Castaño - Tienda Perfecta 🏪
