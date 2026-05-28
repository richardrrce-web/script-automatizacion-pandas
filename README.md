# Script de Automatización y Reportes con Python y Pandas

Este proyecto simula una solución real de nivel freelance. Es un script de automatización en Python que toma un volumen de datos desordenados (en formato CSV), los limpia, filtra bajo condiciones de negocio específicas y genera un reporte ejecutivo en Excel de forma automática.

## 🚀 Características
- **Procesamiento de datos:** Uso de la librería **Pandas** para la manipulación y filtrado eficiente de estructuras de datos.
- **Motor de Excel:** Integración con **Openpyxl** para la exportación limpia a hojas de cálculo `.xlsx`.
- **Lógica de negocio:** Filtra automáticamente registros completados con montos VIP (>= $100) e inyecta metadatos como la fecha exacta del reporte.

## 🛠️ Requisitos e Instalación
1. Clona este repositorio.
2. Instala las librerías necesarias:
   ```bash
   python -m pip install pandas openpyxl