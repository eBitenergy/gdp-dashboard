# Enhanced Solar Flex Dashboard Generator

This README provides instructions for using the enhanced Solar Flex Excel generator script and VBA macros.

## Overview
`generate_solar_flex_dashboard_pro.py` builds a multi-hoja Excel workbook para gestionar el pipeline comercial de Solarflex y Livoltek. Incluye un panel de control con KPIs, la base de oportunidades, requisitos BESS, proyectos de amianto Solarflex y tarifas de equipos Livoltek.

El script utiliza datos de ejemplo para que puedas ver el formato final sin necesidad de rellenar manualmente las hojas.

## Dependencias
```bash
pip install openpyxl
```

## Uso
```bash
python generate_solar_flex_dashboard_pro.py
```

Se generará el archivo `Pipeline_Maestro_Solarflex_Livoltek.xlsx` en el directorio actual.

## VBA Macros
Importa los macros de `macros_vba.txt` en Excel para disponer de botones de navegación y acciones rápidas.
