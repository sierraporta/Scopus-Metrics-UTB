# UTB Scopus Dashboard (>= 2022)

Este repositorio contiene un tablero interactivo (HTML) y tablas (Excel) para explorar la producción científica asociada a docentes de planta de la
Universidad Tecnológica de Bolívar (UTB) a partir de un export de Scopus.

> **Actualización:** estas estadísticas y gráficos están actualizados a fecha de **Mar 1-2026**.

## Metodología (resumen)

- **Fuente de datos:** export CSV desde Scopus (EID, año, tipo de documento, autores).
- **Vinculación a planta UTB:** cruce por **Scopus Author ID** contra una base maestra interna de docentes de planta.
- **Unidad de conteo:** documentos únicos por **EID** (evita dobles conteos por múltiples apariciones del mismo autor).
- **Crédito por Escuela:** una Escuela recibe crédito si al menos un docente de esa Escuela aparece como autor en el documento (un documento puede contar en más de una Escuela si hay coautoría inter-escuelas).
- **Tipos de documento:** agrupación operativa en *Article*, *Conference*, *Review* y *Other* según “Document Type”.

## Abrir el tablero

Abre el archivo principal:

- **Página principal (índice):** utb_scopus_dashboard_single_pretty/index.html

## Gráficos y estadísticas

- **Documentos por año (apilado por tipo):** `utb_scopus_dashboard_single_pretty/overall_papers_per_year_STACKED_by_type_horizontal.html`
- **Documentos por Escuela (apilado + dropdown por año):** `utb_scopus_dashboard_single_pretty/papers_by_escuela_STACKED_by_type_YEAR_dropdown.html`
- **Top autores (apilado + dropdown por año):** `utb_scopus_dashboard_single_pretty/top_authors_STACKED_by_type_YEAR_dropdown.html`
- **Heatmap Escuela × Tipo (con totales + dropdown por año):** `utb_scopus_dashboard_single_pretty/heatmap_escuela_doctype_YEAR_dropdown.html`
- **Top pares planta–planta (dropdown por año):** `utb_scopus_dashboard_single_pretty/top_pairs_faculty_faculty_YEAR_dropdown.html`

## Tablas (Excel)

- **Descargar tablas:** `utb_scopus_dashboard_single_pretty/tables.xlsx`

## Aclaratoria

Este tablero es un ejercicio técnico y personal de análisis bibliométrico basado en un export puntual de Scopus y una lista interna de docentes de planta.
Los resultados son **referenciales** y pueden diferir de cifras institucionales oficiales (cobertura del export, actualización de perfiles, homónimos/duplicados de Author ID, reglas de conteo).
Este material no constituye un reporte oficial ni representa una posición institucional de la UTB.

## Créditos
Desarrollado por **D. Sierra-Porta** © 2026 — Universidad Tecnológica de Bolívar
