---
layout: default
title: Scopus Metrics UTB
nav_order: 1
---

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![Jupyter](https://img.shields.io/badge/Jupyter-Notebooks-orange.svg)](https://jupyter.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

**Desarrollado por:** D. Sierra-Porta (UTB)

# Scopus Metrics UTB
Este repositorio contiene un **notebook/algoritmo** que consolida **tablas resumen de productividad científica** de investigadores de la Universidad, a partir de registros bibliométricos y su **clasificación por cuartil** (**Q1, Q2, Q3, Q4**) y/o **Sin Cuartil (SQ)**.

**Última actualización de estadísticas:** **2025-12-29**.

---

## Objetivo
Generar de forma reproducible un conjunto de **indicadores descriptivos** que permitan:

- resumir la **producción anual** por investigador y por unidad académica;
- comparar la distribución de publicaciones por cuartiles (**Q1–Q4**) y **Sin Cuartil (SQ)**;
- producir visualizaciones estandarizadas (gráficas y heatmaps) para reportes internos y análisis exploratorio.

{: .note }
Este repositorio apunta a **resumen estadístico y visualización**, no a evaluación individual de desempeño ni a ranking normativo.

---

## Definiciones usadas
- **Q1–Q4:** cuartil asignado a la revista (según la fuente de cuartiles integrada en el dataset).
- **SQ (Sin Cuartil):** publicaciones sin cuartil asignado o no clasificadas en el esquema Q1–Q4 (incluye casos sin match de cuartil).
- **Total:** suma anual de publicaciones consideradas por el algoritmo, según reglas de filtrado/limpieza.

---

## Flujo metodológico
1. **Ingesta:** lectura del listado de investigadores (UTB), export Scopus y tabla Scimago/SJR.
2. **Normalización:** estandarización de identificadores y metadatos (año, afiliación, etc.).
3. **Clasificación:** asignación de cuartil **Q1–Q4** (match por ISSN y fallback por título); resto → **SQ**.
4. **Agregación:** tablas por investigador, por escuela/unidad académica y por año.
5. **Visualización:** generación automática de figuras y exportables (CSV/Excel).

---

## Resultados 2025

{: .highlight }
**Producción científica UTB (2025)**  
**Total (contribuciones de autoría): 195**  
Q1: **40** · Q2: **29** · Q3: **19** · Q4: **8** · SQ: **99**

{: .note }
**Importante sobre conteos:**  
El “Total (contribuciones de autoría)” cuenta pares **(autor UTB – publicación)**.  
Si necesitas “papers únicos UTB” (cada paper cuenta 1 vez), el total será menor.

---

# Figuras (2025)

## Escuela de Transformación Digital (ETD)

### ETD — Producción por investigador (authorship basis)
Barras apiladas por investigador con desglose **Q1–Q4** y **SQ**. La longitud total representa la producción anual del investigador bajo el conteo **autor–publicación**.

![]({{ "/outputs/figures/ETD_2025_researchers_stackedbar_authorship.png" | relative_url }})

### ETD — Composición por cuartil (authorship basis)
Distribución porcentual de la producción ETD por cuartil (**Q1–Q4**) y **SQ**.

![]({{ "/outputs/figures/ETD_2025_pie_authorship_quartiles.png" | relative_url }})

---

## Toda la Universidad (UTB)

### UTB — Composición por cuartil (authorship basis)
Distribución institucional por cuartil y **SQ** para 2025 bajo conteo **autor–publicación**.

![]({{ "/outputs/figures/UTB_2025_pie_authorship_quartiles.png" | relative_url }})

### UTB — Comparación por escuelas (participación por escuela)
Barras apiladas por **escuela**, contando pares **(escuela – publicación)**. Un paper puede aportar a más de una escuela si hay coautoría inter-escuela.

![]({{ "/outputs/figures/UTB_2025_schools_stackedbar_participation.png" | relative_url }})

### UTB — Heatmap escuelas × cuartil (participación por escuela)
Mapa de calor con conteos por escuela y cuartil, incluyendo la columna **Total**.

![]({{ "/outputs/figures/UTB_2025_heatmap_school_quartile_participation.png" | relative_url }})

### UTB — Top investigadores (authorship basis)
Top 25 investigadores por producción anual (conteo autor–publicación), desagregado por cuartil.

![]({{ "/outputs/figures/UTB_2025_top25_researchers_stackedbar_authorship.png" | relative_url }})

### UTB — Colaboración entre escuelas
Matriz de coautoría inter-escuela: número de publicaciones 2025 que conectan pares de escuelas (papeles con ≥2 escuelas).

![]({{ "/outputs/figures/UTB_2025_collaboration_matrix_schools.png" | relative_url }})

---

## Reproducibilidad (rápido)
- Ejecuta el notebook principal.
- Las figuras se guardan en `outputs/figures/` y luego pueden copiarse a `assets/figures/2025/` para publicarlas en Pages.

---

## Limitaciones / notas
- La exactitud depende de la calidad del export Scopus (IDs, afiliaciones, duplicados).
- **SQ** puede reflejar tanto producción en fuentes sin cuartil como **vacíos de match** (ISSN/título).
- Interpretaciones comparativas deben considerar diferencias disciplinares y cobertura de indexación.
