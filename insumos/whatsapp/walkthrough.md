# Walkthrough - Análisis Geoespacial de Incidencia Delictiva (911) vs. Inmuebles de Seguridad (DGSPyT)

Se completó con éxito el procesamiento geoespacial, cálculo de distancias ortodrómicas (Haversine), análisis de cobertura y generación de reportes e iteraciones interactivas para contrastar la incidencia delictiva del 911 (`Base 911.xlsx`) contra los inmuebles/bases de la Dirección General de Seguridad Pública y Tránsito (`INMUEBLES GENERAL DGSPYT.xlsx`).

## Archivos Generados

1. **Reporte Multipe pestaña en Excel**:
   - [Reporte_Distancias_911_vs_Bases.xlsx](file:///C:/Users/xXZaB/Documents/C5%20INTELIGENCIA/Reporte_Distancias_911_vs_Bases.xlsx)
2. **Mapa Interactivo HTML (Leaflet / Folium)**:
   - [mapa_interactivo_911_vs_bases.html](file:///C:/Users/xXZaB/Documents/C5%20INTELIGENCIA/mapa_interactivo_911_vs_bases.html)
3. **Script de Procesamiento en Python**:
   - [procesar_distancias_911.py](file:///C:/Users/xXZaB/.gemini/antigravity/scratch/procesar_distancias_911.py)

---

## Principales Resultados y Métricas KPI

| Métrica | Valor |
| :--- | :--- |
| **Total de Folios 911 Procesados** | 1,271 |
| **Incidencias con Coordenadas Válidas** | 1,154 (90.8%) |
| **Incidencias sin Geolocalización (0,0)** | 117 (9.2%) |
| **Bases e Inmuebles DGSPyT Mapeados** | 198 |
| **Distancia Promedio a la Base Más Cercana** | **2.148 km** (2,148 m) |
| **Mediana de Distancia** | **1.621 km** (1,621 m) |
| **Distancia Mínima** | **0.003 km** (3 m) |
| **Distancia Máxima** | **18.729 km** (18,729 m) |
| **Percentil 25% (P25)** | **0.902 km** (902 m) |
| **Percentil 75% (P75)** | **2.880 km** (2.88 km) |
| **Percentil 90% (P90)** | **4.635 km** (4.64 km) |

---

## Distribución por Rango de Cobertura Espacial

- 🟢 **< 500 m (Respuesta Inmediata)**: 6.8% (78 incidencias)
- 🟢 **500 m - 1.5 km (Cobertura Cercana)**: 37.0% (427 incidencias) -> *Total < 1.5km: 43.8%*
- 🟡 **1.5 km - 3.0 km (Cobertura Media)**: 33.8% (391 incidencias) -> *Total < 3.0km: 77.6%*
- 🟠 **3.0 km - 5.0 km (Cobertura Distante)**: 13.2% (152 incidencias) -> *Total < 5.0km: 90.8%*
- 🔴 **> 5.0 km (Fuera de Cobertura Primaria)**: 9.2% (106 incidencias)

---

## Top 5 Municipios con Mayor Volumen de Incidencias

1. **Toluca**: 139 incidencias | Distancia promedio a base: **2.30 km** | Mediana: **1.70 km**
2. **Ecatepec de Morelos**: 115 incidencias | Distancia promedio a base: **1.15 km** | Mediana: **0.92 km**
3. **Naucalpan de Juárez**: 72 incidencias | Distancia promedio a base: **2.12 km** | Mediana: **1.31 km**
4. **Cuautitlán Izcalli**: 65 incidencias | Distancia promedio a base: **1.54 km** | Mediana: **1.49 km**
5. **Nezahualcóyotl**: 63 incidencias | Distancia promedio a base: **1.32 km** | Mediana: **1.22 km**

---

## Top 5 Inmuebles con Mayor Concentración de Delitos Cercanos (< 3 km)

1. **Base #168 (Toluca - Dir. Policía de Tránsito)**: 56 incidencias a <3km (10 a <1km)
2. **Base #173 (Toluca - Base de Seguridad)**: 53 incidencias a <3km (7 a <1km)
3. **Base #52 (Ecatepec - Tercer Agrupamiento Américas)**: 51 incidencias a <3km (35 a <1km)
4. **Base #170 (Toluca - Cuarto Agrupamiento)**: 51 incidencias a <3km (8 a <1km)
5. **Base #58 (Ecatepec - Comandancia Subdir. Operativa Circuito Exterior Mexiquense)**: 51 incidencias a <3km

---

## Verificación de Integridad

- Se validó que el 100% de las 1,154 distancias calculadas sean no negativas y tengan sentido espacial en el marco geográfico del Estado de México.
- El libro Excel contiene las 5 pestañas completas y formateadas.
- El mapa HTML incluye la capa de bases, capa de incidentes codificados por color por proximidad, mapa de calor (Heatmap) y caja de leyenda flotante.
