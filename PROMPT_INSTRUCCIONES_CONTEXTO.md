# PROMPT & CONTEXTO OPERATIVO OFICIAL: SIGEO-HD DGSPYT

Este documento contiene el **Prompt de Contexto Institucional y Requerimientos Técnicos Completos** para desarrollar la plataforma sobre la carpeta del proyecto `sigeo_hd_dgspyt`.

---

## 📌 1. CONTEXTO INSTITUCIONAL & IDENTIDAD

- **Entidad Operativa**: Dirección General de Seguridad Pública y Tránsito (**DGSPYT** - Unidad de Homicidios Dolosos **HD**), en coordinación técnica y operativa con el centro de mando **C5**.
- **Acrónimo del Proyecto**: **SIGEO-HD DGSPYT** *(Sistema de Geointeligencia Operativa HD - DGSPYT)*.
- **Ruta de Trabajo del Proyecto**: `C:\Users\xXZaB\.gemini\antigravity\scratch\sigeo_hd_dgspyt`
- **Enlace de Despliegue en Línea (GitHub Pages)**: `https://sthef0117-ui.github.io/sigeo-hd-dgspyt/`
- **Repositorio de Código**: `https://github.com/sthef0117-ui/sigeo-hd-dgspyt`

---

## 📄 2. TRANSCRIPCIÓN FIEL DE LOS APUNTES MANUSCRITOS (REUNIÓN DE MANDOS C5)

### Imagen 1 (Directrices de Mando y Patrullaje):
- **Eje Central**: Homicidios (`>> Homicidios <<`), alrededor de un deceso. Foco de atención prioritario en Homicidios Dolosos (HD).
- **Herramientas al Director**: Desarrollar aplicativos y visualizaciones nítidas ("Nitidez") para el Director de la DGSPYT.
- **Detector de Zonas Ciegas de Patrullaje**: Analizar delitos conectados y sectores específicos cuestionando *"¿Por qué no se patrulla ahí?"*.
- **Gestión Operativa**: Reuniones de advertencia con coordinadores, presionar a coordinadores territoriales, perfilado de desempeño y bosquejos de presentación continua.

### Imagen 2 (Sistemas, Datos y Algoritmos):
- **Auditoría Forense Tipológica (Suicidios vs HD)**:
  - *`Suicidio -> 25... + Identificar, desaparecidos, Homicidios dolosos. Investigar si son suicidios.`*
  - Detectar e identificar patrones sutiles ("no tan violentos") o muertes catalogadas como suicidio que puedan encubrir homicidios dolosos.
- **Factores Geográficos y Limítrofes**:
  - Mapeo de geografía, orografía y fronteras intermunicipales (*"Cubrir / fuera de territorio"*).
- **Geocodificación Windows Maps NLP**:
  - *Problemática*: *"Puntearlos, la tabla carece de datos suficientes. Algunas traen descripción -> Encontrar la forma de que un sistema lo lea."*
  - *Solución Requerida*: `[Maps de windows (Open maps)] herramienta a usar - porque lee ubicaciones en formato de descripción.`
- **Integración Algorítmica C5**:
  - Encontrar la manera de que los decesos se integren geográficamente con la subbanca del C5 y las llamadas de emergencia al 911.

---

## 📁 3. INSUMOS Y ESTRUCTURA DE LA CARPETA

```
sigeo_hd_dgspyt/
├── PROMPT_INSTRUCCIONES_CONTEXTO.md            (Este documento de contexto)
├── index.html                                   (Plataforma Táctica Operativa)
├── database/
│   └── sigeo_db.sqlite                          (Base de datos relacional SQLite)
├── insumos/
│   ├── excel/
│   │   ├── ACCIONES DE HD JULIO DGSPYT.xlsx      (56 casos corroborados HD julio 2026, 90 FGR)
│   │   └── LLAMADAS_911_CORTE...xlsx              (7,418 llamadas 911 C5 del 26-27 de julio)
│   └── whatsapp/
│       ├── Chat de WhatsApp con C5.txt           (295 registros de seguimiento)
│       ├── INFORME_911_VIOLENCIA_21-22_JUL_2026.docx
│       └── mapa_interactivo_911_vs_bases.html
├── analisis/
│   ├── corroborados_sigeo.json                   (56 eventos HD estructurados)
│   └── llamadas_911_sigeo.json                    (758 llamadas de prioridad C5)
└── src/
    ├── sigeo_schema.sql                          (Esquema DDL SQL)
    ├── sigeo_core.cpp                            (Motor Nativo C++ Haversine)
    └── sigeo_server.py                           (Servidor REST API SQL)
```

---

## ⚙️ 4. REQUERIMIENTOS TÉCNICOS & ARQUITECTURA

1. **Geocodificación Windows Maps**: Transformación de descripciones textuales de observaciones en coordenadas geográficas.
2. **Motor SQL SQLite**: Persistencia relacional con vistas analíticas por municipio (`vw_estadisticas_municipales`) y por móvil (`vw_analisis_moviles`).
3. **Motor Nativo C++ (ISO C++17)**: Cálculo Haversine de distancia espacial entre incidentes del 911 y eventos HD corroborados.
4. **Módulos Operativos Tácticos**:
   - Detector de Zonas Ciegas de Patrullaje (*"¿Por qué no se patrulla ahí?"*).
   - Auditor Forense Algorítmico de Decesos/Suicidios (~25 casos).
   - Generador de Ficha Ejecutiva Imprimible / PDF para la Reunión de Mandos C5.
