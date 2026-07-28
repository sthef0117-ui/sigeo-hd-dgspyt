# SIGEO-HD DGSPYT
### Sistema de Geointeligencia Operativa HD — DGSPYT & C5

Plataforma de geointeligencia operativa y analítica táctica de homicidios dolosos
para la Dirección General de Seguridad Pública y Tránsito (Unidad de Homicidios
Dolosos) en coordinación con el centro de mando C5.

Todo lo que muestra el tablero se **calcula** a partir de los insumos oficiales.
No hay cifras ni casos escritos a mano: cualquier número del informe se reproduce
ejecutando el pipeline.

---

## Puesta en marcha

```bash
pip install openpyxl
python src/etl_sigeo.py
python src/build_dashboard.py
```

El primer comando lee los Excel de `insumos/`, calcula los cuatro módulos y
escribe `analisis/*.json` más las tablas de `database/sigeo_db.sqlite`.
El segundo ensambla `index.html`, un archivo único y autocontenido que se abre
desde GitHub Pages o desde una memoria USB en la sala de juntas, sin servidor.

---

## Módulos

### 1. Mapa geoespacial
Homicidios corroborados, llamadas 911 con violencia, inventario de bases DGSPYT,
mapa de calor y sectores de zona ciega, sobre el mismo esquema de capas base que
Windows Maps: **Carretera**, **Vista aérea** e **Híbrido**. Al seleccionar un
evento, la ficha muestra hechos, acciones SSEM y la cobertura de patrullaje real
del punto (base más cercana, distancia y personal adscrito).

### 2. Extractor NLP de ubicaciones — formato Windows Maps
Resuelve el problema planteado en la reunión de mandos: *«la tabla carece de
datos suficientes, algunas traen descripción — encontrar la forma de que un
sistema lo lea»*.

Lee los tres campos que la cabina sí captura (`DIRECCIÓN` en formato
`;CALLE|COLONIA`, `REFERENCIA DE UBICACIÓN` y `NOTAS` en lenguaje natural),
extrae vialidad, número, cruce, entre-calles, asentamiento y punto de referencia,
y arma la cadena que Windows Maps resuelve, más los URI nativos `bingmaps:`
(abrir en la app) y `ms-drive-to:` (despachar unidad).

Incluye una consola en vivo para el operador de cabina.

**Resultado sobre el corte actual:** de las 4,219 llamadas sin coordenada,
3,756 son mudas, colgadas, de broma o transferidas y no contienen descripción
alguna. El universo recuperable es de **463**; el extractor resuelve **136**
(29.4%) con confianza alta o media. En el subconjunto que importa —llamadas con
violencia sin coordenada— la recuperación es de **25 de 27 (92.6%)**.

El extractor no inventa ubicaciones: sin señal suficiente marca confianza
`NULA` y no emite consulta.

Implementación: `src/nlp_geocoder.py` (lote) y su puerto JavaScript dentro del
tablero (vivo). Ambos aplican las mismas reglas y niveles de confianza.

### 3. Detector de zonas ciegas de patrullaje
*«¿Por qué no se patrulla ahí?»* convertido en métrica reproducible.

El territorio se divide en celdas de ~2.2 km. En cada una se suma la violencia
registrada (homicidio corroborado ×10 y cada llamada 911 según su gravedad), se
mide la distancia a la base DGSPYT **en uso** más cercana del inventario de
inmuebles y se contabiliza el personal adscrito a 3 km:

```
índice_ceguera = violencia × (1 + min(dist_base, 6)/2) × 1/(1 + personal_3km/150)
```

Cada sector recibe un diagnóstico, y cada diagnóstico exige una decisión distinta:

| Diagnóstico | Situación | Decisión |
|---|---|---|
| `DESATENDIDO` | Violencia sin despliegue a distancia útil | Extender cobertura o reasignar cuadrante |
| `SATURADO` | Base y personal a menos de 1.5 km y aun así concentra violencia | Revisar efectividad y horarios del patrullaje, no sumar inmuebles |
| `MIXTO` | Cobertura intermedia | Verificar recorridos y tiempos de respuesta |

La tabla se ordena por brecha de cobertura o por concentración de violencia, y
cada sector expone su evidencia: incidentes registrados, homicidios del sector y
si las acciones SSEM asientan ausencia de cámaras de videovigilancia.

### 4. Auditoría algorítmica de decesos dudosos, suicidios y no localizados
Conforme al apunte *«Suicidio → 25… + identificar, desaparecidos, homicidios
dolosos. Investigar si son suicidios»*.

La cohorte son los **30 reportes reales** del corte C5 clasificados como
suicidio, tentativa o amenaza de suicidio, persona tirada en vía pública —con y
sin huellas de violencia—, persona no localizada y homicidio.

Cada caso se puntúa con indicadores explícitos sobre el texto de cabina y la
geometría del hecho: mención de arma de fuego, indicios de ocultamiento,
participación de terceros, antecedente de violencia de pareja, proximidad
(≤1.5 km) a un homicidio corroborado, proximidad espacio-temporal a otro reporte
por arma de fuego, y también indicadores de descarte (persona localizada,
atención médica con persona consciente, estado etílico).

> **Límite del módulo.** El resultado es una **prioridad de revisión documental**,
> no un dictamen pericial. No afirma la mecánica de la muerte: esa determinación
> corresponde a los servicios periciales y a la FGJEM. Cada caso muestra los
> indicadores exactos que elevaron o bajaron su prioridad, para que el mando
> audite el criterio y no la conclusión.

### 5. Reporte ejecutivo para la reunión de mandos C5
Ficha imprimible (hoja de estilo de impresión incluida) con incidencia por
municipio, cobertura de patrullaje, resultados de geocodificación, casos de
prioridad alta y acuerdos propuestos. Todas las cifras provienen del pipeline.

---

## Protección de datos personales

El tablero se publica en Internet abierto. Las notas de cabina y el desarrollo
de hechos contienen datos personales y sensibles: nombres de víctimas y de
personas desaparecidas, edad, estatura, tatuajes, vestimenta, teléfonos y claves
de operador.

`src/anonimizar.py` los suprime **antes** de escribir los JSON publicables.
La base SQLite local conserva el texto íntegro para el trabajo operativo interno.
El pipeline nunca exporta nombre, apellido ni teléfono del reportante 911: esas
columnas se descartan al leer el Excel.

Para generar los JSON sin suprimir —solo para uso interno, nunca para publicar—:

```bash
python src/etl_sigeo.py --sin-anonimizar
```

---

## Estructura

```
src/
  etl_sigeo.py            Pipeline: Excel -> JSON + SQLite. Cálculo de los 4 módulos.
  nlp_geocoder.py         Extractor NLP de ubicaciones en formato Windows Maps.
  anonimizar.py           Supresión de datos personales para publicación.
  build_dashboard.py      Ensamblador del tablero de archivo único.
  dashboard_template.html Plantilla del tablero (incluye el puerto JS del extractor).
analisis/                 JSON calculados (publicables, anonimizados).
database/sigeo_db.sqlite  Base relacional: HD, llamadas C5, bases, zonas ciegas,
                          auditoría y geocodificación.
insumos/                  Excel oficiales de DGSPYT y C5, sin modificar.
index.html                Tablero generado. No editar a mano: se regenera.
```

## Insumos

| Archivo | Contenido |
|---|---|
| `ACCIONES DE HD JULIO DGSPYT.xlsx` | 56 homicidios corroborados (julio 2026) y 90 registros FGR |
| `LLAMADAS_911_CORTE...xlsx` | 7,414 llamadas C5 del 26-27 de julio de 2026 |
| `INMUEBLES GENERAL DGSPYT.xlsx` | 198 inmuebles DGSPYT; 167 en uso y georreferenciados, 11,147 elementos adscritos |
