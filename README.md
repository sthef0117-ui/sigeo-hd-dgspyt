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

El primer comando lee los Excel de `insumos/`, calcula los indicadores y escribe
`analisis/*.json` más las tablas de `database/sigeo_db.sqlite`.

**Para agregar un corte nuevo del C5 basta copiarlo a `insumos/excel/` y volver a
ejecutar.** El pipeline descubre solo todos los archivos de llamadas, los une y
deduplica por folio: los cortes contiguos repiten folios y sin deduplicar se
contarían dos veces.
El segundo ensambla `index.html`, un archivo único y autocontenido que se abre
desde GitHub Pages o desde una memoria USB en la sala de juntas, sin servidor.

---

## Organización

El tablero se organiza por **decisión**, no por módulo técnico. Seis vistas en tres
bloques, en el orden en que se usan en una reunión de mandos.

### Decidir

**01 · Panorama** — la portada del Director. Abre con la curva de demanda de
emergencia hora por hora de todos los cortes integrados, cuatro cifras del periodo
y, debajo,
*qué atender en esta reunión*: los sectores donde falta cobertura, los sectores
donde ya hay despliegue y aun así hay violencia, y los expedientes que conviene
pedir a la FGJEM. Cada bloque trae el acuerdo sugerido.

**02 · Territorio** — mapa y detector de zonas ciegas en una sola pantalla, porque
son la misma pregunta.

El mapa es **cartografía propia**: se dibujan los 121 polígonos municipales del
Estado de México en lugar de pedir mosaicos a un proveedor. Eso evita tramitar
llave de suscripción, elimina el crédito de terceros impreso sobre el mapa y hace
que se parezca al mural de la Dirección. El municipio se colorea por
**coordinación regional**, por **número de homicidios** o sin relleno.

Capas conmutables (homicidios, llamadas, bases, calor, sectores), perímetro del
estado trazado en toda su orilla, botón de **estado completo** y una lista de
sectores a la derecha; al hacer clic el mapa vuela al sector y abre su evidencia.

**Detalle a nivel calle.** Al abrir la ficha de un hecho el mapa se acerca al
punto y enciende solo la capa de calles: una dirección no sirve si no se ve dónde
cae. Mientras esa capa está apagada el mapa no depende de nadie y no aparece
crédito de proveedor; al encenderla aparece, porque entonces sí hay a quién
acreditar. También se puede prender a mano, en calles o satélite.

> **Fuente de los límites municipales.** Se derivan de un conjunto público de
> límites administrativos y traen tres municipios de creación reciente sin
> polígono: San José del Rincón, Tonanitla y Valle de Chalco. Para uso
> institucional conviene sustituirlos por el **Marco Geoestadístico del INEGI**,
> que es la fuente oficial y de uso libre. `src/preparar_poligonos.py` acepta
> cualquier GeoJSON de municipios: simplifica y normaliza los nombres solo.

*«¿Por qué no se patrulla ahí?»* se calcula así: el territorio se divide en celdas
de ~2.2 km; en cada una se suma la violencia registrada (homicidio corroborado ×10
y cada llamada 911 según su gravedad), se mide la distancia a la base DGSPYT **en
uso** más cercana y se cuenta el personal adscrito a 3 km.

```
índice_ceguera = violencia × (1 + mín(dist_base, 6)/2) × 1/(1 + personal_3km/150)
```

| Diagnóstico | Situación | Decisión |
|---|---|---|
| `DESATENDIDO` | Violencia sin despliegue a distancia útil | Extender cobertura o reasignar cuadrante |
| `SATURADO` | Base y personal a menos de 1.5 km y aun así concentra violencia | Revisar efectividad y horarios; sumar inmuebles no resuelve |
| `MIXTO` | Cobertura intermedia | Verificar recorridos y tiempos de respuesta |

**03 · Coordinaciones** — la herramienta para la mesa con coordinadores
territoriales. Ficha por municipio con homicidios, llamadas violentas, bases,
personal, distancia media al despliegue y sectores con brecha, ordenada por índice
de presión. Al hacer clic en la fila se despliega cómo se compone ese número.

```
índice_presión = (HD×10 + llamadas violentas)
                 × (1 + mín(dist media, 6)/4)
                 × (1 + sectores desatendidos × 0.25)
```

**Clasificación de los municipios.** El inventario es de inmuebles, no de
territorio, así que 22 municipios sin inmueble propio no traían coordinación. Se
completan en dos pasos y cada municipio guarda de dónde salió su asignación:
`catalogo` si viene declarada, `escision` si se hereda del municipio del que se
separó, y `vecindad` si se toma la coordinación mayoritaria entre sus colindantes,
calculada sobre los polígonos. Los 108 municipios quedan clasificados: 86 por
catálogo y 22 por vecindad.

> **Qué mide y qué no.** Mide condiciones del territorio. **No mide el desempeño de
> una persona:** los insumos no contienen asignación nominal de mando, turnos ni
> recorridos, así que atribuir un resultado a un coordinador específico no se
> sostiene con estos datos. Sirve para abrir la conversación con evidencia, no para
> calificar gente.

### Investigar

**04 · Casos a revisar** — auditoría de decesos dudosos, suicidios y personas no
localizadas, conforme al apunte *«Suicidio → 25… + identificar, desaparecidos,
homicidios dolosos. Investigar si son suicidios»*.

La cohorte son los **247 reportes reales** de los cortes C5 clasificados como suicidio,
tentativa o amenaza de suicidio, persona tirada en vía pública —con y sin huellas
de violencia—, persona no localizada y homicidio. Cada caso se puntúa con
indicadores explícitos sobre el texto de cabina y la geometría del hecho: mención
de arma de fuego, indicios de ocultamiento, participación de terceros, antecedente
de violencia de pareja, proximidad (≤1.5 km) a un homicidio corroborado,
proximidad espacio-temporal a otro reporte por arma de fuego, y también
indicadores de descarte (persona localizada, atención médica con persona
consciente, estado etílico).

> **Límite del módulo.** El resultado es una **prioridad de revisión documental**,
> no un dictamen pericial. No afirma la mecánica de la muerte: esa determinación
> corresponde a los servicios periciales y a la FGJEM. Cada caso muestra los
> indicadores exactos que elevaron o bajaron su prioridad, para que el mando
> audite el criterio y no la conclusión.

### Operar

**05 · Ubicaciones** — extractor NLP en formato Windows Maps. Resuelve el problema
planteado en la reunión: *«la tabla carece de datos suficientes, algunas traen
descripción — encontrar la forma de que un sistema lo lea»*.

Lee los tres campos que la cabina sí captura (`DIRECCIÓN` en formato
`;CALLE|COLONIA`, `REFERENCIA DE UBICACIÓN` y `NOTAS` en lenguaje natural), extrae
vialidad, número, cruce, entre-calles, asentamiento y punto de referencia, y arma
la cadena que Windows Maps resuelve, más los URI nativos `bingmaps:` (abrir en la
app) y `ms-drive-to:` (despachar unidad). Incluye consola en vivo para el operador.

**Resultado sobre los cortes integrados:** la mayoría de las llamadas sin
coordenada son mudas, colgadas, de broma o transferidas y no contienen descripción
alguna. El universo recuperable es de **2,590**; el extractor resuelve **970**
(37.5%) con confianza alta o media. En el subconjunto que importa —llamadas con
violencia sin coordenada— la recuperación es de **182 de 189 (96.3%)**.

El extractor no inventa ubicaciones: sin señal suficiente marca confianza `NULA` y
no emite consulta. Implementación en `src/nlp_geocoder.py` (lote) y su puerto
JavaScript dentro del tablero (vivo); ambos aplican las mismas reglas.

**06 · Informe de mandos** — hoja ejecutiva imprimible con hoja de estilo de
impresión propia: presión territorial, cobertura y zonas ciegas, geocodificación,
casos de prioridad alta y acuerdos propuestos. Todas las cifras provienen del
pipeline.

---

## Estructura territorial oficial

La Dirección es estatal y el tablero cubre el Estado de México completo. El mapa
abre con el **perímetro del estado trazado en toda su orilla** y la agregación
usa la división real de mando, tomada de `INSTALACIONES_DGSPYT`: **8
coordinaciones regionales** territoriales (Metropolitana, Ecatepec, Valle Toluca,
Oriente, Chalco, Atlacomulco, Ixtapan, Valle de Bravo), 22 subdirecciones y 95
municipios asignados. Es la misma división del mapa mural de la Dirección.

Los municipios sin inmueble en el inventario quedan agrupados como «sin
coordinación asignada» en lugar de asignarles una a la fuerza.

---

## Power BI

`python src/export_powerbi.py` genera en `powerbi/` un modelo en estrella en CSV
(3 dimensiones y 4 tablas de hechos) listo para cargar en Power BI Desktop.
`powerbi/MODELO.md` trae las relaciones, las medidas DAX y el armado de las
cuatro páginas.

El visual de mapa de Power BI usa **Bing Maps de forma nativa**, sin tramitar
llave. El tablero HTML resuelve lo mismo por otra vía: dibuja su propia
cartografía vectorial y no depende de ningún proveedor de mosaicos.

---

## Diseño

Interfaz institucional sobria, no de maqueta: navegación lateral por decisión,
tipografía con numeración tabular, tablas densas y paleta guinda institucional.
Modo claro por defecto —lee mejor en sala y en impresión— con modo oscuro para el
videowall del C5.

---

## Protección de datos personales

El acceso al tablero es interno, por enlace. Aun así conviene mantener la
supresión activa: un repositorio público en GitHub Pages es indexable por
buscadores aunque nadie comparta el enlace, y las notas de cabina y el desarrollo
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
  etl_sigeo.py            Pipeline: Excel -> JSON + SQLite. Todos los cálculos.
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
| `CONCENTRADO_HISTORICO_LLAMADAS_911_089_C5.xlsx` + cortes diarios | 22,811 llamadas únicas del C5, del 22 al 28 de julio de 2026 (120 h) |
| `INSTALACIONES_DGSPYT_290426_QR L.xlsx` | 198 inmuebles con coordinación regional, subdirección, coordenadas y personal |
| `poligonos/perimetro_edomex.geojson` | Perímetro del Estado de México (201 vértices) |
| `INMUEBLES GENERAL DGSPYT.xlsx` | Inventario anterior, sin coordinación regional. Se usa solo como respaldo |
