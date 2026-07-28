# SIGEO-HD en Power BI — guía de armado

## Camino corto: abrir el proyecto ya armado

**Doble clic en `SIGEO-HD.pbip`.** El proyecto trae el modelo completo: las 7
tablas, las 8 relaciones, las 17 medidas DAX y las 4 páginas con sus visuales.
No hay que arrastrar campos ni escribir medidas.

Requisitos: **Power BI Desktop de 2024 en adelante** (el formato de proyecto
`.pbip` no existe en versiones anteriores). Descarga en
<https://powerbi.microsoft.com/desktop/> o desde Microsoft Store.

La ruta de los CSV vive en el parámetro **RutaDatos** (Inicio → Transformar
datos → Administrar parámetros). Si mueven la carpeta de sitio, se cambia ahí y
se actualiza; no hay que tocar cada consulta.

> **Advertencia honesta.** Este proyecto se generó por archivo, sin abrirlo en
> Power BI Desktop, porque la aplicación no está instalada en el equipo donde se
> construyó. El modelo está validado contra los CSV —columnas, relaciones y
> referencias DAX resuelven— pero el acomodo de los visuales puede requerir
> ajustes al abrirlo. Si alguna página sale vacía, el modelo sigue bueno: se
> arrastran los campos y listo. Repórtelo y lo corrijo.

---

## Camino largo: armarlo a mano

Si prefieren construirlo paso a paso, o si su versión de Power BI Desktop no
abre proyectos `.pbip`, los CSV de esta carpeta ya son un modelo en estrella. No
hay que transformarlos: se cargan y se relacionan.

El mapa de Power BI usa **Bing Maps de forma nativa**, sin tramitar ninguna
llave. Eso resuelve el requisito de usar cartografía de Microsoft.

Regenerar todo:

```bash
python src/etl_sigeo.py
python src/export_powerbi.py
python src/build_pbip.py
```

---

## 1. Cargar

**Inicio → Obtener datos → Texto/CSV**, uno por archivo. En la vista previa
confirmar que el **origen de archivo sea `65001: Unicode (UTF-8)`** — si los
acentos salen rotos, los nombres de municipio dejan de coincidir entre tablas y
el modelo se rompe en silencio.

| Archivo | Tipo | Filas |
|---|---|---|
| `dim_calendario.csv` | Dimensión | 28 |
| `dim_municipio.csv` | Dimensión | 108 |
| `dim_coordinacion.csv` | Dimensión | 8 |
| `fact_homicidios.csv` | Hechos | 56 |
| `fact_llamadas.csv` | Hechos | 3,735 |
| `fact_sectores.csv` | Hechos | 363 |
| `fact_auditoria.csv` | Hechos | 247 |

## 2. Relacionar

En **Vista de modelo**, arrastrar. Todas son de **uno a varios**, con el filtro
cruzado en dirección **única**, desde la dimensión hacia los hechos:

```
dim_calendario[fecha]        →  fact_homicidios[fecha]
dim_calendario[fecha]        →  fact_llamadas[fecha]
dim_municipio[municipio]     →  fact_homicidios[municipio]
dim_municipio[municipio]     →  fact_llamadas[municipio]
dim_municipio[municipio]     →  fact_sectores[municipio]
dim_municipio[municipio]     →  fact_auditoria[municipio]
dim_coordinacion[coordinacion] → dim_municipio[coordinacion]
```

Marcar `dim_calendario` como **tabla de fechas** (clic derecho → Marcar como
tabla de fechas → columna `fecha`). Sin eso, la inteligencia de tiempo no
funciona.

Ocultar de la vista de informe las columnas `latitud`, `longitud` y las claves,
para que quien arme visuales no las arrastre por error.

## 3. Medidas DAX

Crear una tabla vacía llamada `_Medidas` (**Escribir consulta** → `= {1}`,
luego ocultar la columna) y meter ahí todas:

```dax
Homicidios = COUNTROWS ( fact_homicidios )

Víctimas = SUM ( fact_homicidios[victimas] )

Llamadas con violencia = COUNTROWS ( fact_llamadas )

Llamadas por arma de fuego =
CALCULATE ( [Llamadas con violencia], fact_llamadas[familia] = "arma_fuego" )

Homicidios nocturnos =
CALCULATE ( [Homicidios], fact_homicidios[franja] = "nocturna (22-06)" )

% nocturnos = DIVIDE ( [Homicidios nocturnos], [Homicidios] )

Personal desplegado = SUM ( dim_municipio[personal_total] )

Carga de violencia = [Homicidios] * 10 + [Llamadas con violencia]

Carga por 100 elementos =
DIVIDE ( [Carga de violencia] * 100, [Personal desplegado] )

Distancia mediana a base =
MEDIANX ( fact_homicidios, fact_homicidios[dist_base_km] )

Homicidios lejos de base =
CALCULATE ( [Homicidios], fact_homicidios[dist_base_km] > 3 )

Sectores desatendidos =
CALCULATE ( COUNTROWS ( fact_sectores ), fact_sectores[clasificacion] = "DESATENDIDO" )

Sectores saturados =
CALCULATE ( COUNTROWS ( fact_sectores ), fact_sectores[clasificacion] = "SATURADO" )

Casos prioridad alta =
CALCULATE ( COUNTROWS ( fact_auditoria ), fact_auditoria[nivel_revision] = "ALTA" )

Homicidios periodo anterior =
CALCULATE ( [Homicidios], DATEADD ( dim_calendario[fecha], -7, DAY ) )

Variación semanal =
DIVIDE ( [Homicidios] - [Homicidios periodo anterior], [Homicidios periodo anterior] )
```

**`Carga por 100 elementos` es la medida que sirve en la mesa.** Dice cuánta
violencia carga cada coordinación respecto de su propio despliegue, no en
absoluto. Metropolitana carga cinco veces lo que Ixtapan.

## 4. Páginas

Cuatro páginas, en el orden en que se usan en la reunión.

### Página 1 · Homicidios dolosos

El eje del análisis. Nada de 911 arriba.

- Cuatro tarjetas: `Homicidios`, `Víctimas`, `% nocturnos`, `Distancia mediana a base`
- Gráfico de columnas: `Homicidios` por `dim_calendario[fecha]`
- Barras horizontales: `Homicidios` por `dim_municipio[municipio]`, top 10
- Anillo: `Homicidios` por `fact_homicidios[movil]`
- Columnas: `Homicidios` por `fact_homicidios[hora_num]`
- Segmentaciones: `dim_coordinacion[coordinacion]`, `dim_calendario[fecha]`

### Página 2 · Territorio

- **Mapa** (el visual nativo, que usa Bing): `latitud`/`longitud` de
  `fact_homicidios`, tamaño por `victimas`. Estilo de mapa: **Escala de grises**
  o **Carretera**.
- Segundo **mapa** con `fact_sectores`: tamaño por `indice_ceguera`, leyenda por
  `clasificacion`.
- Tabla de `fact_sectores` ordenada por `indice_ceguera`, con `diagnostico`.
- Segmentación por `clasificacion`.

> El visual de mapa de Power BI no dibuja el perímetro del estado por sí solo.
> Si lo quieren trazado, se hace con el visual **Shape Map** o **Azure Maps**
> cargando `poligonos/perimetro_edomex.geojson`; Azure Maps sí pide llave.

### Página 3 · Coordinaciones

La mesa con coordinadores.

- Tabla de `dim_coordinacion` con `Carga de violencia`, `Personal desplegado`,
  **`Carga por 100 elementos`**, `Sectores desatendidos`, `Sectores saturados`
- Formato condicional en barras sobre `Carga por 100 elementos`
- Al hacer clic en una coordinación, el resto de la página se filtra a sus municipios

Poner en la página, como texto fijo, la advertencia de alcance:

> Mide condiciones del territorio, no el desempeño de una persona. Los insumos no
> contienen asignación nominal de mando, turnos ni recorridos.

### Página 4 · Casos a revisar

- Tarjeta: `Casos prioridad alta`
- Tabla de `fact_auditoria` filtrada a `nivel_revision = "ALTA"`, con la columna
  `indicadores` visible
- Segmentación por `nivel_revision` y `incidente`

Texto fijo:

> Prioridad de revisión documental, no dictamen pericial. La determinación de la
> mecánica de los hechos corresponde a los servicios periciales y a la FGJEM.

## 5. Qué no se traslada

Tres cosas del tablero HTML no existen en Power BI y conviene decidirlas antes:

1. **La consola del extractor NLP.** Es interactiva: el operador pega una nota de
   cabina y obtiene la consulta de Windows Maps. Power BI no ejecuta ese código.
   Se queda en el tablero HTML o pasa a una app aparte.
2. **El perímetro del estado trazado.** Requiere Shape Map o Azure Maps.
3. **Los enlaces `bingmaps:` y `ms-drive-to:`** para despachar unidad. En Power BI
   se puede poner la columna como URL web, pero no dispara la app nativa.

Por eso conviene conservar el tablero HTML para la operación y usar Power BI para
la presentación y el seguimiento.

## 6. Actualización

Cada corte nuevo del C5:

1. Copiar el Excel a `insumos/excel/`
2. `python src/etl_sigeo.py`
3. `python src/export_powerbi.py`
4. En Power BI: **Actualizar**

Si el archivo `.pbix` se publica en el servicio de Power BI, los CSV deben vivir
en OneDrive o SharePoint para que la actualización programada los alcance.
