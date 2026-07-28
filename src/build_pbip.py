"""
SIGEO-HD DGSPYT — Generador del proyecto de Power BI.

Escribe un proyecto Power BI (.pbip) completo: modelo semantico con sus
tablas, relaciones y medidas DAX, mas el informe con sus paginas y visuales.
Se abre con doble clic en Power BI Desktop y ya viene armado; no hay que
arrastrar campos ni escribir medidas a mano.

    python src/etl_sigeo.py
    python src/export_powerbi.py
    python src/build_pbip.py

Salida en powerbi/:
    SIGEO-HD.pbip                 <- abrir este
    SIGEO-HD.SemanticModel/       modelo, particiones M, relaciones, medidas
    SIGEO-HD.Report/              paginas y visuales

El modelo lee los CSV de esta misma carpeta. Si el proyecto se mueve de sitio,
basta cambiar el parametro RutaDatos dentro de Power BI.
"""

import json
import shutil
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
PBI = RAIZ / "powerbi"
NOMBRE = "SIGEO-HD"

# --------------------------------------------------------------------------
# Esquema de las tablas: columna -> tipo de datos del modelo tabular
# --------------------------------------------------------------------------
TEXTO, ENTERO, DECIMAL, FECHA = "string", "int64", "double", "dateTime"

TABLAS = {
    "dim_calendario": {
        "archivo": "dim_calendario.csv",
        "columnas": [("fecha", FECHA), ("anio", ENTERO), ("mes", ENTERO),
                     ("nombre_mes", TEXTO), ("dia", ENTERO), ("dia_semana", TEXTO),
                     ("num_dia_semana", ENTERO), ("es_fin_de_semana", TEXTO)],
        "tabla_fechas": True,
    },
    "dim_municipio": {
        "archivo": "dim_municipio.csv",
        "columnas": [("municipio", TEXTO), ("coordinacion", TEXTO),
                     ("subdireccion", TEXTO), ("bases_en_uso", ENTERO),
                     ("personal_total", ENTERO), ("sectores_desatendidos", ENTERO),
                     ("sectores_saturados", ENTERO),
                     ("dist_media_hd_base_km", DECIMAL), ("indice_presion", DECIMAL)],
    },
    "dim_coordinacion": {
        "archivo": "dim_coordinacion.csv",
        "columnas": [("coordinacion", TEXTO), ("total_municipios", ENTERO),
                     ("bases_en_uso", ENTERO), ("personal_total", ENTERO),
                     ("carga_violencia", ENTERO),
                     ("carga_por_100_elementos", DECIMAL),
                     ("sectores_desatendidos", ENTERO),
                     ("sectores_saturados", ENTERO), ("ranking", ENTERO)],
    },
    "fact_homicidios": {
        "archivo": "fact_homicidios.csv",
        "columnas": [("id", ENTERO), ("fecha", FECHA), ("hora", TEXTO),
                     ("hora_num", ENTERO), ("franja", TEXTO), ("municipio", TEXTO),
                     ("coordinacion", TEXTO), ("colonia", TEXTO), ("calle", TEXTO),
                     ("latitud", DECIMAL), ("longitud", DECIMAL),
                     ("victimas", ENTERO), ("sexo", TEXTO), ("movil", TEXTO),
                     ("observaciones", TEXTO), ("dist_base_km", DECIMAL),
                     ("windows_maps_query", TEXTO)],
    },
    "fact_llamadas": {
        "archivo": "fact_llamadas.csv",
        "columnas": [("folio", TEXTO), ("fecha", FECHA), ("hora", TEXTO),
                     ("hora_num", ENTERO), ("franja", TEXTO), ("municipio", TEXTO),
                     ("coordinacion", TEXTO), ("incidente", TEXTO),
                     ("familia", TEXTO), ("peso_violencia", DECIMAL),
                     ("latitud", DECIMAL), ("longitud", DECIMAL),
                     ("geo_confianza", TEXTO)],
    },
    "fact_sectores": {
        "archivo": "fact_sectores.csv",
        "columnas": [("sector_id", TEXTO), ("ranking", ENTERO), ("municipio", TEXTO),
                     ("coordinacion", TEXTO), ("colonias", TEXTO),
                     ("clasificacion", TEXTO), ("latitud", DECIMAL),
                     ("longitud", DECIMAL), ("indice_ceguera", DECIMAL),
                     ("indice_violencia", DECIMAL), ("eventos_hd", ENTERO),
                     ("llamadas_violentas", ENTERO), ("dist_base_km", DECIMAL),
                     ("personal_3km", ENTERO), ("limitrofe", TEXTO),
                     ("diagnostico", TEXTO)],
    },
    "fact_auditoria": {
        "archivo": "fact_auditoria.csv",
        "columnas": [("folio", TEXTO), ("fecha", FECHA), ("hora", TEXTO),
                     ("municipio", TEXTO), ("coordinacion", TEXTO),
                     ("incidente", TEXTO), ("nivel_revision", TEXTO),
                     ("puntaje", ENTERO), ("num_indicadores", ENTERO),
                     ("indicadores", TEXTO), ("accion_sugerida", TEXTO),
                     ("latitud", DECIMAL), ("longitud", DECIMAL)],
    },
}

TIPO_M = {TEXTO: "type text", ENTERO: "Int64.Type",
          DECIMAL: "type number", FECHA: "type date"}

RELACIONES = [
    ("dim_calendario", "fecha", "fact_homicidios", "fecha"),
    ("dim_calendario", "fecha", "fact_llamadas", "fecha"),
    ("dim_calendario", "fecha", "fact_auditoria", "fecha"),
    ("dim_municipio", "municipio", "fact_homicidios", "municipio"),
    ("dim_municipio", "municipio", "fact_llamadas", "municipio"),
    ("dim_municipio", "municipio", "fact_sectores", "municipio"),
    ("dim_municipio", "municipio", "fact_auditoria", "municipio"),
    ("dim_coordinacion", "coordinacion", "dim_municipio", "coordinacion"),
]

MEDIDAS = [
    ("Homicidios", "COUNTROWS ( fact_homicidios )", "#,0"),
    ("Víctimas", "SUM ( fact_homicidios[victimas] )", "#,0"),
    ("Llamadas con violencia", "COUNTROWS ( fact_llamadas )", "#,0"),
    ("Llamadas por arma de fuego",
     'CALCULATE ( [Llamadas con violencia], fact_llamadas[familia] = "arma_fuego" )', "#,0"),
    ("Homicidios nocturnos",
     'CALCULATE ( [Homicidios], fact_homicidios[franja] = "nocturna (22-06)" )', "#,0"),
    ("% nocturnos", "DIVIDE ( [Homicidios nocturnos], [Homicidios] )", "0.0%"),
    ("Personal desplegado", "SUM ( dim_municipio[personal_total] )", "#,0"),
    ("Carga de violencia", "[Homicidios] * 10 + [Llamadas con violencia]", "#,0"),
    ("Carga por 100 elementos",
     "DIVIDE ( [Carga de violencia] * 100, [Personal desplegado] )", "#,0.0"),
    ("Distancia mediana a base",
     "MEDIANX ( fact_homicidios, fact_homicidios[dist_base_km] )", "#,0.00"),
    ("Homicidios lejos de base",
     "CALCULATE ( [Homicidios], fact_homicidios[dist_base_km] > 3 )", "#,0"),
    ("Sectores desatendidos",
     'CALCULATE ( COUNTROWS ( fact_sectores ), fact_sectores[clasificacion] = "DESATENDIDO" )', "#,0"),
    ("Sectores saturados",
     'CALCULATE ( COUNTROWS ( fact_sectores ), fact_sectores[clasificacion] = "SATURADO" )', "#,0"),
    ("Casos prioridad alta",
     'CALCULATE ( COUNTROWS ( fact_auditoria ), fact_auditoria[nivel_revision] = "ALTA" )', "#,0"),
    ("Casos auditados", "COUNTROWS ( fact_auditoria )", "#,0"),
    ("Homicidios periodo anterior",
     "CALCULATE ( [Homicidios], DATEADD ( dim_calendario[fecha], -7, DAY ) )", "#,0"),
    ("Variación semanal",
     "DIVIDE ( [Homicidios] - [Homicidios periodo anterior], [Homicidios periodo anterior] )", "0.0%"),
]


def particion_m(nombre, cfg):
    """Consulta Power Query que lee el CSV y tipa las columnas."""
    tipos = ", ".join(f'{{"{c}", {TIPO_M[t]}}}' for c, t in cfg["columnas"])
    return (
        "let\n"
        f'    Origen = Csv.Document(File.Contents(RutaDatos & "\\{cfg["archivo"]}"), '
        "[Delimiter=\",\", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),\n"
        "    Encabezados = Table.PromoteHeaders(Origen, [PromoteAllScalars=true]),\n"
        f"    Tipos = Table.TransformColumnTypes(Encabezados, {{{tipos}}})\n"
        "in\n"
        "    Tipos"
    )


def construir_modelo():
    tablas = []
    for nombre, cfg in TABLAS.items():
        columnas = []
        for c, t in cfg["columnas"]:
            col = {
                "name": c, "dataType": t, "sourceColumn": c,
                "lineageTag": f"col-{nombre}-{c}",
                "summarizeBy": "none",
            }
            if t == FECHA:
                col["formatString"] = "yyyy-mm-dd"
            if c in ("latitud",):
                col["dataCategory"] = "Latitude"
            if c in ("longitud",):
                col["dataCategory"] = "Longitude"
            if c == "municipio":
                col["dataCategory"] = "City"
            columnas.append(col)

        tabla = {
            "name": nombre,
            "lineageTag": f"tbl-{nombre}",
            "columns": columnas,
            "partitions": [{
                "name": nombre, "mode": "import",
                "source": {"type": "m", "expression": particion_m(nombre, cfg)},
            }],
        }
        if cfg.get("tabla_fechas"):
            tabla["dataCategory"] = "Time"
            for col in tabla["columns"]:
                if col["name"] == "fecha":
                    col["isKey"] = True
        tablas.append(tabla)

    # Las medidas viven en la tabla de coordinaciones para no crear una tabla
    # vacia que Power BI Desktop no puede generar desde el archivo.
    destino = next(t for t in tablas if t["name"] == "dim_coordinacion")
    destino["measures"] = [
        {"name": n, "expression": e, "formatString": f,
         "lineageTag": f"med-{i}"}
        for i, (n, e, f) in enumerate(MEDIDAS)
    ]

    relaciones = [{
        "name": f"rel-{i}",
        "fromTable": ft, "fromColumn": fc,
        "toTable": tt, "toColumn": tc,
        "crossFilteringBehavior": "oneDirection",
    } for i, (tt, tc, ft, fc) in enumerate(RELACIONES)]

    # La ruta de los CSV va como parametro de Power Query (expresion del
    # modelo), no como tabla: asi aparece editable en "Administrar parametros"
    # y el proyecto se puede mover de carpeta sin tocar cada consulta.
    ruta = str(PBI).replace("\\", "\\\\")
    expresiones = [{
        "name": "RutaDatos",
        "kind": "m",
        "expression": (f'"{ruta}" meta [IsParameterQuery=true, Type="Text", '
                       'IsParameterQueryRequired=true]'),
        "lineageTag": "expr-ruta-datos",
        "annotations": [{"name": "PBI_NavigationStepName", "value": "Navegación"},
                        {"name": "PBI_ResultType", "value": "Text"}],
    }]

    return {
        "compatibilityLevel": 1567,
        "model": {
            "culture": "es-MX",
            "dataAccessOptions": {"legacyRedirects": True,
                                  "returnErrorValuesAsNull": True},
            "defaultPowerBIDataSourceVersion": "powerBI_V3",
            "sourceQueryCulture": "es-MX",
            "expressions": expresiones,
            "tables": tablas,
            "relationships": relaciones,
            "annotations": [{"name": "PBI_QueryOrder",
                             "value": json.dumps(["RutaDatos"] + list(TABLAS))}],
        },
    }


# --------------------------------------------------------------------------
# Informe: paginas y visuales
# --------------------------------------------------------------------------

def visual(x, y, ancho, alto, tipo, proyecciones, titulo=None, orden=0):
    """Contenedor de visual en el formato que guarda Power BI Desktop."""
    cfg = {
        "name": f"v{orden}-{tipo}-{x}-{y}",
        "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": orden,
                                           "width": ancho, "height": alto}}],
        "singleVisual": {
            "visualType": tipo,
            "projections": proyecciones,
            "drillFilterOtherVisuals": True,
        },
    }
    if titulo:
        cfg["singleVisual"]["vcObjects"] = {
            "title": [{"properties": {
                "text": {"expr": {"Literal": {"Value": f"'{titulo}'"}}},
                "show": {"expr": {"Literal": {"Value": "true"}}},
            }}]
        }
    return {"x": x, "y": y, "z": orden, "width": ancho, "height": alto,
            "config": json.dumps(cfg, ensure_ascii=False),
            "filters": "[]"}


def campo(tabla, columna):
    return {"queryRef": f"{tabla}.{columna}",
            "active": True} if False else {"queryRef": f"{tabla}.{columna}"}


def medida(nombre):
    return {"queryRef": f"dim_coordinacion.{nombre}"}


def construir_informe():
    A = 1280.0  # ancho de pagina
    paginas = []

    # ---- Pagina 1: Homicidios dolosos ----
    v, o = [], 0
    for i, m in enumerate(["Homicidios", "Víctimas", "% nocturnos",
                           "Distancia mediana a base"]):
        o += 1
        v.append(visual(20 + i * 305, 20, 285, 120, "card",
                        {"Values": [medida(m)]}, m, o))
    o += 1
    v.append(visual(20, 155, 610, 270, "columnChart",
                    {"Category": [campo("dim_calendario", "fecha")],
                     "Y": [medida("Homicidios")]},
                    "Homicidios dolosos por día", o))
    o += 1
    v.append(visual(645, 155, 615, 270, "barChart",
                    {"Category": [campo("dim_municipio", "municipio")],
                     "Y": [medida("Homicidios")]},
                    "Homicidios por municipio", o))
    o += 1
    v.append(visual(20, 440, 400, 240, "columnChart",
                    {"Category": [campo("fact_homicidios", "hora_num")],
                     "Y": [medida("Homicidios")]},
                    "Homicidios por hora del día", o))
    o += 1
    v.append(visual(435, 440, 400, 240, "donutChart",
                    {"Category": [campo("fact_homicidios", "movil")],
                     "Y": [medida("Homicidios")]},
                    "Móvil de la agresión", o))
    o += 1
    v.append(visual(850, 440, 410, 240, "slicer",
                    {"Values": [campo("dim_coordinacion", "coordinacion")]},
                    "Coordinación regional", o))
    paginas.append(("Homicidios dolosos", v))

    # ---- Pagina 2: Territorio ----
    v, o = [], 0
    o += 1
    v.append(visual(20, 20, 620, 400, "map",
                    {"Category": [campo("fact_homicidios", "municipio")],
                     "Size": [medida("Homicidios")]},
                    "Homicidios dolosos en el territorio", o))
    o += 1
    v.append(visual(655, 20, 605, 400, "map",
                    {"Category": [campo("fact_sectores", "sector_id")],
                     "Size": [campo("fact_sectores", "indice_ceguera")],
                     "Series": [campo("fact_sectores", "clasificacion")]},
                    "Sectores de zona ciega", o))
    o += 1
    v.append(visual(20, 435, 940, 245, "tableEx",
                    {"Values": [campo("fact_sectores", "ranking"),
                                campo("fact_sectores", "municipio"),
                                campo("fact_sectores", "clasificacion"),
                                campo("fact_sectores", "indice_ceguera"),
                                campo("fact_sectores", "dist_base_km"),
                                campo("fact_sectores", "personal_3km"),
                                campo("fact_sectores", "diagnostico")]},
                    "Sectores ordenados por brecha de cobertura", o))
    o += 1
    v.append(visual(975, 435, 285, 245, "slicer",
                    {"Values": [campo("fact_sectores", "clasificacion")]},
                    "Diagnóstico", o))
    paginas.append(("Territorio", v))

    # ---- Pagina 3: Coordinaciones ----
    v, o = [], 0
    o += 1
    v.append(visual(20, 20, 1240, 300, "tableEx",
                    {"Values": [campo("dim_coordinacion", "ranking"),
                                campo("dim_coordinacion", "coordinacion"),
                                campo("dim_coordinacion", "total_municipios"),
                                medida("Carga de violencia"),
                                medida("Personal desplegado"),
                                medida("Carga por 100 elementos"),
                                medida("Sectores desatendidos"),
                                medida("Sectores saturados")]},
                    "Coordinaciones regionales", o))
    o += 1
    v.append(visual(20, 335, 620, 345, "barChart",
                    {"Category": [campo("dim_coordinacion", "coordinacion")],
                     "Y": [medida("Carga por 100 elementos")]},
                    "Carga de violencia por cada 100 elementos adscritos", o))
    o += 1
    v.append(visual(655, 335, 605, 345, "tableEx",
                    {"Values": [campo("dim_municipio", "municipio"),
                                campo("dim_municipio", "coordinacion"),
                                campo("dim_municipio", "indice_presion"),
                                campo("dim_municipio", "personal_total"),
                                campo("dim_municipio", "sectores_desatendidos")]},
                    "Detalle por municipio", o))
    paginas.append(("Coordinaciones", v))

    # ---- Pagina 4: Casos a revisar ----
    v, o = [], 0
    for i, m in enumerate(["Casos auditados", "Casos prioridad alta"]):
        o += 1
        v.append(visual(20 + i * 305, 20, 285, 120, "card",
                        {"Values": [medida(m)]}, m, o))
    o += 1
    v.append(visual(630, 20, 630, 120, "slicer",
                    {"Values": [campo("fact_auditoria", "nivel_revision")]},
                    "Prioridad de revisión", o))
    o += 1
    v.append(visual(20, 155, 1240, 525, "tableEx",
                    {"Values": [campo("fact_auditoria", "folio"),
                                campo("fact_auditoria", "municipio"),
                                campo("fact_auditoria", "incidente"),
                                campo("fact_auditoria", "nivel_revision"),
                                campo("fact_auditoria", "puntaje"),
                                campo("fact_auditoria", "indicadores"),
                                campo("fact_auditoria", "accion_sugerida")]},
                    "Decesos dudosos, suicidios y personas no localizadas", o))
    paginas.append(("Casos a revisar", v))

    secciones = []
    for i, (nombre, visuales) in enumerate(paginas):
        secciones.append({
            "name": f"pagina{i}",
            "displayName": nombre,
            "filters": "[]",
            "ordinal": i,
            "visualContainers": visuales,
            "config": json.dumps({"visibility": 0}),
            "displayOption": 1,
            "width": A,
            "height": 720.0,
        })

    return {
        "id": 0,
        "resourcePackages": [],
        "sections": secciones,
        "config": json.dumps({
            "version": "5.43",
            "themeCollection": {"baseTheme": {"name": "CY24SU10"}},
            "activeSectionIndex": 0,
            "defaultDrillFilterOtherVisuals": True,
            "settings": {"useStylableVisualContainerHeader": True},
        }),
        "layoutOptimization": 0,
    }


# --------------------------------------------------------------------------

def main():
    faltantes = [c["archivo"] for c in TABLAS.values()
                 if not (PBI / c["archivo"]).exists()]
    if faltantes:
        print("Faltan los CSV: " + ", ".join(faltantes))
        print("Ejecuta primero:  python src/export_powerbi.py")
        return 1

    modelo_dir = PBI / f"{NOMBRE}.SemanticModel"
    informe_dir = PBI / f"{NOMBRE}.Report"
    for d in (modelo_dir, informe_dir):
        if d.exists():
            shutil.rmtree(d)
        d.mkdir(parents=True)

    def escribir(ruta, contenido):
        texto = (json.dumps(contenido, ensure_ascii=False, indent=2)
                 if not isinstance(contenido, str) else contenido)
        ruta.write_text(texto, encoding="utf-8")

    # --- Proyecto ---
    escribir(PBI / f"{NOMBRE}.pbip", {
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NOMBRE}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    })

    # --- Modelo semantico ---
    escribir(modelo_dir / "model.bim", construir_modelo())
    escribir(modelo_dir / "definition.pbism",
             {"version": "1.0", "settings": {}})
    escribir(modelo_dir / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "SemanticModel", "displayName": NOMBRE},
        "config": {"version": "2.0", "logicalId": "00000000-0000-0000-0000-00000000d001"},
    })

    # --- Informe ---
    escribir(informe_dir / "report.json", construir_informe())
    escribir(informe_dir / "definition.pbir", {
        "version": "1.0",
        "datasetReference": {"byPath": {"path": f"../{NOMBRE}.SemanticModel"}},
    })
    escribir(informe_dir / ".platform", {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
        "metadata": {"type": "Report", "displayName": NOMBRE},
        "config": {"version": "2.0", "logicalId": "00000000-0000-0000-0000-00000000d002"},
    })
    (informe_dir / "StaticResources" / "RegisteredResources").mkdir(parents=True)

    modelo = construir_modelo()
    print("SIGEO-HD DGSPYT · proyecto de Power BI generado")
    print(f"  tablas ..................... {len(modelo['model']['tables'])}")
    print(f"  relaciones ................. {len(modelo['model']['relationships'])}")
    print(f"  medidas DAX ................ {len(MEDIDAS)}")
    print(f"  páginas del informe ........ {len(construir_informe()['sections'])}")
    print(f"  abrir ...................... {PBI / (NOMBRE + '.pbip')}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
