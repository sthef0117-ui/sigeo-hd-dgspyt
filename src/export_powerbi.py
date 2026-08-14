"""
SIGEO-HD DGSPYT — Exportacion a Power BI.

Convierte los JSON calculados en un modelo en estrella listo para cargar en
Power BI Desktop, en CSV plano.

    python src/etl_sigeo.py
    python src/export_powerbi.py

Salida en powerbi/:
    dim_calendario.csv      Una fila por dia del periodo
    dim_municipio.csv       Municipio con su coordinacion y despliegue
    dim_coordinacion.csv    Las 8 coordinaciones territoriales
    fact_homicidios.csv     Un renglon por homicidio doloso corroborado
    fact_llamadas.csv       Un renglon por llamada con violencia
    fact_sectores.csv       Un renglon por sector de zona ciega
    fact_auditoria.csv      Un renglon por caso de la auditoria de decesos

Por que CSV y no una conexion directa: el area puede abrir, revisar y validar
cada archivo en Excel antes de cargarlo, y el modelo queda auditable. La
relacion entre tablas se arma en Power BI con el campo municipio.

Codificacion UTF-8 con BOM: sin el BOM, Power BI y Excel abren los acentos
rotos y los nombres de municipio dejan de coincidir entre tablas.
"""

import csv
import json
import math
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
ANALISIS = RAIZ / "analisis"
SALIDA = RAIZ / "powerbi"
INSUMOS = RAIZ / "insumos"
EVID = SALIDA / "evidencia"
# El proyecto se abre desde la copia principal, no desde el worktree. La ruta
# de las fotos (columna ImageUrl) debe apuntar ahí para que rendericen.
COPIA_PRINCIPAL = Path(r"C:\Users\xXZaB\.gemini\antigravity\scratch\sigeo_hd_dgspyt")


def _foto_uri(stem):
    ruta = COPIA_PRINCIPAL / "powerbi" / "evidencia" / (stem + ".png")
    return "file:///" + str(ruta).replace("\\", "/")


def _fecha_de_nombre(nombre):
    """Deriva la fecha de la captura desde su nombre: IMG-20260723-... o
    'Captura ... 2026-08-13 ...'. Es la fecha de la imagen, no del hecho."""
    m = re.search(r"(20\d{2})[-_]?(\d{2})[-_]?(\d{2})", nombre)
    if not m:
        return ""
    y, mo, d = m.group(1), m.group(2), m.group(3)
    if "01" <= mo <= "12" and "01" <= d <= "31":
        return f"{y}-{mo}-{d}"
    return ""


def construir_evidencia():
    """
    Recorta la foto real de cada captura de WhatsApp y la deja en powerbi/
    evidencia/ para anexarla al Power BI como evidencia (uso interno, reservado).
    Devuelve las filas de dim_evidencia. Best-effort: si algo no se puede leer,
    se omite sin romper el resto.
    """
    sys.path.insert(0, str(Path(__file__).parent))
    try:
        from recorte import recortar_foto
    except Exception:
        print("  evidencia .................. omitida (falta Pillow/recorte)")
        return []
    EVID.mkdir(parents=True, exist_ok=True)
    exts = {".jpg", ".jpeg", ".png", ".webp"}
    origen = []
    for sub in ("whatsapp", "bandeja"):
        d = INSUMOS / sub
        if d.exists():
            for p in sorted(d.iterdir()):
                if p.is_file() and p.suffix.lower() in exts and p.name != "CATALOGO.md":
                    origen.append(p)
    filas = []
    for p in origen:
        destino = EVID / (p.stem + ".png")
        try:
            recortar_foto(p, destino)
        except Exception:
            continue
        filas.append({
            "fecha": _fecha_de_nombre(p.name),
            "municipio": "",
            "lugar": "",
            "caso": p.stem[:60],
            "foto": _foto_uri(p.stem),
        })
    print(f"  evidencia .................. {len(filas)} foto(s) en {EVID}")
    return filas

DIAS_ES = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
MESES_ES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
            "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def leer(nombre):
    ruta = ANALISIS / nombre
    if not ruta.exists():
        print(f"Falta {ruta.name}. Ejecuta primero: python src/etl_sigeo.py")
        sys.exit(1)
    return json.loads(ruta.read_text(encoding="utf-8"))


def escribir(nombre, campos, filas):
    SALIDA.mkdir(parents=True, exist_ok=True)
    ruta = SALIDA / nombre
    # utf-8-sig: Power BI y Excel necesitan el BOM para no romper los acentos.
    with ruta.open("w", encoding="utf-8-sig", newline="") as f:
        escritor = csv.DictWriter(f, fieldnames=campos, extrasaction="ignore")
        escritor.writeheader()
        escritor.writerows(filas)
    print(f"  {nombre:<24} {len(filas):>6} filas")
    return ruta


def haversine_km(lat1, lng1, lat2, lng2):
    r = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    a = (math.sin((p2 - p1) / 2) ** 2
         + math.cos(p1) * math.cos(p2) * math.sin(math.radians(lng2 - lng1) / 2) ** 2)
    return 2 * r * math.asin(math.sqrt(a))


def hora_num(hhmmss):
    try:
        return int(str(hhmmss)[:2])
    except (ValueError, TypeError):
        return None


def franja(h):
    if h is None:
        return "sin dato"
    if 22 <= h or h < 6:
        return "nocturna (22-06)"
    if 6 <= h < 14:
        return "matutina (06-14)"
    return "vespertina (14-22)"


def main():
    hd = leer("corroborados_sigeo.json")
    llamadas = leer("llamadas_911_sigeo.json")
    bases = leer("bases_dgspyt.json")
    sectores = leer("zonas_ciegas.json")
    auditoria = leer("auditoria_decesos.json")
    territorio = leer("perfil_territorial.json")
    coord = leer("perfil_coordinaciones.json")

    print("SIGEO-HD DGSPYT · exportación a Power BI")

    coordinacion_de = {t["municipio"]: t["coordinacion"] for t in territorio}
    bases_geo = [b for b in bases if b["lat"] and b["en_uso"]]

    def dist_base(lat, lng):
        if not lat or not bases_geo:
            return None
        return round(min(haversine_km(lat, lng, b["lat"], b["lng"])
                         for b in bases_geo), 2)

    # --- Dimensiones -------------------------------------------------------
    # Las cifras de personal no se exportan: se usa el nivel de despliegue
    # (banda). La precisión numérica se reserva a los eventos de homicidio.
    escribir("dim_municipio.csv",
             ["municipio", "coordinacion", "subdireccion", "bases_en_uso",
              "despliegue_banda", "sectores_desatendidos", "sectores_saturados",
              "dist_media_hd_base_km", "indice_presion"],
             [{**t, "coordinacion": t["coordinacion"] or "SIN CATALOGAR"}
              for t in territorio])

    escribir("dim_coordinacion.csv",
             ["coordinacion", "total_municipios", "bases_en_uso", "despliegue_banda",
              "carga_violencia", "presion_banda",
              "sectores_desatendidos", "sectores_saturados", "ranking"],
             coord["coordinaciones"])

    # --- Calendario --------------------------------------------------------
    fechas = sorted({d["fecha"] for d in hd if d["fecha"]}
                    | {l["fecha"] for l in llamadas if l["fecha"]})
    if fechas:
        ini = datetime.strptime(fechas[0], "%Y-%m-%d").date()
        fin = datetime.strptime(fechas[-1], "%Y-%m-%d").date()
        calendario = []
        d = ini
        while d <= fin:
            calendario.append({
                "fecha": d.isoformat(),
                "anio": d.year,
                "mes": d.month,
                "nombre_mes": MESES_ES[d.month - 1],
                "dia": d.day,
                "dia_semana": DIAS_ES[d.weekday()],
                "num_dia_semana": d.weekday() + 1,
                "es_fin_de_semana": "Sí" if d.weekday() >= 5 else "No",
            })
            d += timedelta(days=1)
        escribir("dim_calendario.csv", list(calendario[0].keys()), calendario)

    # --- Hechos ------------------------------------------------------------
    filas_hd = []
    for d in hd:
        h = hora_num(d["hora"])
        filas_hd.append({
            "id": d["id"],
            "fecha": d["fecha"],
            "hora": d["hora"],
            "hora_num": h,
            "franja": franja(h),
            "municipio": d["municipio"],
            "coordinacion": coordinacion_de.get(d["municipio"], "") or "SIN CATALOGAR",
            "colonia": d["colonia"],
            "calle": d["calle"],
            "latitud": d["lat"],
            "longitud": d["lng"],
            "victimas": d["total_hd"],
            "sexo": d["sexo"],
            "movil": d["movil"],
            "observaciones": d["observaciones"],
            "dist_base_km": dist_base(d["lat"], d["lng"]),
            "windows_maps_query": d["windows_maps_query"],
        })
    escribir("fact_homicidios.csv", list(filas_hd[0].keys()), filas_hd)

    filas_ll = []
    for l in llamadas:
        if l["peso_violencia"] <= 0:
            continue
        h = hora_num(l["hora"])
        filas_ll.append({
            "folio": l["folio"],
            "fecha": l["fecha"],
            "hora": l["hora"],
            "hora_num": h,
            "franja": franja(h),
            "municipio": l["municipio"],
            "coordinacion": coordinacion_de.get(l["municipio"], "") or "SIN CATALOGAR",
            "incidente": l["incidente"],
            "familia": l["familia"],
            "peso_violencia": l["peso_violencia"],
            "latitud": l["lat"],
            "longitud": l["lng"],
            "geo_confianza": l["geo_confianza"],
        })
    escribir("fact_llamadas.csv", list(filas_ll[0].keys()), filas_ll)

    filas_sec = [{
        "sector_id": s["sector_id"],
        "ranking": s["ranking"],
        "municipio": s["municipio"],
        "coordinacion": coordinacion_de.get(s["municipio"], "") or "SIN CATALOGAR",
        "colonias": ", ".join(s["colonias"]),
        "clasificacion": s["clasificacion"],
        "latitud": s["lat"],
        "longitud": s["lng"],
        "indice_ceguera": s["indice_ceguera"],
        "indice_violencia": s["indice_violencia"],
        "eventos_hd": s["eventos_hd"],
        "llamadas_violentas": s["llamadas_violentas"],
        "dist_base_km": s["dist_base_km"],
        "despliegue_banda": s.get("despliegue_banda",""),
        "limitrofe": "Sí" if s["limitrofe"] else "No",
        "diagnostico": s["diagnostico"],
    } for s in sectores]
    escribir("fact_sectores.csv", list(filas_sec[0].keys()), filas_sec)

    filas_aud = [{
        "folio": a["folio"],
        "fecha": a["fecha"],
        "hora": a["hora"],
        "municipio": a["municipio"],
        "coordinacion": coordinacion_de.get(a["municipio"], "") or "SIN CATALOGAR",
        "incidente": a["incidente"],
        "nivel_revision": a["nivel_revision"],
        "puntaje": a["puntaje"],
        "num_indicadores": len(a["indicadores"]),
        "indicadores": " | ".join(i["motivo"] for i in a["indicadores"]),
        "accion_sugerida": a["accion_sugerida"],
        "latitud": a["lat"],
        "longitud": a["lng"],
    } for a in auditoria]
    escribir("fact_auditoria.csv", list(filas_aud[0].keys()), filas_aud)

    # --- Evidencia (fotos de escena, reservadas, uso interno) --------------
    evid = construir_evidencia()
    if evid:
        escribir("dim_evidencia.csv", ["fecha", "municipio", "lugar", "caso", "foto"], evid)

    print(f"  carpeta .................. {SALIDA}")
    print("  Siguiente paso: powerbi/MODELO.md")


if __name__ == "__main__":
    main()
