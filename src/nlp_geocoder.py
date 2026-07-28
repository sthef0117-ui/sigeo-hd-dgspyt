"""
SIGEO-HD DGSPYT — Extractor NLP de ubicaciones en formato Windows Maps.

Resuelve el problema planteado en la reunion de mandos C5:

    "Puntearlos, la tabla carece de datos suficientes. Algunas traen descripcion
     -> Encontrar la forma de que un sistema lo lea."
    "[Maps de windows (Open maps)] herramienta a usar - porque lee ubicaciones
     en formato de descripcion."

El modulo toma los tres campos que el C5 si captura con regularidad
(DIRECCION, REFERENCIA DE UBICACION y NOTAS de cabina) y construye una
cadena de consulta que Windows Maps / Bing Maps resuelve directamente,
mas los URI nativos `bingmaps:` y `ms-drive-to:`.

No inventa ubicaciones: si no hay senal suficiente, marca confianza NULA.

La misma logica esta replicada en JavaScript dentro del dashboard
(modulo 2) para uso en vivo por el operador de cabina.
"""

import re
import unicodedata
from urllib.parse import quote

ESTADO = "Estado de México"
PAIS = "México"

# Ruido tipico de la exportacion C5: celdas vacias marcadas como "0".
NULOS = {"", "0", "0.0", "NONE", "NULL", "N/A", "NA", "SIN DATO", "SIN INFORMACION"}

# --- Vocabulario vial -------------------------------------------------------

TIPOS_VIA = [
    "CARRETERA", "AUTOPISTA", "BOULEVARD", "BOULEVARD", "BLVD", "AVENIDA", "AV",
    "CALZADA", "CALLE", "CERRADA", "CDA", "PRIVADA", "PRIV", "ANDADOR", "AND",
    "CALLEJON", "PROLONGACION", "PROL", "CIRCUITO", "RETORNO", "EJE", "VIA",
    "CAMINO", "PASEO", "PLAZA", "GLORIETA", "PUENTE",
]

TIPOS_ASENTAMIENTO = [
    "COLONIA", "COL", "FRACCIONAMIENTO", "FRACC", "BARRIO", "PUEBLO", "EJIDO",
    "UNIDAD HABITACIONAL", "U HABITACIONAL", "CONJUNTO URBANO", "RANCHERIA",
    "COMUNIDAD", "SECCION", "AMPLIACION", "AMPL", "RESIDENCIAL", "VILLAS",
]

# Puntos de interes que Windows Maps resuelve bien como ancla de busqueda.
PUNTOS_INTERES = [
    "OXXO", "SEVEN ELEVEN", "7 ELEVEN", "AURRERA", "BODEGA AURRERA", "SORIANA",
    "CHEDRAUI", "WALMART", "ELEKTRA", "COPPEL", "FARMACIA", "GASOLINERA",
    "GASOLINERIA", "PEMEX", "MERCADO", "TIANGUIS", "PANTEON", "IGLESIA",
    "CAPILLA", "PARROQUIA", "HOSPITAL", "CLINICA", "IMSS", "ISSSTE", "CRUZ ROJA",
    "ESCUELA", "PRIMARIA", "SECUNDARIA", "PREPARATORIA", "KINDER", "JARDIN DE NINOS",
    "CECYTEM", "CONALEP", "CBT", "UNIVERSIDAD", "DEPORTIVO", "PARQUE", "CANCHA",
    "PLAZA COMERCIAL", "CENTRO COMERCIAL", "TERMINAL", "PARADERO", "BASE DE TAXIS",
    "PUENTE PEATONAL", "DISTRIBUIDOR VIAL", "CASETA", "PALACIO MUNICIPAL",
    "COMANDANCIA", "TECALLI", "MODULO DE VIGILANCIA", "KIOSCO", "ZOCALO",
]

_RE_ESPACIOS = re.compile(r"\s+")
_RE_BASURA = re.compile(r"[<>*;|]+")


def _norm(texto):
    """Mayusculas sin acentos, para comparar. Nunca se devuelve al usuario."""
    if texto is None:
        return ""
    s = unicodedata.normalize("NFD", str(texto))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return _RE_ESPACIOS.sub(" ", s.upper()).strip()


def _limpio(texto):
    """Limpia para mostrar: quita separadores de exportacion y espacios dobles."""
    if texto is None:
        return ""
    s = _RE_BASURA.sub(" ", str(texto))
    s = _RE_ESPACIOS.sub(" ", s).strip(" ,.-/")
    return s.strip()


def _vacio(texto):
    return _norm(texto) in NULOS


def _titulo_via(texto):
    """Normaliza abreviaturas viales para que Maps las resuelva mejor."""
    s = _limpio(texto)
    if not s:
        return ""
    reemplazos = [
        (r"^C\s+", "CALLE "), (r"^AV\s+", "AVENIDA "), (r"^AV\.\s*", "AVENIDA "),
        (r"^CDA\s+", "CERRADA "), (r"^PRIV\s+", "PRIVADA "),
        (r"^PROL\s+", "PROLONGACION "), (r"^BLVD\s+", "BOULEVARD "),
        (r"^CALZ\s+", "CALZADA "), (r"^AND\s+", "ANDADOR "),
        (r"^CARR\s+", "CARRETERA "), (r"^COL\s+", "COLONIA "),
        (r"^FRACC\s+", "FRACCIONAMIENTO "),
    ]
    up = _norm(s)
    for patron, destino in reemplazos:
        if re.match(patron, up):
            return re.sub(patron, destino, up, count=1).strip()
    return up


# --- Parseo del campo DIRECCION del C5 --------------------------------------
# Formatos observados en la exportacion:
#   ";GOMEZ FARIAS|TEMOAYA"
#   "; BENITO JUAREZ 27/MARIPOSA MONARCA|SAN JOSE DEL RINCON"
#   ";AV LA MAGDALENA / FRANCISCO I MADERO|CENTRO"
# Izquierda del "|" = vialidad (con posible cruce tras "/"), derecha = colonia.

def parse_direccion_c5(direccion):
    """Descompone el campo DIRECCION del C5. Devuelve dict con partes."""
    salida = {"vialidad": "", "cruce": "", "numero": "", "colonia": ""}
    if _vacio(direccion):
        return salida

    crudo = str(direccion).lstrip(";").strip()
    if "|" in crudo:
        izq, der = crudo.split("|", 1)
        colonia = _titulo_via(der)
        # La cabina captura "OTRA <MUNICIPIO>" cuando desconoce el asentamiento:
        # es un marcador de dato faltante, no un nombre de colonia.
        if not re.match(r"^OTR[AO]\b", colonia):
            salida["colonia"] = colonia
    else:
        izq = crudo

    izq = _limpio(izq)
    # Cruce de vialidades: "A / B" o "A ESQ B" o "A ESQUINA B"
    partes = re.split(r"\s*/\s*|\s+ESQ(?:UINA)?\.?\s+(?:CON\s+)?", izq, maxsplit=1)
    principal = partes[0].strip() if partes else ""
    if len(partes) > 1:
        salida["cruce"] = _titulo_via(partes[1])

    # Numero exterior al final de la vialidad principal ("BENITO JUAREZ 27").
    m = re.search(r"\s(\d{1,5}[A-Z]?)$", _norm(principal))
    if m:
        salida["numero"] = m.group(1)
        principal = principal[: m.start()].strip()

    salida["vialidad"] = _titulo_via(principal)
    return salida


# --- Extraccion NLP sobre texto libre (NOTAS / REFERENCIA) ------------------

def _buscar_via(texto_norm):
    """Primera vialidad explicita mencionada en el texto."""
    alternancia = "|".join(sorted(TIPOS_VIA, key=len, reverse=True))
    patron = re.compile(
        r"\b(" + alternancia + r")\.?\s+"
        r"([A-Z0-9NÑ'\.\- ]{3,45}?)"
        r"(?=\s*(?:,|\.|\bY\b|\bENTRE\b|\bESQ\b|\bESQUINA\b|\bCOL\b|\bCOLONIA\b|\bFRENTE\b|\bNUMERO\b|\bNO\b|\bMUNICIPIO\b|$))"
    )
    m = patron.search(texto_norm)
    if not m:
        return ""
    return _limpio(f"{m.group(1)} {m.group(2)}")


def _buscar_asentamiento(texto_norm):
    alternancia = "|".join(sorted(TIPOS_ASENTAMIENTO, key=len, reverse=True))
    patron = re.compile(
        r"\b(" + alternancia + r")\.?\s+"
        r"([A-Z0-9NÑ'\.\- ]{3,45}?)"
        r"(?=\s*(?:,|\.|\bY\b|\bENTRE\b|\bESQ\b|\bCALLE\b|\bAV\b|\bAVENIDA\b|\bFRENTE\b|\bMUNICIPIO\b|$))"
    )
    m = patron.search(texto_norm)
    if not m:
        return ""
    nombre = _limpio(m.group(2))
    return nombre if nombre else ""


def _buscar_entre_calles(texto_norm):
    m = re.search(
        r"\bENTRE\s+([A-Z0-9NÑ'\.\- ]{3,40}?)\s+Y\s+([A-Z0-9NÑ'\.\- ]{3,40}?)"
        r"(?=\s*(?:,|\.|\bCOL\b|\bCOLONIA\b|\bFRENTE\b|$))",
        texto_norm,
    )
    if not m:
        return []
    return [_limpio(m.group(1)), _limpio(m.group(2))]


def _buscar_numero(texto_norm):
    m = re.search(r"\b(?:NUMERO|NUM|NO|#)\.?\s*(\d{1,5}[A-Z]?)\b", texto_norm)
    return m.group(1) if m else ""


def _buscar_kilometro(texto_norm):
    m = re.search(r"\bKM\.?\s*(\d{1,3}(?:[\.\+]\d{1,3})?)\b", texto_norm)
    return f"KM {m.group(1)}" if m else ""


def _buscar_punto_interes(texto_norm):
    """Ancla de busqueda tipo negocio/equipamiento; Maps los resuelve bien."""
    # Preferir la forma "FRENTE AL X" / "A UN COSTADO DE X" que ya viene acotada.
    m = re.search(
        r"\b(?:FRENTE A(?:L)?|A UN COSTADO DE(?:L)?|JUNTO A(?:L)?|"
        r"A LA ALTURA DE(?:L)?|CERCA DE(?:L)?|ATRAS DE(?:L)?|AFUERA DE(?:L)?)\s+"
        r"([A-Z0-9NÑ'\.\- ]{3,40}?)"
        r"(?=\s*(?:,|\.|\bY\b|\bEN\b|\bSE\b|\bEL\b|$))",
        texto_norm,
    )
    if m:
        candidato = _limpio(m.group(1))
        if len(candidato) >= 3:
            return candidato

    for poi in sorted(PUNTOS_INTERES, key=len, reverse=True):
        if re.search(r"\b" + re.escape(poi) + r"\b", texto_norm):
            return poi
    return ""


def extraer_de_texto(texto):
    """Extrae componentes de ubicacion de una descripcion en lenguaje natural."""
    t = _norm(texto)
    if not t or t in NULOS:
        return {}
    return {
        "vialidad": _buscar_via(t),
        "colonia": _buscar_asentamiento(t),
        "entre_calles": _buscar_entre_calles(t),
        "numero": _buscar_numero(t),
        "kilometro": _buscar_kilometro(t),
        "punto_interes": _buscar_punto_interes(t),
    }


# --- Construccion de la consulta Windows Maps -------------------------------

def _uri_bing(query, lat=None, lng=None):
    """URI nativo de Windows Maps (protocolo bingmaps:)."""
    if lat is not None and lng is not None:
        return f"bingmaps:?collection=point.{lat}_{lng}_{quote(query)}&lvl=17"
    return f"bingmaps:?where={quote(query)}&lvl=16"


def _uri_ruta(query, lat=None, lng=None):
    """URI de navegacion para despacho de unidad (ms-drive-to:)."""
    if lat is not None and lng is not None:
        return (
            f"ms-drive-to:?destination.latitude={lat}"
            f"&destination.longitude={lng}&destination.name={quote(query)}"
        )
    return f"ms-drive-to:?destination.name={quote(query)}"


def geocodificar(municipio="", direccion="", referencia="", notas="",
                 lat=None, lng=None, colonia="", calle=""):
    """
    Punto de entrada del extractor.

    Combina, en orden de confiabilidad: campos estructurados > DIRECCION del C5
    > REFERENCIA DE UBICACION > NOTAS de cabina. Devuelve la cadena lista para
    pegar en Windows Maps, los URI nativos y la trazabilidad de que campo
    aporto cada componente.
    """
    fuentes = []
    partes = {
        "vialidad": "", "numero": "", "cruce": "", "entre_calles": [],
        "colonia": "", "punto_interes": "", "kilometro": "",
    }

    def _tomar(origen, datos):
        usado = False
        for clave, valor in datos.items():
            if not valor:
                continue
            if clave == "entre_calles":
                if not partes["entre_calles"]:
                    partes["entre_calles"] = valor
                    usado = True
                continue
            if not partes.get(clave):
                partes[clave] = valor
                usado = True
        if usado and origen not in fuentes:
            fuentes.append(origen)

    # 1. Campos ya estructurados (caso de la tabla de HD corroborados).
    _tomar("campos_estructurados", {
        "vialidad": _titulo_via(calle) if not _vacio(calle) else "",
        "colonia": _titulo_via(colonia) if not _vacio(colonia) else "",
    })

    # 2. Campo DIRECCION del C5 (";CALLE|COLONIA").
    _tomar("direccion_c5", parse_direccion_c5(direccion))

    # 3. Referencia de ubicacion capturada por la cabina.
    _tomar("referencia_c5", extraer_de_texto(referencia))

    # 4. Notas de cabina en lenguaje natural (el caso dificil del apunte).
    _tomar("notas_nlp", extraer_de_texto(notas))

    municipio_limpio = _limpio(municipio)
    if _vacio(municipio_limpio):
        municipio_limpio = ""

    # --- Armado de la cadena de consulta ---
    tokens = []
    via = partes["vialidad"]
    if via:
        if partes["numero"]:
            via = f"{via} {partes['numero']}"
        if partes["kilometro"]:
            via = f"{via} {partes['kilometro']}"
        if partes["cruce"]:
            via = f"{via} esquina {partes['cruce']}"
        tokens.append(via)
    elif partes["kilometro"]:
        tokens.append(partes["kilometro"])

    if not via and partes["entre_calles"]:
        tokens.append(" y ".join(partes["entre_calles"]))

    if partes["punto_interes"] and not via:
        tokens.insert(0, partes["punto_interes"])

    # La exportacion C5 repite el municipio en el campo de colonia; no duplicar.
    colonia_util = bool(
        partes["colonia"] and _norm(partes["colonia"]) != _norm(municipio_limpio)
    )
    if colonia_util:
        tokens.append(partes["colonia"])
    if municipio_limpio:
        tokens.append(municipio_limpio)
    tokens.append(ESTADO)
    tokens.append(PAIS)

    query = ", ".join(t for t in tokens if t)

    # --- Nivel de confianza ---
    tiene_via = bool(via)
    tiene_col = colonia_util
    tiene_ancla = bool(partes["punto_interes"] or partes["entre_calles"])

    if lat is not None and lng is not None:
        confianza = "COORDENADA"
    elif tiene_via and tiene_col and municipio_limpio:
        confianza = "ALTA"
    elif municipio_limpio and (tiene_via or tiene_col or tiene_ancla):
        confianza = "MEDIA"
    elif tiene_via and tiene_col:
        # Sin municipio capturado, pero calle + colonia acotan lo suficiente
        # para que Maps resuelva dentro del Estado de Mexico.
        confianza = "MEDIA"
    elif municipio_limpio or tiene_via or tiene_col or tiene_ancla:
        confianza = "BAJA"
    else:
        confianza = "NULA"

    return {
        "query": query,
        "confianza": confianza,
        "fuentes": fuentes,
        "componentes": {
            "vialidad": partes["vialidad"],
            "numero": partes["numero"],
            "cruce": partes["cruce"],
            "entre_calles": partes["entre_calles"],
            "colonia": partes["colonia"],
            "punto_interes": partes["punto_interes"],
            "kilometro": partes["kilometro"],
            "municipio": municipio_limpio,
        },
        "uri_windows_maps": _uri_bing(query, lat, lng),
        "uri_ruta_despacho": _uri_ruta(query, lat, lng),
    }


if __name__ == "__main__":  # verificacion rapida
    casos = [
        dict(municipio="TEMOAYA", direccion=";GOMEZ FARIAS|TEMOAYA",
             referencia="EN LA GLORIETA , POR EL CHARCO", notas=""),
        dict(municipio="ZINACANTEPEC",
             direccion=";AV LA MAGDALENA / FRANCISCO I MADERO|CENTRO",
             referencia="FRENTE AL 3B", notas=""),
        dict(municipio="ECATEPEC", direccion="", referencia="",
             notas="USUARIO REPORTA DETONACIONES SOBRE CALLE MORELOS ENTRE "
                   "HIDALGO Y ZARAGOZA, COLONIA SAN AGUSTIN, FRENTE AL OXXO"),
        dict(municipio="", direccion="", referencia="", notas="LLAMADA MUDA"),
    ]
    for c in casos:
        r = geocodificar(**c)
        print(f"[{r['confianza']:>10}] {r['query']}")
        print(f"             fuentes={r['fuentes']}")
