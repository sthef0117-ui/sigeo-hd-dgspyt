"""
SIGEO-HD DGSPYT — Lector de tarjetas informativas.

Convierte el texto de una tarjeta de WhatsApp (leida por OCR o pegada) en un
evento estructurado y ANONIMIZADO. Tu das el insumo crudo; el sistema saca el
dato limpio. La captura no se guarda como grafica: se limpia a texto.

    from tarjetas import parsear_tarjeta
    evento = parsear_tarjeta(texto_de_la_tarjeta)

El evento resultante alimenta la ficha de inteligencia igual que un homicidio
corroborado, pero marcado como "no corroborado" hasta que entre a la tabla.

Datos personales (nombres, edades de terceros, placas) se suprimen aqui mismo.
La edad de la victima se conserva por ser dato estadistico.
"""

import re
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from anonimizar import anonimizar_texto  # noqa: E402

MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
         "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10,
         "noviembre": 11, "diciembre": 12}


def _norm(t):
    s = unicodedata.normalize("NFD", str(t or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").upper()


def _municipios_conocidos():
    import json
    p = Path(__file__).resolve().parent.parent / "analisis" / "perfil_territorial.json"
    if not p.exists():
        return {}
    datos = json.loads(p.read_text(encoding="utf-8"))
    return {_norm(t["municipio"]): t["municipio"] for t in datos}


def _municipio(texto, munis):
    t = _norm(texto)
    # La tarjeta suele abrir con el municipio en mayusculas: "ALMOLOYA DE JUAREZ (..."
    ini = re.match(r"^\s*([A-ZÁÉÍÓÚÑ .]{4,40}?)\s*\(", texto)
    if ini:
        cand = _norm(ini.group(1)).strip()
        # recortar sufijos como en el pipeline
        for suf in (" DE JUAREZ", " DE MORELOS", " DE ZARAGOZA", " DE BAZ"):
            if cand.endswith(suf) and cand not in munis:
                base = cand[: -len(suf)].strip()
                if base in munis:
                    return munis[base]
        if cand in munis:
            return munis[cand]
    # si no, el primer municipio conocido que aparezca
    for k in sorted(munis, key=len, reverse=True):
        if re.search(r"\b" + re.escape(k) + r"\b", t):
            return munis[k]
    return ""


def _lugar(texto):
    m = re.search(r"\b(?:en|sobre|ubicad[oa] en)\s+((?:priv(?:ada)?|calle|av(?:enida)?|"
                  r"carretera|camino|cerrada|blvd|boulevard|fraccionamiento|colonia|"
                  r"barrio|pueblo)[^,\.]{3,70})", texto, re.IGNORECASE)
    partes = []
    for m in re.finditer(r"\b((?:Privada|Priv\.?|Calle|Avenida|Av\.?|Carretera|Cerrada|"
                         r"Boulevard|Blvd\.?|Fraccionamiento|Fracc\.?|Colonia|Col\.?|"
                         r"Barrio|Pueblo|Ejido)\s+[A-ZÁÉÍÓÚa-záéíóúÑñ0-9 ]{3,45})",
                         texto):
        p = re.sub(r"\s+", " ", m.group(1)).strip(" ,.")
        if p not in partes:
            partes.append(p)
    return ", ".join(partes[:3])


def _fecha(texto):
    m = re.search(r"\b(\d{1,2})\s+de\s+([a-záéíóú]+)\s+(?:de\s+)?(20\d{2})\b",
                  texto, re.IGNORECASE)
    if m and m.group(2).lower() in MESES:
        return f"{m.group(3)}-{MESES[m.group(2).lower()]:02d}-{int(m.group(1)):02d}"
    m = re.search(r"\b(20\d{2})[-/](\d{1,2})[-/](\d{1,2})\b", texto)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.search(r"\b(\d{1,2})[/](\d{1,2})[/](20\d{2})\b", texto)
    if m:
        return f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    return ""


def _horas(texto):
    return re.findall(r"\b([01]?\d|2[0-3]):([0-5]\d)\s*(?:hrs?|horas)?\b", texto)


def _arma(texto):
    t = _norm(texto)
    if re.search(r"IMPACTO DE BALA|ARMA DE FUEGO|PROYECTIL|DISPARO|BALAZO|CALIBRE", t):
        return "Arma de fuego"
    if re.search(r"ARMA BLANCA|CUCHILL|NAVAJA|MACHETE|PUNZOCORTANTE|APUÑAL", t):
        return "Arma blanca"
    if re.search(r"GOLPE|CONTUNDENTE|ASFIXIA|ESTRANGUL", t):
        return "Golpes / asfixia"
    return "No determinada"


def _victima(texto):
    edad = None
    m = re.search(r"de\s+(\d{1,3})\s+a[nñ]os", texto, re.IGNORECASE)
    if m:
        edad = int(m.group(1))
    t = _norm(texto)
    sexo = ("Masculino" if re.search(r"\bMASCULIN|\bHOMBRE|MUERTO\b|OCCISO", t)
            else "Femenino" if re.search(r"\bFEMENIN|\bMUJER|MUERTA\b|OCCISA", t)
            else "No determinado")
    return edad, sexo


def _agresores(texto):
    m = re.search(r"\b(\d{1,2}|un|dos|tres|cuatro)\s+(?:sujetos?|personas?|"
                  r"individuos?|masculinos?)\b", texto, re.IGNORECASE)
    if not m:
        return None
    palabra = {"un": 1, "dos": 2, "tres": 3, "cuatro": 4}
    v = m.group(1).lower()
    return int(v) if v.isdigit() else palabra.get(v)


def _movil(texto):
    t = _norm(texto)
    if re.search(r"INGIRIENDO BEBIDAS|EMBRIAGANT|RI[NÑ]A|DISCUSION|PELEA", t):
        return "Posible riña / entre conocidos"
    if re.search(r"AJUSTE DE CUENTAS|EJECUCION|LEVANTAD|CELUL[AO]|NARCOMENSAJE", t):
        return "Posible ajuste de cuentas"
    if re.search(r"ROBO|ASALT|DESPOJO", t):
        return "Posible robo"
    return "Se desconoce el móvil de la agresión"


def parsear_tarjeta(texto, munis=None):
    """Texto crudo de una tarjeta -> evento estructurado y anonimizado."""
    if munis is None:
        munis = _municipios_conocidos()
    edad, sexo = _victima(texto)
    horas = _horas(texto)
    return {
        "municipio": _municipio(texto, munis),
        "lugar": _lugar(texto),
        "fecha": _fecha(texto),
        "hora": (f"{int(horas[-1][0]):02d}:{horas[-1][1]}:00" if horas else ""),
        "horas_mencionadas": [f"{int(h):02d}:{m}" for h, m in horas],
        "arma": _arma(texto),
        "victima_edad": edad,
        "victima_sexo": sexo,
        "num_agresores": _agresores(texto),
        "movil": _movil(texto),
        "narrativa": anonimizar_texto(texto),
        "corroborado": False,
        "fuente": "Tarjeta informativa (WhatsApp)",
    }


if __name__ == "__main__":  # prueba con el texto del ejercicio
    ejercicio = (
        "ALMOLOYA DE JUÁREZ (PERSONA MUERTA POR IMPACTO DE BALA) Siendo las 20:37 "
        "horas, el Policía PELCASTRE PÉREZ JUAN CARLOS, en la unidad ME490A1, informó "
        "que a las 06:57 horas, hizo contacto con el Primer Respondiente Policía "
        "Municipal SOTO ZAVALETA FRAIN ANTONIO, en la unidad AJ-089, indicando que en "
        "Privada Cerro Tláloc, Fraccionamiento Colinas del Sol, resultó muerto por "
        "impacto de bala en el pecho, quien en vida respondía al nombre de JUAN DANIEL "
        "RODRÍGUEZ ÁVILA de 27 años, datos proporcionados por su novia MELANIA VANESA "
        "GINES GUTIÉRREZ de 19 años, manifestando que los responsables fueron 2 sujetos "
        "con los que se encontraban ingiriendo bebidas embriagantes, dándose a la fuga."
    )
    import json
    print(json.dumps(parsear_tarjeta(ejercicio), ensure_ascii=False, indent=1))
