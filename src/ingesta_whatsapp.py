# -*- coding: utf-8 -*-
"""
SIGEO-HD DGSPYT — Ingesta de tarjetas informativas de WhatsApp.

Las tarjetas de WhatsApp son HOMICIDIOS REALES REPORTADOS, no solo texto: cada
mensaje del chat institucional es una tarjeta de un hecho. Este módulo abre el
export del chat (.zip/.txt), lo parte en tarjetas, las vuelve eventos
estructurados con tarjetas.parsear_tarjeta y los clasifica (HD por arma vs otras
causas a revisar). El resultado alimenta el buscador y las fichas del tablero.

    python src/ingesta_whatsapp.py

Salida:
    analisis/tarjetas_whatsapp.json   eventos estructurados (uso del tablero)
    entregables/TARJETAS_WHATSAPP.md  lista legible para revisión
"""
import json
import re
import sys
import unicodedata
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
INSUMOS = RAIZ / "insumos"
ANALISIS = RAIZ / "analisis"
ENTREGABLES = RAIZ / "entregables"
sys.path.insert(0, str(Path(__file__).parent))
from tarjetas import parsear_tarjeta, _municipios_conocidos  # noqa: E402

# Inicio de cada mensaje del export de WhatsApp:
#   13/08/26 1:30 p. m. - Remitente: mensaje
_RE_MSG = re.compile(
    r"(\d{1,2}/\d{1,2}/\d{2,4})\s+\d{1,2}:\d{2}\s*[ap]\.?\s*m\.?\s*-\s*(.*?):\s",
    re.IGNORECASE)

_MUERTE = ("PERSONA MUERTA", "OCCISO", "OCCISA", "SIN VIDA", "IMPACTO DE BALA",
           "PRIVAD", "HOMICIDIO", "CADAVER", "CUERPO")


def _norm(t):
    return "".join(c for c in unicodedata.normalize("NFD", str(t))
                   if unicodedata.category(c) != "Mn").upper()


def _causa_titulo(texto):
    """Extrae lo que va entre paréntesis del encabezado: '(… POR IMPACTO DE BALA)'."""
    m = re.search(r"\(([^)]*(?:MUERT|OCCIS|VIDA|BALA|HOMICID|PRIVAD)[^)]*)\)",
                  texto, re.IGNORECASE)
    return m.group(1).strip() if m else ""


def _clasificar(texto, arma):
    t = _norm(texto)
    if arma in ("Arma de fuego", "Arma blanca"):
        return "HOMICIDIO DOLOSO"
    if "ATROPELL" in t or "ACCIDENTE" in t or "VOLCAD" in t or "ARROLL" in t:
        return "NO HD — hecho de tránsito"
    if "SUICID" in t or "AHORCA" in t or "COLGAD" in t:
        return "NO HD — probable suicidio"
    if "CAUSAS A DETERMINAR" in t or "NATURAL" in t or "AL PARECER NATURAL" in t:
        return "REVISAR — causa a determinar"
    return "REVISAR — sin arma clara"


def separar_mensajes(txt):
    """Corta el chat en (fecha, remitente, cuerpo) por cada marca de mensaje."""
    marcas = list(_RE_MSG.finditer(txt))
    for i, m in enumerate(marcas):
        ini = m.end()
        fin = marcas[i + 1].start() if i + 1 < len(marcas) else len(txt)
        cuerpo = txt[ini:fin].strip()
        yield m.group(1), m.group(2).strip(), cuerpo


def leer_chats():
    textos = []
    for z in INSUMOS.rglob("*.zip"):
        try:
            zf = zipfile.ZipFile(z)
            for n in zf.namelist():
                if n.lower().endswith(".txt"):
                    textos.append((z.name, zf.read(n).decode("utf-8", "ignore")))
        except Exception:
            pass
    for p in INSUMOS.rglob("*.txt"):
        if p.name.lower() != "leeme.txt":
            textos.append((p.name, p.read_text(encoding="utf-8", errors="ignore")))
    return textos


def main():
    munis = _municipios_conocidos()
    eventos = []
    for origen, txt in leer_chats():
        for fecha_msg, remitente, cuerpo in separar_mensajes(txt):
            tu = _norm(cuerpo)
            # es tarjeta de un hecho fatal si menciona muerte y trae encabezado
            if not any(k in tu for k in _MUERTE):
                continue
            if len(cuerpo) < 120:  # descarta menciones sueltas, no tarjetas
                continue
            ev = parsear_tarjeta(cuerpo, munis)
            ev["causa_reportada"] = _causa_titulo(cuerpo)
            ev["clasificacion"] = _clasificar(cuerpo, ev["arma"])
            ev["fecha_mensaje"] = fecha_msg
            ev["origen"] = origen
            eventos.append(ev)

    # dedup por (municipio, lugar, narrativa[:80])
    vistos, unicos = set(), []
    for e in eventos:
        clave = (_norm(e["municipio"]), _norm(e["lugar"])[:40], _norm(e["narrativa"])[:80])
        if clave in vistos:
            continue
        vistos.add(clave)
        unicos.append(e)

    ANALISIS.mkdir(parents=True, exist_ok=True)
    (ANALISIS / "tarjetas_whatsapp.json").write_text(
        json.dumps(unicos, ensure_ascii=False, indent=1), encoding="utf-8")

    hd = [e for e in unicos if e["clasificacion"] == "HOMICIDIO DOLOSO"]
    rev = [e for e in unicos if e["clasificacion"].startswith("REVISAR")]
    nohd = [e for e in unicos if e["clasificacion"].startswith("NO HD")]

    L = ["# Tarjetas de WhatsApp — hechos reportados", "",
         f"Total de tarjetas leídas: **{len(unicos)}**", "",
         f"- Homicidios dolosos (por arma): **{len(hd)}**",
         f"- A revisar (causa a determinar / sin arma clara): **{len(rev)}**",
         f"- No HD (tránsito / suicidio): **{len(nohd)}**", "",
         "## Homicidios dolosos", ""]
    for e in hd:
        L.append(f"- **{e['municipio'] or '¿?'}** · {e['arma']} · "
                 f"{('víctima '+str(e['victima_edad'])+' años') if e['victima_edad'] else 'edad s/d'}"
                 f"{' · '+e['lugar'] if e['lugar'] else ''}  \n"
                 f"  _{e['movil']}_ · causa: {e['causa_reportada'] or '—'}")
    L += ["", "## A revisar", ""]
    for e in rev:
        L.append(f"- {e['municipio'] or '¿?'} · causa: {e['causa_reportada'] or '—'}"
                 f"{' · '+e['lugar'] if e['lugar'] else ''}")
    L += ["", "## No HD (otra competencia)", ""]
    for e in nohd:
        L.append(f"- {e['municipio'] or '¿?'} · {e['clasificacion']} · causa: {e['causa_reportada'] or '—'}")

    ENTREGABLES.mkdir(parents=True, exist_ok=True)
    (ENTREGABLES / "TARJETAS_WHATSAPP.md").write_text("\n".join(L) + "\n", encoding="utf-8")

    print("SIGEO-HD DGSPYT - ingesta de WhatsApp")
    print(f"  tarjetas leidas ............ {len(unicos)}")
    print(f"  homicidios dolosos ......... {len(hd)}")
    print(f"  a revisar .................. {len(rev)}")
    print(f"  no HD (transito/suicidio) .. {len(nohd)}")
    print(f"  JSON ....................... {ANALISIS / 'tarjetas_whatsapp.json'}")
    print(f"  lista legible .............. {ENTREGABLES / 'TARJETAS_WHATSAPP.md'}")
    almoloya = [e for e in unicos if "ALMOLOYA" in _norm(e["municipio"])]
    print(f"  ejercicio ALMOLOYA ......... {'ENCONTRADO' if almoloya else 'no'} "
          f"({len(almoloya)} tarjeta/s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
