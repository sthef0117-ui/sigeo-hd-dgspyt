"""
SIGEO-HD DGSPYT — Bandeja de información.

Deja cualquier cosa revuelta en insumos/bandeja/ (Excel, capturas de WhatsApp,
Word, PDF, notas) y este organizador la lee, dice que es, extrae lo que puede
y la vincula al homicidio que le corresponde.

    python src/bandeja.py

Salidas:
    insumos/bandeja/CATALOGO.md   catalogo legible, agrupado por caso
    analisis/bandeja.json         catalogo estructurado (uso interno)

No borra ni mueve los archivos originales: solo los cataloga. Todo lo que
contenga datos personales se queda en la bandeja, que no se publica.

Principio del proyecto: simple de entender, preciso en el dato.
"""

import json
import re
import sys
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
BANDEJA = RAIZ / "insumos" / "bandeja"
ANALISIS = RAIZ / "analisis"

IGNORAR = {"LEEME.txt", "CATALOGO.md"}
EXT_EXCEL = {".xlsx", ".xls", ".xlsm"}
EXT_IMG = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"}
EXT_TXT = {".txt", ".csv", ".md"}


def norm(t):
    s = unicodedata.normalize("NFD", str(t or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.upper().split())


def municipios_conocidos():
    p = ANALISIS / "perfil_territorial.json"
    if not p.exists():
        return {}
    datos = json.loads(p.read_text(encoding="utf-8"))
    return {norm(t["municipio"]): t["municipio"] for t in datos}


def homicidios():
    p = ANALISIS / "corroborados_sigeo.json"
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else []


# --- Extraccion de senales de un texto o nombre de archivo ------------------

_RE_FOLIO = re.compile(r"\b([A-Z]{3}\d{11})\b")
_RE_FECHA = re.compile(r"\b(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\b"
                       r"|\b(\d{1,2})[-/.](\d{1,2})[-/.](20\d{2})\b")
_RE_FECHA_WA = re.compile(r"(20\d{2})(\d{2})(\d{2})")  # IMG-20260723-...
_RE_HDID = re.compile(r"\bHD[- ]?(\d{1,4})\b", re.IGNORECASE)


def fechas_en(texto):
    salida = set()
    for m in _RE_FECHA.finditer(texto):
        if m.group(1):
            salida.add(f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}")
        else:
            salida.add(f"{m.group(6)}-{int(m.group(5)):02d}-{int(m.group(4)):02d}")
    for m in _RE_FECHA_WA.finditer(texto):
        a, mes, d = m.group(1), int(m.group(2)), int(m.group(3))
        if 1 <= mes <= 12 and 1 <= d <= 31:
            salida.add(f"{a}-{mes:02d}-{d:02d}")
    return salida


def senales(texto, munis):
    t = norm(texto)
    return {
        "folios": sorted(set(_RE_FOLIO.findall(t))),
        "fechas": sorted(fechas_en(texto)),
        "hd_ids": sorted({int(x) for x in _RE_HDID.findall(texto)}),
        "municipios": sorted({munis[k] for k in munis if k in t}),
    }


# --- Lectores por tipo ------------------------------------------------------

def leer_excel(ruta):
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    hojas, muestra = [], ""
    for s in wb.sheetnames:
        ws = wb[s]
        hojas.append(f"{s} ({ws.max_row}f)")
        for fila in ws.iter_rows(min_row=1, max_row=6, values_only=True):
            muestra += " ".join(str(v) for v in fila if v is not None) + "\n"
    wb.close()
    tipo = ("Homicidios" if "DESARROLLO DE LOS HECHOS" in norm(muestra)
            or "TOTAL DE HOMICIDIOS" in norm(muestra)
            else "Llamadas 911" if "INCIDENTE" in norm(muestra) and "FOLIO" in norm(muestra)
            else "Tabla de datos")
    return f"Excel · {tipo} · hojas: {', '.join(hojas[:4])}", muestra[:4000]


def leer_texto(ruta):
    txt = ruta.read_text(encoding="utf-8", errors="ignore")
    return f"Texto · {len(txt)} caracteres", txt[:4000]


def leer_docx(ruta):
    from docx import Document
    doc = Document(ruta)
    txt = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    return f"Word · {len(doc.paragraphs)} párrafos", txt[:4000]


def leer_pdf(ruta):
    from pypdf import PdfReader
    r = PdfReader(str(ruta))
    txt = ""
    for pg in r.pages[:6]:
        try:
            txt += (pg.extract_text() or "") + "\n"
        except Exception:
            pass
    marca = f"PDF · {len(r.pages)} páginas"
    if not txt.strip():
        marca += " · sin texto (posible escaneo, requiere revisión manual)"
    return marca, txt[:4000]


def _ocr_disponible():
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def leer_imagen(ruta):
    """
    Las imágenes son evidencia reservada: su contenido gráfico NUNCA se muestra
    ni se publica. Si el equipo tiene Tesseract instalado, se lee el texto de la
    captura (tarjetas de WhatsApp) solo para extraer municipio, fecha o folio y
    poder vincularla; ese texto se usa internamente, no se vuelca al catálogo.
    """
    from PIL import Image
    try:
        with Image.open(ruta) as im:
            dim = f"{im.width}×{im.height}"
            texto = ""
            if _ocr_disponible():
                import pytesseract
                try:
                    texto = pytesseract.image_to_string(im, lang="spa")
                except Exception:
                    texto = pytesseract.image_to_string(im)
    except Exception:
        return "Imagen (no legible) · evidencia reservada", ""
    if texto.strip():
        return f"Imagen · {dim} · texto leído por OCR · evidencia reservada", texto + " " + ruta.name
    return (f"Imagen · {dim} · sin OCR (instale Tesseract para leerla) · "
            "evidencia reservada", ruta.name)


def clasificar(ruta):
    ext = ruta.suffix.lower()
    try:
        if ext in EXT_EXCEL:
            return leer_excel(ruta)
        if ext in EXT_IMG:
            return leer_imagen(ruta)
        if ext == ".docx":
            return leer_docx(ruta)
        if ext == ".pdf":
            return leer_pdf(ruta)
        if ext in EXT_TXT:
            return leer_texto(ruta)
    except Exception as e:
        return f"No se pudo leer ({e.__class__.__name__})", ""
    return f"Tipo no reconocido ({ext or 'sin extensión'})", ""


# --- Vinculacion al homicidio -----------------------------------------------

def vincular(sig, hd):
    """
    Sugiere el homicidio que empata con las senales extraidas.

    Regla de precision: un id explicito manda; si no, se exige municipio Y
    fecha juntos (un solo dato empata con demasiados casos). Un documento que
    menciona muchos municipios o fechas es un concentrado general, no la
    evidencia de un caso, y no se fuerza a ninguno.
    """
    if sig["hd_ids"]:
        directos = [h for h in hd if h["id"] in sig["hd_ids"]]
        if directos:
            return [(h, "id explícito") for h in directos], "vinculado"

    masivo = len(sig["municipios"]) > 3 or len(sig["fechas"]) > 5
    if masivo:
        return [], "documento general"

    munis = {norm(m) for m in sig["municipios"]}
    if munis and sig["fechas"]:
        cand = [(h, "municipio + fecha") for h in hd
                if norm(h["municipio"]) in munis and h["fecha"] in sig["fechas"]]
        if cand:
            return cand[:6], "vinculado"

    # Senales claras pero sin homicidio corroborado que empate: puede ser un
    # hecho aun no corroborado o un caso sin carpeta.
    if (munis or sig["fechas"]) and (sig["folios"] or (munis and sig["fechas"])):
        return [], "sin homicidio que empate"

    return [], "sin vincular"


# --- Programa ---------------------------------------------------------------

def main():
    if not BANDEJA.exists():
        BANDEJA.mkdir(parents=True)
    munis = municipios_conocidos()
    hd = homicidios()

    archivos = [p for p in sorted(BANDEJA.rglob("*"))
                if p.is_file() and p.name not in IGNORAR
                and "_organizado" not in p.parts]
    if not archivos:
        print("La bandeja está vacía.")
        print(f"Deja archivos en: {BANDEJA}")
        print("Luego vuelve a correr: python src/bandeja.py")
        return 0

    items = []
    for ruta in archivos:
        desc, contenido = clasificar(ruta)
        sig = senales(contenido + " " + ruta.name, munis)
        vinc, estado = vincular(sig, hd)
        items.append({
            "archivo": ruta.name,
            "ruta": str(ruta.relative_to(RAIZ)),
            "descripcion": desc,
            "estado": estado,
            "senales": sig,
            "vinculos": [{"id": h["id"], "fecha": h["fecha"],
                          "municipio": h["municipio"], "colonia": h["colonia"],
                          "por": razon} for h, razon in vinc],
        })

    # --- Catalogo estructurado (uso interno) ---
    ANALISIS.mkdir(parents=True, exist_ok=True)
    (ANALISIS / "bandeja.json").write_text(
        json.dumps(items, ensure_ascii=False, indent=1), encoding="utf-8")

    # --- Catalogo legible, agrupado ---
    por_caso = {}
    generales = [it for it in items if it["estado"] == "documento general"]
    sin_carpeta = [it for it in items if it["estado"] == "sin homicidio que empate"]
    sueltos = [it for it in items if it["estado"] == "sin vincular"]
    for it in items:
        for v in it["vinculos"][:1]:
            por_caso.setdefault(v["id"], {"info": v, "items": []})
            por_caso[v["id"]]["items"].append(it)
    n_vinc = sum(len(g["items"]) for g in por_caso.values())

    def pistas(it):
        s, p = it["senales"], []
        if s["municipios"]:
            p.append("municipios: " + ", ".join(s["municipios"][:4]))
        if s["fechas"]:
            p.append("fechas: " + ", ".join(s["fechas"][:4]))
        if s["folios"]:
            p.append("folios: " + ", ".join(s["folios"]))
        return (" — " + " · ".join(p)) if p else ""

    L = ["# Catálogo de la bandeja — SIGEO-HD", "",
         f"_Generado {datetime.now().strftime('%Y-%m-%d %H:%M')} · {len(items)} archivo(s)._",
         "",
         f"- Vinculados a un homicidio: **{n_vinc}**",
         f"- Con señales fuertes pero sin homicidio corroborado que empate: **{len(sin_carpeta)}**",
         f"- Documentos generales (concentrados, chats): **{len(generales)}**",
         f"- Sin señales suficientes: **{len(sueltos)}**", ""]

    if por_caso:
        L.append("## Vinculados por caso")
        for hid, g in sorted(por_caso.items()):
            v = g["info"]
            L.append(f"\n### HD-{hid} · {v['municipio']} · {v['fecha']}"
                     f"{' · ' + v['colonia'] if v['colonia'] else ''}")
            for it in g["items"]:
                L.append(f"- **{it['archivo']}** — {it['descripcion']} "
                         f"_(por {it['vinculos'][0]['por']})_")

    if sin_carpeta:
        L.append("\n## Señales fuertes sin homicidio corroborado")
        L.append("_Traen municipio, fecha o folio, pero no empatan con ningún homicidio "
                 "de la tabla. Posible hecho no corroborado o caso sin carpeta: revisar._")
        for it in sin_carpeta:
            L.append(f"- **{it['archivo']}** — {it['descripcion']}{pistas(it)}")

    if generales:
        L.append("\n## Documentos generales")
        L.append("_Mencionan muchos lugares o fechas; son concentrados, no evidencia de un caso._")
        for it in generales:
            L.append(f"- **{it['archivo']}** — {it['descripcion']}")

    if sueltos:
        L.append("\n## Sin señales suficientes")
        L.append("_Falta municipio + fecha o un folio para amarrarlos. Vincular a mano._")
        for it in sueltos:
            L.append(f"- **{it['archivo']}** — {it['descripcion']}{pistas(it)}")

    (BANDEJA / "CATALOGO.md").write_text("\n".join(L) + "\n", encoding="utf-8")

    print("SIGEO-HD DGSPYT · bandeja organizada")
    print(f"  archivos leídos ............ {len(items)}")
    print(f"  vinculados a un caso ....... {n_vinc}")
    print(f"  sin carpeta que empate ..... {len(sin_carpeta)}")
    print(f"  documentos generales ....... {len(generales)}")
    print(f"  sin señales ................ {len(sueltos)}")
    print(f"  catálogo legible ........... {BANDEJA / 'CATALOGO.md'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
