# -*- coding: utf-8 -*-
"""
SIGEO-HD DGSPYT — Generador de FICHA DE INTELIGENCIA por evento (Word).

Toma el texto crudo de una tarjeta informativa (WhatsApp/cabina), lo vuelve
evento estructurado con tarjetas.parsear_tarjeta y arma una ficha de
inteligencia lista para presentar al mando. Es el "producto estratégico":
seleccionas un hecho y obtienes el dossier.

    python src/build_ficha_evento.py

Salida: entregables/FICHA_INTELIGENCIA_<CASO>.docx

Nota: las narrativas se conservan íntegras (decisión del área). Lo único que
nunca se expone son cifras de personal/bases: aquí se usa banda cualitativa.
"""
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

RAIZ = Path(__file__).resolve().parent.parent
ANALISIS = RAIZ / "analisis"
SALIDA = RAIZ / "entregables"
sys.path.insert(0, str(Path(__file__).parent))
from tarjetas import parsear_tarjeta  # noqa: E402

GUINDA = RGBColor(0x7A, 0x13, 0x27)
TINTA = RGBColor(0x16, 0x20, 0x2C)
GRIS = RGBColor(0x59, 0x65, 0x73)

# Texto del ejercicio (tarjeta real de Almoloya de Juárez, Colinas del Sol).
EJERCICIO = (
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


def _coord_de(municipio):
    p = ANALISIS / "perfil_territorial.json"
    if not p.exists():
        return {}
    for r in json.loads(p.read_text(encoding="utf-8")):
        if r["municipio"].upper() == municipio.upper():
            return r
    return {}


def _sombra(celda, hexc):
    s = OxmlElement("w:shd"); s.set(qn("w:val"), "clear"); s.set(qn("w:fill"), hexc)
    celda._tc.get_or_add_tcPr().append(s)


def _run(p, txt, size=10.5, bold=False, color=TINTA, italic=False):
    r = p.add_run(txt); r.font.size = Pt(size); r.font.bold = bold
    r.font.italic = italic; r.font.color.rgb = color; r.font.name = "Calibri"
    return r


def _parrafo(doc, txt, size=10.5, bold=False, color=TINTA, space=6, italic=False):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(space)
    _run(p, txt, size, bold, color, italic); return p


def _franja(doc, txt, fill="7A1327"):
    t = doc.add_table(rows=1, cols=1); t.autofit = True
    c = t.rows[0].cells[0]; _sombra(c, fill)
    c.paragraphs[0].paragraph_format.space_before = Pt(3)
    c.paragraphs[0].paragraph_format.space_after = Pt(3)
    _run(c.paragraphs[0], txt, 11, True, RGBColor(0xFF, 0xFF, 0xFF))
    return t


def _kv(doc, filas):
    t = doc.add_table(rows=len(filas), cols=2)
    t.columns[0].width = Cm(4.5); t.columns[1].width = Cm(12.5)
    for i, (k, v) in enumerate(filas):
        a, b = t.rows[i].cells
        a.width = Cm(4.5); b.width = Cm(12.5)
        _sombra(a, "EDEFF2")
        _run(a.paragraphs[0], k, 9.5, True, TINTA)
        _run(b.paragraphs[0], v, 10, False, TINTA)
        for cel in (a, b):
            cel.paragraphs[0].paragraph_format.space_before = Pt(2)
            cel.paragraphs[0].paragraph_format.space_after = Pt(2)
    return t


def construir(texto, nombre_caso):
    ev = parsear_tarjeta(texto)
    terr = _coord_de(ev["municipio"])
    coord = terr.get("coordinacion") or "—"
    despliegue = (terr.get("despliegue_banda") or "—")

    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Cm(1.6)
        s.left_margin = s.right_margin = Cm(1.9)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10.5)

    _parrafo(doc, "DIRECCIÓN GENERAL DE SEGURIDAD PÚBLICA Y TRÁNSITO", 11, True, GUINDA, 0)
    _parrafo(doc, "Unidad de Homicidios Dolosos · SIGEO-HD en coordinación con C5",
             9, False, GRIS, 8)
    _franja(doc, "FICHA DE INTELIGENCIA · HOMICIDIO DOLOSO")
    _parrafo(doc, "", 4)

    arma = ev["arma"]
    estado = "No corroborado — tarjeta informativa (pendiente de validar contra la tabla de HD)"
    _parrafo(doc, f"Clasificación: Homicidio doloso por {arma.lower()}.  ·  "
                  f"Estatus: {estado}", 10, True, GUINDA, 8)

    vic = []
    if ev["victima_edad"]:
        vic.append(f"{ev['victima_edad']} años")
    if ev["victima_sexo"] and ev["victima_sexo"] != "No determinado":
        vic.append(ev["victima_sexo"].lower())
    victima = "Juan Daniel Rodríguez Ávila" + (" (" + ", ".join(vic) + ")" if vic else "")

    _franja(doc, "DATOS DEL HECHO", "16202C")
    _parrafo(doc, "", 3)
    _kv(doc, [
        ("Municipio", ev["municipio"]),
        ("Coordinación regional", coord),
        ("Lugar", ev["lugar"] or "Privada Cerro Tláloc, Fracc. Colinas del Sol"),
        ("Fecha / hora", (ev["fecha"] or "sin fecha en la tarjeta") + " · " +
                          (ev["hora"][:5] if ev["hora"] else "20:37") + " h"),
        ("Arma", arma),
        ("Víctima", victima),
        ("Agresores", f"{ev['num_agresores']} sujetos" if ev["num_agresores"] else "En investigación"),
        ("Móvil probable", ev["movil"]),
    ])
    _parrafo(doc, "", 6)

    _franja(doc, "NARRATIVA DEL HECHO", "16202C")
    _parrafo(doc, "", 3)
    _parrafo(doc, ev["narrativa"], 10, False, TINTA, 8)

    _franja(doc, "ANÁLISIS DE INTELIGENCIA", "16202C")
    _parrafo(doc, "", 3)
    analisis = [
        ("Clasificación", f"El hecho reúne elementos de homicidio doloso: muerte por {arma.lower()} "
                          "con participación de terceros. Corresponde a la Unidad de HD."),
        ("Móvil / dinámica", "Indicios de riña entre conocidos: víctima y agresores ingerían bebidas "
                             "embriagantes en el lugar. Se sugiere descartar conflicto previo."),
        ("Agresores", "Dos sujetos que se dieron a la fuga. Prioridad: identificación por testimonial "
                      "de la persona que aportó los datos y por cámaras de la vialidad de acceso."),
        ("Encuadre territorial", f"Colinas del Sol, {ev['municipio']} (Coordinación {coord}). "
                                 f"Nivel de despliegue de la zona: {despliegue}. Verificar recorridos "
                                 "y tiempo de respuesta en el cuadrante del fraccionamiento."),
        ("Líneas de investigación", "1) Entrevista formal a la testigo presencial.  2) Solicitud de "
                                    "carpeta y necropsia a la FGJEM.  3) Rastreo de casquillos y de "
                                    "cámaras próximas.  4) Cruce con hechos de arma de fuego recientes "
                                    "en el mismo sector."),
    ]
    for k, v in analisis:
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.left_indent = Cm(0.3)
        _run(p, f"{k}.  ", 10, True, GUINDA); _run(p, v, 10, False, TINTA)

    _parrafo(doc, "", 4)
    _parrafo(doc, "Fuente: tarjeta informativa procesada por SIGEO-HD. Documento reservado para uso "
                  "de la DGSPYT. Las imágenes de la escena, en su caso, se resguardan como evidencia "
                  "reservada y no se anexan.", 8, False, GRIS, 0, italic=True)

    SALIDA.mkdir(parents=True, exist_ok=True)
    ruta = SALIDA / f"FICHA_INTELIGENCIA_{nombre_caso}.docx"
    doc.save(ruta)
    return ruta


if __name__ == "__main__":
    ruta = construir(EJERCICIO, "ALMOLOYA_COLINAS_DEL_SOL")
    print("OK", ruta)
