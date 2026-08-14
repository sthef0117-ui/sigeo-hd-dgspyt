"""
SIGEO-HD DGSPYT — Recorte de la foto real en una captura de WhatsApp.

Una captura de WhatsApp trae tres cosas mezcladas: la burbuja de texto (la
tarjeta informativa), el fondo oscuro del chat y —abajo— la FOTO real de la
escena. Aquí solo nos interesa aislar esa foto, recortada, para tratarla como
evidencia reservada. El texto no se recorta: ese se lee por OCR aparte.

    from recorte import recortar_foto
    r = recortar_foto("captura.png", "foto_escena.png")
    # r["recortada"] == True si logró aislar la foto con confianza

Cómo distingue la foto del resto, sin librerías de visión ni modelos:

  - La foto es una región con MUCHA textura (varianza de luz alta) Y color
    (una escena real trae rojos, verdes, naranjas), en un bloque macizo.
  - La burbuja de texto es gris plano con letras finas: textura media, color
    casi nulo.
  - El fondo del chat es plano y oscuro: textura y color casi nulos.

Se puntea una rejilla, se marcan las celdas "fotográficas", se toma el bloque
macizo más grande y se recorta su rectángulo. Las capturas son de baja calidad
y a veces desenfocadas, así que si NO logra aislar un bloque con confianza,
NO recorta a ciegas: devuelve recortada=False y conserva la captura completa,
que se sigue tratando como reservada. Vale más una foto de más que perder la
escena por un recorte mal hecho.
"""

import sys
from pathlib import Path

from PIL import Image, ImageMath

ANCHO_ANALISIS = 360      # se analiza a esta anchura; suficiente y rápido
CELDA_PX = 6              # lado de celda en la imagen de análisis
MARGEN = 0.015           # margen que se deja alrededor de la foto al recortar


def _rejilla(imr):
    """Devuelve (columnas, filas, textura[], color[]) de la imagen reducida."""
    w, h = imr.size
    gw, gh = max(8, w // CELDA_PX), max(8, h // CELDA_PX)

    ev = ImageMath.unsafe_eval
    lf = imr.convert("L").convert("F")
    media = lf.resize((gw, gh), Image.BOX)
    media_sq = ev("a*a", a=lf).resize((gw, gh), Image.BOX)

    r, g, b = (c.convert("F") for c in imr.split()[:3])
    color = ev("max(max(a,b),c)-min(min(a,b),c)", a=r, b=g, c=b) \
        .resize((gw, gh), Image.BOX)

    m = list(media.getdata())
    msq = list(media_sq.getdata())
    col = list(color.getdata())
    textura = [max(0.0, msq[i] - m[i] * m[i]) ** 0.5 for i in range(len(m))]
    return gw, gh, textura, col


def _pct(valores, p):
    if not valores:
        return 1.0
    s = sorted(valores)
    return s[min(len(s) - 1, int(p * len(s)))] or 1.0


COLOR_MIN = 3.0      # una foto real trae color (max-min RGB); el texto gris no
TEXTURA_MIN = 5.0    # y trae textura; el fondo plano del chat no


def _celdas_foto(gw, gh, textura, color):
    """
    Marca cada celda como fotográfica. El color es el discriminador limpio:
    medido en las capturas reales, la burbuja de texto da color ~1.6 y el
    fondo del chat ~1.1, mientras la foto de la escena da ~10 con picos de 60.
    Un umbral absoluto de color separa la foto del texto sin ambigüedad; la
    textura descarta además las zonas planas.
    """
    return [color[i] >= COLOR_MIN and textura[i] >= TEXTURA_MIN
            for i in range(gw * gh)]


def _componentes(marca, gw, gh):
    """Componentes conexas (4-vecindad) de las celdas marcadas."""
    visto = [False] * (gw * gh)
    comps = []
    for inicio in range(gw * gh):
        if not marca[inicio] or visto[inicio]:
            continue
        pila, celdas = [inicio], []
        visto[inicio] = True
        while pila:
            i = pila.pop()
            celdas.append(i)
            x, y = i % gw, i // gw
            for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                nx, ny = x + dx, y + dy
                if 0 <= nx < gw and 0 <= ny < gh:
                    j = ny * gw + nx
                    if marca[j] and not visto[j]:
                        visto[j] = True
                        pila.append(j)
        comps.append(celdas)
    return comps


def _mejor_bloque(comps, gw, gh):
    """
    Elige el bloque macizo más grande: una foto llena su rectángulo; una
    mancha de texto es rala. Se exige tamaño y densidad mínimos.
    """
    mejor, mejor_area = None, 0
    for celdas in comps:
        xs = [i % gw for i in celdas]
        ys = [i // gw for i in celdas]
        x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
        ancho, alto = x1 - x0 + 1, y1 - y0 + 1
        area = ancho * alto
        densidad = len(celdas) / area
        # descarta tiras finas (texto) y bloques poco llenos
        if ancho < gw * 0.18 or alto < gh * 0.10:
            continue
        if densidad < 0.42:
            continue
        if area > mejor_area:
            mejor, mejor_area = (x0, y0, x1, y1), area
    return mejor


def recortar_foto(ruta_entrada, ruta_salida):
    """
    Aísla y recorta la foto de la escena. Devuelve un dict con:
      recortada: True si logró aislar la foto; False si conservó la captura.
      caja:      (izq, arriba, der, abajo) en píxeles del original, o None.
      salida:    ruta escrita.
    """
    ruta_entrada, ruta_salida = Path(ruta_entrada), Path(ruta_salida)
    with Image.open(ruta_entrada) as im:
        im = im.convert("RGB")
        W, H = im.size
        escala = ANCHO_ANALISIS / W if W > ANCHO_ANALISIS else 1.0
        imr = im.resize((max(1, int(W * escala)), max(1, int(H * escala))),
                        Image.BILINEAR) if escala != 1.0 else im.copy()

        gw, gh, textura, color = _rejilla(imr)
        marca = _celdas_foto(gw, gh, textura, color)
        caja = _mejor_bloque(_componentes(marca, gw, gh), gw, gh)

        ruta_salida.parent.mkdir(parents=True, exist_ok=True)

        if caja is None:
            # No se pudo aislar: se conserva la captura completa (reservada).
            im.save(ruta_salida)
            return {"recortada": False, "caja": None, "salida": str(ruta_salida),
                    "motivo": "no se aisló una foto con confianza; captura conservada"}

        # celdas -> píxeles del original, con un margen
        cx, cy = W / gw, H / gh
        x0 = max(0, int((caja[0] - MARGEN * gw) * cx))
        y0 = max(0, int((caja[1] - MARGEN * gh) * cy))
        x1 = min(W, int((caja[2] + 1 + MARGEN * gw) * cx))
        y1 = min(H, int((caja[3] + 1 + MARGEN * gh) * cy))

        cobertura = (x1 - x0) * (y1 - y0) / (W * H)
        # Si el bloque casi cubre todo, la imagen probablemente YA era una foto
        # (no una captura de chat): se guarda tal cual, sin recorte cosmético.
        if cobertura >= 0.92:
            im.save(ruta_salida)
            return {"recortada": False, "caja": (0, 0, W, H),
                    "salida": str(ruta_salida),
                    "motivo": "la imagen ya era una foto; sin chat que recortar"}

        im.crop((x0, y0, x1, y1)).save(ruta_salida)
        return {"recortada": True, "caja": (x0, y0, x1, y1),
                "salida": str(ruta_salida),
                "motivo": f"foto aislada ({int(cobertura*100)}% de la captura)"}


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("uso: python src/recorte.py <captura> [salida]")
        raise SystemExit(1)
    entrada = Path(sys.argv[1])
    salida = Path(sys.argv[2]) if len(sys.argv) > 2 else \
        entrada.with_name(entrada.stem + "_foto.png")
    r = recortar_foto(entrada, salida)
    print(f"{'RECORTADA' if r['recortada'] else 'CONSERVADA'} - {r['motivo']}")
    print(f"  caja: {r['caja']}")
    print(f"  -> {r['salida']}")
