"""
SIGEO-HD DGSPYT — Supresion de datos personales para publicacion.

El tablero se publica en GitHub Pages, es decir, en Internet abierto. Las notas
de cabina del C5 y el desarrollo de hechos de los HD contienen datos personales
y datos personales sensibles: nombres de victimas y de personas desaparecidas,
edad, estatura, tatuajes, vestimenta, domicilio, telefonos, claves de operador
y nombres de servidores publicos.

Difundirlos abiertamente es incompatible con la Ley General de Proteccion de
Datos Personales en Posesion de Sujetos Obligados (arts. 3 fr. X, 6, 16, 21) y
con el interes superior de las victimas y sus familias.

Este modulo enmascara esos elementos ANTES de escribir los JSON publicables.
La base SQLite local conserva el texto integro para el trabajo operativo
interno de la DGSPYT.

El enmascarado es deliberadamente amplio: ante la duda, suprime. Perder una
palabra de contexto operativo cuesta menos que publicar la identidad de una
victima.
"""

import re

MARCA = "[DATO PERSONAL SUPRIMIDO]"
MARCA_VICTIMA = "[VÍCTIMA]"
MARCA_TEL = "[TELÉFONO]"
MARCA_OPERADOR = "[OPERADOR C5]"

_MAY = r"A-ZÁÉÍÓÚÜÑ"
_MIN = r"a-záéíóúüñ"

REGLAS = [
    # Claves de operador y despachador del C5 (C5ECA911IGMORENOG, C5TOLCRO...).
    (re.compile(r"\bC5[A-Z0-9]{6,}\b"), MARCA_OPERADOR),

    # Telefonos: (55) 1234-5678, 5512345678, 55 1234 5678.
    (re.compile(r"\(\d{2,3}\)\s*\d{3,4}[\s\-]?\d{4}"), MARCA_TEL),
    (re.compile(r"\b\d{10}\b"), MARCA_TEL),

    # Extensiones con nombre de servidor publico: "EXT.- 13104 LIC. EMMA".
    (re.compile(r"\bEXT\.?\s*-?\s*\d{3,6}\s*(?:AT|LIC\.?|C\.?)?\s*"
                rf"[{_MAY}][{_MAY}{_MIN}]+"), MARCA),

    # Nombre precedido de tratamiento: "C. ANDREA ABIGAIL GALLARDO IRIARTE",
    # "LIC. EMMA", "SR. JUAN PEREZ".
    (re.compile(rf"\b(?:EL\s+|LA\s+)?(?:C|LIC|SR|SRA|SRITA|MTRO|MTRA|DR|DRA|"
                rf"OFICIAL|CMTE|COMANDANTE)\.?\s+"
                rf"[{_MAY}][{_MAY}{_MIN}]+(?:\s+[{_MAY}][{_MAY}{_MIN}]+){{1,4}}\b"),
     MARCA),

    # Nombre entre parentesis tras el marcador de persona desaparecida.
    (re.compile(rf"\(\s*[{_MAY}][{_MAY}{_MIN}]+"
                rf"(?:\s+[{_MAY}][{_MAY}{_MIN}]+){{2,4}}\s*\)"), f"({MARCA})"),

    # Parentesco + nombre propio: "SU HIJA ANDREA ABIGAIL GALLARDO IRIARTE".
    (re.compile(rf"\b(SU\s+(?:HIJ[AO]|ESPOS[AO]|HERMAN[AO]|MADRE|PADRE|"
                rf"NIET[AO]|SOBRIN[AO]|PAREJA|EXPAREJA|CUÑAD[AO]|TI[AO]))\s+"
                rf"[{_MAY}][{_MAY}{_MIN}]+(?:\s+[{_MAY}][{_MAY}{_MIN}]+){{1,4}}\b"),
     r"\1 " + MARCA),

    # Senas particulares que reidentifican a una persona desaparecida.
    (re.compile(r"\bTATUAJE[S]?\b[^\.]{0,120}"), "[SEÑAS PARTICULARES SUPRIMIDAS]"),
    (re.compile(r"\bESTATURA\s+\d[\.,]?\d*\s*(?:M|MTS|METROS|CM)?"), MARCA),

    # CURP y RFC.
    (re.compile(r"\b[A-Z]{4}\d{6}[HM][A-Z]{5}[A-Z0-9]{2}\b"), MARCA),
    (re.compile(r"\b[A-Z]{3,4}\d{6}[A-Z0-9]{3}\b"), MARCA),

    # Placas vehiculares del Estado de Mexico.
    (re.compile(r"\bPLACAS?\s*:?\s*[A-Z0-9\-]{5,10}\b"), "PLACAS " + MARCA),
]

# Nombre de victima al inicio del desarrollo de hechos, con edad entre
# parentesis: "Jonathan Yair Estrada Nava (26) fue localizado sin vida".
_RE_VICTIMA_EDAD = re.compile(
    rf"\b[{_MAY}][{_MIN}]+(?:\s+[{_MAY}][{_MIN}]+){{1,4}}\s*\((\d{{1,3}})\)")

# Variante en mayusculas dentro de notas de cabina.
_RE_VICTIMA_EDAD_MAY = re.compile(
    rf"\b[{_MAY}][{_MAY}{_MIN}]+(?:\s+[{_MAY}][{_MAY}{_MIN}]+){{2,4}}\s+"
    rf"DE\s+(\d{{1,3}})\s*(?:AÑOS|ANOS)")


def anonimizar_texto(texto):
    """Enmascara datos personales en texto libre. Conserva el hecho, no la identidad."""
    if not texto:
        return texto
    s = str(texto)

    # La edad es dato estadistico util; el nombre no. Se conserva la edad.
    s = _RE_VICTIMA_EDAD.sub(lambda m: f"{MARCA_VICTIMA} ({m.group(1)})", s)
    s = _RE_VICTIMA_EDAD_MAY.sub(
        lambda m: f"{MARCA_VICTIMA} DE {m.group(1)} AÑOS", s)

    for patron, reemplazo in REGLAS:
        s = patron.sub(reemplazo, s)

    return re.sub(r"\s{2,}", " ", s).strip()


CAMPOS_LIBRES = (
    "notas", "referencia", "desarrollo_hechos", "observaciones",
    "informacion_adicional", "acciones_ssem",
)


def anonimizar_registro(registro):
    """Devuelve copia del registro con los campos de texto libre enmascarados."""
    salida = dict(registro)
    for campo in CAMPOS_LIBRES:
        if campo in salida and isinstance(salida[campo], str):
            salida[campo] = anonimizar_texto(salida[campo])
    return salida


def anonimizar_lista(registros):
    return [anonimizar_registro(r) for r in registros]


if __name__ == "__main__":  # verificacion rapida
    muestras = [
        "Jonathan Yair Estrada Nava (26) fue localizado sin vida con impactos "
        "de arma de fuego, al interior de un domicilio",
        "REPORTA QUE SU HIJA C. ANDREA ABIGAIL GALLARDO IRIARTE DE 19, AÑOS "
        "SALIÓ DEL DOMICILIO, ESTATURA 1.52, TATUAJES EN PIERNAS, ANDRE Y "
        "ALEJANDRO. EXT.- 13104 LIC. EMMA. El usuario C5ECA911IGMORENOG ha "
        "cambiado el estatus de la persona desaparecida #1 (ANDREA ABIGAIL "
        "GALLARDO IRIARTE) a Persona extraviada. TEL (722) 535-3842",
    ]
    for m in muestras:
        print(anonimizar_texto(m))
        print("-" * 70)
