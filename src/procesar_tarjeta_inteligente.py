import os
import glob
import json
import re
import pandas as pd
from datetime import datetime

base_dir = r"C:\Users\xXZaB\.gemini\antigravity\scratch\sigeo_hd_dgspyt"
historico_dir = os.path.join(base_dir, "historico_c5")
ti_dir = os.path.join(historico_dir, "tarjetas_informativas")

master_ti_excel = os.path.join(ti_dir, "CONCENTRADO_MASTER_TARJETAS_INFORMATIVAS.xlsx")
master_911_excel = os.path.join(historico_dir, "concentrados", "CONCENTRADO_HISTORICO_LLAMADAS_911_089_C5.xlsx")
out_super_fichas_html = os.path.join(historico_dir, "pdf_reportes", "SUPER_FICHAS_360_VINCULADAS.html")

os.makedirs(os.path.join(ti_dir, "texto"), exist_ok=True)
os.makedirs(os.path.join(ti_dir, "imagenes"), exist_ok=True)

def parse_whatsapp_card(raw_text, image_path=None):
    """ Extrae metadatos detallados de la narrativa de una Tarjeta Informativa de WhatsApp """
    lines = [line.strip() for line in raw_text.strip().split('\n') if line.strip()]

    # Municipio
    mun_match = re.search(r'([A-ZÁÉÍÓÚÑ\s]+)\s*\((?:PERSONA|HOMICIDIO|CADAVER|SUICIDIO)', raw_text, re.IGNORECASE)
    municipio = mun_match.group(1).strip().upper() if mun_match else "DESCONOCIDO"

    # Victim Name
    victim_match = re.search(r'nombre de\s+([A-ZÁÉÍÓÚÑ\s]{5,40})(?:\s+de\s+\d+|\s*,|\s*localiz)', raw_text, re.IGNORECASE)
    victima = victim_match.group(1).strip().upper() if victim_match else "DESCONOCIDO"

    # Age
    age_match = re.search(r'(\d{1,2})\s*años', raw_text, re.IGNORECASE)
    edad = int(age_match.group(1)) if age_match else None

    # Shells / Caliber
    shells_match = re.search(r'(\d+)\s*casquillos\s*percutidos\s*(?:calibre|de)?\s*([0-9a-zA-Z\s]{2,10})', raw_text, re.IGNORECASE)
    casquillos = f"{shells_match.group(1)} casquillos {shells_match.group(2).strip()}" if shells_match else "No especificado"

    # Police Officers / Units
    police_match = re.search(r'Policía\s+([A-ZÁÉÍÓÚÑ\s]{5,35})\s+unidad\s+([A-Z0-9\-]+)', raw_text, re.IGNORECASE)
    policia_unidad = f"{police_match.group(1).strip()} ({police_match.group(2).strip()})" if police_match else "No especificado"

    # Location (Street / Colonia)
    loc_match = re.search(r'en\s+(calle\s+[^,]+,\s*colonia\s+[^,]+)', raw_text, re.IGNORECASE)
    ubicacion = loc_match.group(1).strip() if loc_match else "No especificada"

    # Time
    time_match = re.search(r'A las\s+(\d{1,2}:\d{2})\s*horas', raw_text, re.IGNORECASE)
    hora = time_match.group(1) if time_match else "No especificada"

    return {
        "municipio": municipio,
        "victima": victima,
        "edad": edad,
        "casquillos": casquillos,
        "policia_unidad": policia_unidad,
        "ubicacion": ubicacion,
        "hora": hora,
        "narrativa_completa": raw_text.strip(),
        "imagen_path": image_path or "",
        "fecha_registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

def cross_reference_with_911(card_data):
    """ Cruza la tarjeta informativa con el Concentrado Máster de llamadas 911 C5 """
    matched_call = None
    if os.path.exists(master_911_excel):
        try:
            df_911 = pd.read_excel(master_911_excel, sheet_name='CONCENTRADO_911')
            
            mun = card_data['municipio']
            filtered = df_911[df_911['MUNICIPIO'].astype(str).str.upper().str.contains(mun, na=False)]
            
            for idx, row in filtered.iterrows():
                notes = str(row.get('NOTAS', '')).upper()
                addr = str(row.get('DIRECCION ', '')).upper()
                ref = str(row.get('REFERENCIA DE UBICACION', '')).upper()
                full_text = f"{notes} {addr} {ref}"

                # Match if victim name or street keywords appear
                if (card_data['victima'] != "DESCONOCIDO" and card_data['victima'] in full_text) or \
                   (card_data['ubicacion'] != "No especificada" and any(w in full_text for w in card_data['ubicacion'].upper().split() if len(w) > 4)):
                    matched_call = {
                        "folio_c5": str(row.get('FOLIO', '')),
                        "incidente": str(row.get('INCIDENTE', '')),
                        "hora_911": str(row.get('HORA', '')),
                        "direccion_911": str(row.get('DIRECCION ', '')),
                        "notas_911": notes[:250]
                    }
                    break
        except Exception as e:
            print(f"Error cruzando con 911: {e}")

    return matched_call

def ingestar_tarjeta_inteligente(raw_text, image_path=None):
    data = parse_whatsapp_card(raw_text, image_path)
    
    # 1. Deduplication check in Master TI Excel
    if os.path.exists(master_ti_excel):
        df_master = pd.read_excel(master_ti_excel)
        if 'victima' in df_master.columns and data['victima'] != "DESCONOCIDO":
            dups = df_master[df_master['victima'].astype(str).str.upper() == data['victima']]
            if not dups.empty:
                print(f"[DUPLICADO DETECTADO] La tarjeta de '{data['victima']}' ya existe en el concentrado. Se omite duplicado.")
                return None

    # 2. Perform Cross-Referencing with 911 Calls
    match_911 = cross_reference_with_911(data)
    if match_911:
        print(f"[CRUCE EXITOSO] Vinculada llamada 911 C5 Folio #{match_911['folio_c5']} con Tarjeta de {data['victima']}")
        data['folio_911_vinculado'] = match_911['folio_c5']
        data['incidente_911'] = match_911['incidente']
        data['coincidencia_estatus'] = "VINCULADA 911 C5"
    else:
        data['folio_911_vinculado'] = "N/A"
        data['incidente_911'] = "N/A"
        data['coincidencia_estatus'] = "SIN REGISTRO PREVIO 911"

    # Save to Master Excel
    df_new = pd.DataFrame([data])
    df_combined = pd.concat([pd.read_excel(master_ti_excel), df_new], ignore_index=True) if os.path.exists(master_ti_excel) else df_new

    with pd.ExcelWriter(master_ti_excel, engine='openpyxl') as writer:
        df_combined.to_excel(writer, sheet_name='TARJETAS_INFORMATIVAS', index=False)

    print(f"[OK] Tarjeta Informativa de '{data['victima']}' ingestada y concatenada exitosamente.")
    return data

# Ingest test cards from WhatsApp screenshots provided by the user
card_ecatepec = """ECATEPEC (PERSONA MUERTA POR IMPACTOS DE BALA)
A las 10:55 horas, el Policía GUTIERREZ RAMIREZ JUAN CARLOS unidad ME334A, informó que, en calle Colorines esquina calle Encino, colonia Tierra Blanca, hizo contacto con el Primer Respondiente Policía Municipal FERMÍN GUZMÁN FERNANDO unidad sector 3-21, indicando de una persona muerta por impactos de bala, quien respondía al nombre de RICARDO HERNÁNDEZ MAYA REYNOSA de 33 años, localizándose en el lugar 5 casquillos percutidos de calibre 9mm; arribo personal de Protección Civil unidad PC-2021, al mando del Paramédico JUAN CARLOS GALVÁN, quien certifica que ya no cuentan con signos vitales."""

card_tecamac = """TECÁMAC (PERSONA MUERTA POR IMPACTO DE BALA)
A las 23:55 horas, el Policía MALDONADO SUÁREZ JUAN TRINIDAD, unidad ME937A5, informó que a las 06:15 horas, en calle Jardines de La Viña esquina 2da. Cerrada de Jardines de La Viña, fraccionamiento Héroes Tecámac, hizo contacto con la Primer Respondiente Policía Municipal NORMA HERNÁNDEZ CASTRO unidad C-072, quien toma conocimiento de una persona muerta quien en vida respondía al nombre de LUIS EDUARDO RUIZ ROJAS de 35 años, el cual se encuentra a bordo de un vehículo marca Chevrolet, tipo Chevy Monza, color blanco, placas PUB383E, con bandera de taxi, localizando 7 casquillos percutidos calibre 9mm, arribo la Unidad Médica PC-70, al mando de la Paramédico ROCIÓ ZARATE RANGEL, indicando que ya no cuenta con signos vitales, debido a los impactos que presenta en la cara, indica la Primer Respondiente que se trató de un intento de asalto y al oponer resistencia le disparan."""

print("=========================================================================")
print("  SIGEO-HD DGSPYT: PROCESADOR INTELIGENTE Y CRUCE 360 DE TARJETAS       ")
print("=========================================================================")

res1 = ingestar_tarjeta_inteligente(card_ecatepec, image_path="whatsapp_ecatepec_colorines.png")
res2 = ingestar_tarjeta_inteligente(card_tecamac, image_path="whatsapp_tecamac_chevy.png")
