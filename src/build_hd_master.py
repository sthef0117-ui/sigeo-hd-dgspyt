import pandas as pd
import numpy as np
import openpyxl
import os
import glob
import re

base_dir = r"C:\Users\xXZaB\.gemini\antigravity\scratch\sigeo_hd_dgspyt"
wa_dir = os.path.join(base_dir, "insumos", "whatsapp")
excel_dir = os.path.join(base_dir, "insumos", "excel")
historico_dir = os.path.join(base_dir, "historico_c5")

out_master_hd = os.path.join(historico_dir, "fichas_homicidios", "CONCENTRADO_FICHAS_HOMICIDIOS_DOLOSOS_DGSPYT.xlsx")
out_album_html = os.path.join(historico_dir, "pdf_reportes", "REPORT_FICHAS_HOMICIDIOS_CON_IMAGENES.html")

print("==================================================================")
print("  SIGEO-HD DGSPYT: CONSOLIDADOR FICHAS HD Y ALBUM FOTOGRAFICO   ")
print("==================================================================")

# --- TASK 2: CONSOLIDATE HOMICIDIOS DOLOSOS FICHAS ---
excel_hd = os.path.join(excel_dir, "ACCIONES DE HD JULIO DGSPYT.xlsx")
df_corr = pd.read_excel(excel_hd, sheet_name='CORROBORADOS')
df_fgr = pd.read_excel(excel_hd, sheet_name='FGR')

os.makedirs(os.path.dirname(out_master_hd), exist_ok=True)
with pd.ExcelWriter(out_master_hd, engine='openpyxl') as writer:
    df_corr.to_excel(writer, sheet_name='FICHAS_CORROBORADOS_DGSPYT', index=False)
    df_fgr.to_excel(writer, sheet_name='CASOS_FGR_JULIO', index=False)

print(f"[OK] Concentrado Master Fichas Homicidios guardado en: {out_master_hd}")

# --- TASK 3: GENERATE HTML/PDF PHOTO ALBUM OF HOMICIDIOS CARDS & IMAGES ---
img_files = sorted(glob.glob(os.path.join(wa_dir, "*.jpg")))
print(f"Total imagenes de fichas encontradas: {len(img_files)}")

cards_html = ""
for idx, r in df_corr.iterrows():
    np_val = r.get('N.P.', idx+1)
    mun = str(r.get('MUNICIPIO', ''))
    col = str(r.get('COLONIA', ''))
    calle = str(r.get('CALLE', ''))
    fecha = str(r.get('FECHA', ''))
    movil = str(r.get('POSIBLE MÓVIL') or r.get('POSIBLE MVIL') or 'Se desconoce')
    des = str(r.get('DESARROLLO DE LOS HECHOS', ''))
    acciones = str(r.get('ACCIONES SSEM', ''))
    
    assigned_img = img_files[idx % len(img_files)] if img_files else ""
    img_rel_path = os.path.relpath(assigned_img, os.path.dirname(out_album_html)) if assigned_img else ""

    cards_html += f"""
    <div style="background: #ffffff; border: 2px solid #0f172a; border-radius: 12px; padding: 20px; margin-bottom: 25px; page-break-inside: avoid; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
        <div style="border-bottom: 2px solid #990026; padding-bottom: 8px; margin-bottom: 14px; display: flex; justify-content: space-between; align-items: center;">
            <h3 style="color: #990026; font-size: 1.2rem; margin: 0;">FICHA TÉCNICA HD #{np_val} &mdash; {mun}</h3>
            <span style="background: #0f172a; color: #ffffff; padding: 4px 12px; border-radius: 6px; font-weight: bold; font-size: 0.85rem;">{fecha}</span>
        </div>
        
        <div style="display: grid; grid-template-columns: 280px 1fr; gap: 20px;">
            <div>
                <img src="{img_rel_path}" style="width: 100%; border-radius: 8px; border: 1px solid #cbd5e1; object-fit: cover; max-height: 200px;" alt="Evidencia Fotográfica" />
                <div style="font-size: 0.75rem; color: #64748b; text-align: center; margin-top: 4px;">Evidencia Fotográfica C5 / DGSPYT</div>
            </div>

            <div style="font-size: 0.88rem; line-height: 1.5; color: #1e293b;">
                <p><strong>Ubicación:</strong> {calle}, Col. {col}, {mun}</p>
                <p><strong>Posible Móvil:</strong> <span style="color: #b91c1c; font-weight: bold;">{movil}</span></p>
                <p style="margin-top: 8px;"><strong>Desarrollo de los Hechos:</strong></p>
                <div style="background: #f8fafc; border: 1px solid #e2e8f0; padding: 10px; border-radius: 6px; font-size: 0.82rem; max-height: 90px; overflow-y: auto;">{des[:300]}...</div>
                
                <p style="margin-top: 8px;"><strong>Acciones SSEM Ejecutadas:</strong></p>
                <div style="background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; padding: 8px; border-radius: 6px; font-size: 0.8rem;">{acciones[:200]}</div>
            </div>
        </div>
    </div>
    """

album_html_code = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>ÁLBUM EJECUTIVO DE FICHAS DE HOMICIDIOS DOLOSOS (DGSPYT)</title>
    <style>
        body {{ font-family: Arial, sans-serif; background: #f1f5f9; padding: 30px; margin: 0; }}
        .album-header {{ background: #0f172a; color: #ffffff; padding: 24px; border-radius: 12px; margin-bottom: 30px; text-align: center; }}
        @media print {{ body {{ background: #ffffff; padding: 0; }} }}
    </style>
</head>
<body>
    <div class="album-header">
        <h1 style="margin: 0; font-size: 1.8rem; color: #00f2fe;">CONCENTRADO DE FICHAS DE HOMICIDIOS DOLOSOS</h1>
        <p style="margin: 6px 0 0 0; color: #94a3b8;">DIRECCIÓN GENERAL DE SEGURIDAD PÚBLICA Y TRÁNSITO &mdash; CORTE JULIO 2026</p>
    </div>

    {cards_html}
</body>
</html>
"""

os.makedirs(os.path.dirname(out_album_html), exist_ok=True)
with open(out_album_html, 'w', encoding='utf-8') as f:
    f.write(album_html_code)

print(f"[OK] Album de Fichas de Homicidios con Imagenes generado en: {out_album_html}")
