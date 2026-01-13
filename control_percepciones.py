import pdfplumber
import re
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import os
import calendar

# ===============================
# CONFIGURACIÓN GENERAL
# ===============================
PERIODO = "12.2025"          # <<< COLOCAR PERIODO A CONTROLAR
MES, ANIO = PERIODO.split(".")

BASE = fr"Z:\IMPUESTOS\Control_impuestos\{PERIODO}\control_percepciones"

# ===============================
# 1. PADRONES
# ===============================
padron_files = {
    "BA": os.path.join(BASE, "padrones", f"PadronRGSPer{MES}{ANIO}.txt"),
    "CABA": os.path.join(BASE, "padrones", f"ARDJU008{MES}{ANIO}.txt")
}

padrones = {}

for jurisdiccion, file in padron_files.items():
    if not os.path.exists(file):
        raise FileNotFoundError(f"No existe padrón {jurisdiccion}: {file}")

    padron_data = {}

    with open(file, "r", encoding="latin1") as f:
        for line in f:
            if jurisdiccion == "BA":
                parts = line.strip().split(";")
                if parts[0] == "P":
                    cuit = parts[4]
                    alicuota = Decimal(parts[8].replace(",", "."))
                    padron_data[cuit] = alicuota

            elif jurisdiccion == "CABA":
                cuit = line[27:38].strip()
                alicuota_str = line[45:49].replace(",", ".").strip()
                alicuota = Decimal(alicuota_str) if alicuota_str else Decimal("0.00")
                padron_data[cuit] = alicuota

    padrones[jurisdiccion] = padron_data

# ===============================
# 2. PDFs
# ===============================
pdf_files = {
    "BA": os.path.join(BASE, "bejerman", f"Perc {PERIODO}.pdf"),
    "CABA": os.path.join(BASE, "bejerman", f"Perc caba {PERIODO}.pdf")
}

for j, ruta in pdf_files.items():
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No existe PDF {j}: {ruta}")

# ===============================
# 3. REGEX
# ===============================
cuit_pattern = re.compile(r"(\d{2}-\d{8}-\d)")
line_pattern = re.compile(
    r"(\d{2}/\d{2}/\d{2}).*?(-?\d{1,3}(?:\.\d{3})*,\d{2})\s+(-?\d{1,3}(?:\.\d{3})*,\d{2})$"
)

resultados_final = []

# ===============================
# 4. PROCESAMIENTO
# ===============================
for jurisdiccion, pdf_file in pdf_files.items():
    padron_data = padrones[jurisdiccion]

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text().splitlines()
            current_cuit = None

            for line in text:
                match_cuit = cuit_pattern.search(line)
                if match_cuit:
                    current_cuit = match_cuit.group(1).replace("-", "")

                match_line = line_pattern.search(line)
                if current_cuit and match_line:
                    fecha = match_line.group(1)
                    neto = Decimal(match_line.group(2).replace(".", "").replace(",", "."))
                    perc = Decimal(match_line.group(3).replace(".", "").replace(",", "."))

                    perc_abs = abs(perc)  # <<< IMPORTANTE PARA NEGATIVOS

                    if current_cuit in padron_data:
                        alicuota = padron_data[current_cuit]
                        esperado = (neto * alicuota / Decimal(100)).quantize(
                            Decimal("0.01"), rounding=ROUND_HALF_UP
                        )
                        diferencia = (esperado - perc_abs).quantize(
                            Decimal("0.01"), rounding=ROUND_HALF_UP
                        )
                        resultado = (
                            "OK"
                            if abs(diferencia) <= Decimal("0.01")
                            else f"DIFERENCIA:{diferencia:.2f}"
                        )
                    else:
                        alicuota = Decimal("0.00")
                        esperado = Decimal("0.00")
                        diferencia = Decimal("0.00")
                        resultado = "SIN PADRON"

                    resultados_final.append({
                        "Jurisdiccion": jurisdiccion,
                        "CUIT": current_cuit,
                        "Fecha": fecha,
                        "Neto": float(neto),
                        "Percibido": float(perc),
                        "Percibido ABS": float(perc_abs),
                        "Alicuota": float(alicuota),
                        "Esperado": float(esperado),
                        "Diferencia": float(diferencia),
                        "Resultado": resultado
                    })

# ===============================
# 5. EXPORTAR EXCEL
# ===============================
mes_num = int(MES)
mes_nombre = calendar.month_name[mes_num].lower()

meses_es = {
    "january": "enero", "february": "febrero", "march": "marzo",
    "april": "abril", "may": "mayo", "june": "junio",
    "july": "julio", "august": "agosto", "september": "septiembre",
    "october": "octubre", "november": "noviembre", "december": "diciembre"
}

mes_nombre = meses_es[mes_nombre]

df_final = pd.DataFrame(resultados_final)

OUTPUT_PATH = os.path.join(
    BASE,
    f"resultado_cruce_{mes_nombre}_{ANIO}.xlsx"
)

with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
    df_final.to_excel(
        writer,
        sheet_name=f"{mes_nombre.capitalize()} {ANIO}",
        index=False
    )

print("✅ Cruce de percepciones finalizado")
print("📁 Archivo generado en:", OUTPUT_PATH)
