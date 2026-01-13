import pdfplumber
import re
import glob
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import os
import calendar


# =========================================================
# CONFIGURACIÓN GENERAL 
# =========================================================
PERIODO = "12.2025"          # <<< COLOCAR PERIODO A CONTROLAR
MES, ANIO = PERIODO.split(".")
MES_ANIO = f"{MES}{ANIO}"

BASE = fr"Z:\IMPUESTOS\Control_impuestos\{PERIODO}\control_retenciones"
BASE_PADRONES = os.path.join(BASE, "padrones")
BASE_PDFS = os.path.join(BASE, "bejerman")

# =========================================================
# 1. PADRONES
# =========================================================
padron_files = {
    "BA": os.path.join(BASE_PADRONES, f"PadronRGSRet{MES_ANIO}.txt"),
    "CABA": os.path.join(BASE_PADRONES, f"ARDJU008{MES_ANIO}.txt"),
    "ER": os.path.join(BASE_PADRONES, f"PadronRetPer{ANIO}{MES}")
}

padrones = {}

for jurisdiccion, file in padron_files.items():
    padron_data = {}

    if jurisdiccion in ["BA", "CABA"]:
        if not os.path.exists(file):
            raise FileNotFoundError(f"No existe el padrón {jurisdiccion}: {file}")

        with open(file, "r", encoding="latin1") as f:
            for line in f:
                parts = line.strip().split(";")

                if jurisdiccion == "BA" and parts[0] == "R":
                    cuit = parts[4]
                    alicuota = Decimal(parts[8].replace(",", "."))
                    padron_data[cuit] = alicuota

                elif jurisdiccion == "CABA" and len(parts) >= 12:
                    cuit = parts[3]
                    alicuota = Decimal(parts[8].replace(",", "."))
                    padron_data[cuit] = alicuota

    elif jurisdiccion == "ER":
        archivos_posibles = glob.glob(file + ".xls") + glob.glob(file + ".xlsx")
        if not archivos_posibles:
            raise FileNotFoundError(
                f"No se encontró el padrón ER (.xls/.xlsx): {file}"
            )

        padron_path = archivos_posibles[0]
        df = pd.read_excel(padron_path)

        for _, row in df.iterrows():
            cuit = str(row["CUIT"]).strip()
            alicuota = Decimal(str(row["ALICUOTA RETENCION"]).replace(",", "."))
            padron_data[cuit] = alicuota

    padrones[jurisdiccion] = padron_data

# =========================================================
# 2. PDFs BEJERMAN
# =========================================================
pdf_files = {
    "BA": os.path.join(BASE_PDFS, f"Ret {PERIODO}.pdf"),
    "CABA": os.path.join(BASE_PDFS, f"Ret caba {PERIODO}.pdf"),
    "ER": os.path.join(BASE_PDFS, f"Ret DGR Entre Rios {PERIODO}.pdf")
}

# Validar existencia de PDFs
for j, ruta in pdf_files.items():
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No existe el PDF {j}: {ruta}")

# =========================================================
# 3. PROCESAMIENTO PDFs
# =========================================================
cuit_pattern = re.compile(r"(\d{2}-\d{8}-\d)")
line_pattern = re.compile(
    r"(\d{2}/\d{2}/\d{2}).*?"
    r"(-?\d{1,3}(?:\.\d{3})*,\d{2})\s+"
    r"(-?\d{1,3}(?:\.\d{3})*,\d{2})\s+"
    r"(-?\d{1,3}(?:\.\d{3})*,\d{2})$"
)

resultados = []

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
                    neto = Decimal(
                        match_line.group(3).replace(".", "").replace(",", ".")
                    )
                    retenido = Decimal(
                        match_line.group(4).replace(".", "").replace(",", ".")
                    )

                    if current_cuit in padron_data:
                        alicuota = padron_data[current_cuit]
                        esperado = (
                            neto * alicuota / Decimal(100)
                        ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

                        diferencia = (
                            esperado - retenido
                        ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

                        resultado = (
                            "OK"
                            if abs(diferencia) <= Decimal("0.01")
                            else f"DIFERENCIA {diferencia}"
                        )
                    else:
                        alicuota = Decimal("0.00")
                        esperado = Decimal("0.00")
                        resultado = "SIN PADRON"

                    resultados.append({
                        "Jurisdiccion": jurisdiccion,
                        "CUIT": current_cuit,
                        "Fecha": fecha,
                        "Neto": float(neto),
                        "Retenido": float(retenido),
                        "Alicuota": float(alicuota),
                        "Esperado": float(esperado),
                        "Resultado": resultado
                    })

# =========================================================
# 4. EXPORTAR A EXCEL
# =========================================================
mes_num = int(MES)
mes_nombre = calendar.month_name[mes_num].lower()

meses_es = {
    "january": "enero", "february": "febrero", "march": "marzo",
    "april": "abril", "may": "mayo", "june": "junio",
    "july": "julio", "august": "agosto", "september": "septiembre",
    "october": "octubre", "november": "noviembre", "december": "diciembre"
}
mes_nombre = meses_es[mes_nombre]

df_resultados = pd.DataFrame(resultados)

OUTPUT_PATH = os.path.join(
    BASE,
    f"resultado_cruce_{mes_nombre}_{ANIO}.xlsx"
)

with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
    df_resultados.to_excel(
        writer,
        index=False,
        sheet_name=f"{mes_nombre.capitalize()} {ANIO}"
    )

print("✅ Cruce finalizado")
print("📄 Archivo generado en:", OUTPUT_PATH)
