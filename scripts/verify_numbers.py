from pathlib import Path

import pandas as pd
import pdfplumber

src = Path("templates/source")


def read_pdf_jami_kuzatuv(path: Path):
    with pdfplumber.open(path) as pdf:
        table = pdf.pages[0].extract_tables()[0]
    for row in table:
        if row and row[0] and str(row[0]).strip().startswith("Jami"):
            return row
    return None


def read_pdf_jami_osago(path: Path):
    with pdfplumber.open(path) as pdf:
        table = pdf.pages[0].extract_tables()[0]
    for row in table:
        if row and str(row[1] or "").strip().startswith("Jami"):
            return row
    return None


kuz_pdf = next(src.glob("KUZATUV*.pdf"))
osa_pdf = next(src.glob("OSAGO*.pdf"))
print("KUZATUV jami", read_pdf_jami_kuzatuv(kuz_pdf))
print("OSAGO jami", read_pdf_jami_osago(osa_pdf))

kuz = pd.read_excel(src / "KUZATUV.xlsx", header=0)
osa = pd.read_excel(src / "OSAGO.xlsx", header=0)
print("KUZATUV cols", list(kuz.columns))
print(kuz.to_string())
print("OSAGO cols", list(osa.columns))
print(osa.to_string())

aripov_kuz = ["TEMIR YOL", "OZBEKINVEST", "XALQ", "INFINITY"]
aripov_osa = ["ALFA", "INSON"]


def match(df, keys, name_col):
    mask = df[name_col].astype(str).str.upper().apply(
        lambda v: any(k.upper() in v.replace("ʻ", "'").replace("'", "") for k in keys)
    )
    return df[mask]


print("Ariпов KUZ matched")
print(match(kuz, aripov_kuz, kuz.columns[0]).to_string())
print("Ariпов OSA matched")
print(match(osa, aripov_osa, osa.columns[0]).to_string())
