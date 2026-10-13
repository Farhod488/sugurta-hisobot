from pathlib import Path

from docx import Document

from app.tushum import Inspector, TushumInputs, generate_tushum_report


def content_texts(doc: Document) -> list[str]:
    started = False
    result: list[str] = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not started:
            normalized = text.upper().replace("ʼ", "'").replace("'", "")
            if "MALUMOT" in normalized:
                started = True
            continue
        if text:
            result.append(text)
    return result


root = Path(__file__).resolve().parents[1]
src = root / "templates" / "source"
base_paras = content_texts(Document(root / "templates" / "tushum_aripov.docx"))

inputs = TushumInputs(
    kuzatuv_pdf=None,
    kuzatuv_xlsx=src / "KUZATUV.xlsx",
    osago_pdf=None,
    osago_xlsx=src / "OSAGO.xlsx",
    period_text="1-7-sentyabr",
    templates_dir=root / "templates",
    output_dir=root / "output",
)
out = generate_tushum_report(Inspector.ARIPOV, inputs)
paras = content_texts(Document(out))

print("OUT", out.name)
print("P3 same as template", paras[3] == base_paras[3])
print("P7 same as template", paras[7] == base_paras[7])
print("P4 period updated", "1-7-sentyabr" in paras[4])
print("P8 period updated", "1-7-sentyabr" in paras[8])
print("P9 sum updated", "1-7-sentyabr" in paras[9])
print("P3:", paras[3][:100])
print("P4:", paras[4][:100])
