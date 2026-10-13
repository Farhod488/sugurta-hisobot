from pathlib import Path

from app.tushum import Inspector, TushumInputs, generate_tushum_report
from docx import Document

root = Path(__file__).resolve().parents[1]
src = root / "templates" / "source"

inputs = TushumInputs(
    kuzatuv_pdf=next(src.glob("KUZATUV*.pdf")),
    kuzatuv_xlsx=src / "KUZATUV.xlsx",
    osago_pdf=next(src.glob("OSAGO*.pdf")),
    osago_xlsx=src / "OSAGO.xlsx",
    period_text="24-31-avgust",
    templates_dir=root / "templates",
    output_dir=root / "output",
)

lines: list[str] = []
for inspector in Inspector:
    path = generate_tushum_report(inspector, inputs)
    lines.append(f"=== {inspector.value} -> {path.name} ===")
    doc = Document(path)
    started = False
    idx = 0
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        norm = text.upper().replace("ʼ", "'").replace("'", "")
        if "MALUMOT" in norm:
            started = True
            continue
        if started:
            lines.append(f"[{idx}] {text}")
            idx += 1
    lines.append("")

out = root / "templates" / "_extracted" / "generated_preview.txt"
out.write_text("\n".join(lines), encoding="utf-8")
print("wrote", out)
