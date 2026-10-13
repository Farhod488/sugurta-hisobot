from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from app.tushum import Inspector, TushumInputs, generate_tushum_report


def content_paras(doc: Document):
    started = False
    for p in doc.paragraphs:
        t = p.text.strip()
        if not started:
            if "MALUMOT" in t.upper().replace("ʼ", "'").replace("'", ""):
                started = True
            continue
        if t:
            yield p


root = Path(__file__).resolve().parents[1]
src = root / "templates" / "source"
out = generate_tushum_report(
    Inspector.ARIPOV,
    TushumInputs(
        kuzatuv_pdf=next(src.glob("KUZATUV*.pdf")),
        kuzatuv_xlsx=src / "KUZATUV.xlsx",
        osago_pdf=next(src.glob("OSAGO*.pdf")),
        osago_xlsx=src / "OSAGO.xlsx",
        period_text="24-31-avgust",
        templates_dir=root / "templates",
        output_dir=root / "output",
    ),
)

paras = list(content_paras(Document(out)))
# Check paragraph 3 (totals) runs
p = paras[3]
print("FILE", out.name)
print("TEXT", p.text[:120])
for i, run in enumerate(p.runs[:20]):
    color = None
    rpr = run._element.rPr
    if rpr is not None:
        c = rpr.find(qn("w:color"))
        if c is not None:
            color = c.get(qn("w:val"))
    print(f"{i}: bold={run.bold} color={color} {run.text!r}")
