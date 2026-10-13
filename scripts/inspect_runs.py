from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


def dump_runs(path: Path, out: Path) -> None:
    doc = Document(path)
    lines: list[str] = []
    started = False
    idx = 0
    for p in doc.paragraphs:
        text = p.text.strip()
        if not started:
            n = text.upper().replace("ʼ", "'").replace("'", "")
            if "MALUMOT" in n:
                started = True
            continue
        if not text:
            continue
        lines.append(f"=== P{idx} ===")
        lines.append(text[:200])
        for ri, run in enumerate(p.runs):
            rpr = run._element.rPr
            color = None
            highlight = None
            if rpr is not None:
                c = rpr.find(qn("w:color"))
                if c is not None:
                    color = c.get(qn("w:val"))
                h = rpr.find(qn("w:highlight"))
                if h is not None:
                    highlight = h.get(qn("w:val"))
                # theme color?
                if c is not None:
                    color = {
                        "val": c.get(qn("w:val")),
                        "theme": c.get(qn("w:themeColor")),
                    }
            lines.append(
                f"  run[{ri}] bold={run.bold} italic={run.italic} "
                f"color={color} highlight={highlight} "
                f"font={run.font.name} size={run.font.size} "
                f"text={run.text!r}"
            )
        idx += 1
        lines.append("")
    out.write_text("\n".join(lines), encoding="utf-8")
    print("wrote", out)


root = Path(__file__).resolve().parents[1]
dump_runs(
    root / "templates" / "tushum_aripov.docx",
    root / "templates" / "_extracted" / "aripov_runs.txt",
)
dump_runs(
    root / "templates" / "tushum_abdimurodov.docx",
    root / "templates" / "_extracted" / "abdi_runs.txt",
)
