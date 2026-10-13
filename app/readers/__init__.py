from __future__ import annotations

from pathlib import Path

from app.core.models import SourceDocument
from app.readers.csv_reader import read_csv
from app.readers.docx_reader import read_docx
from app.readers.excel_reader import read_excel
from app.readers.pdf_reader import read_pdf


def load_document(path: Path | str) -> SourceDocument:
    source = Path(path)
    suffix = source.suffix.lower()

    if suffix in {".xlsx", ".xls"}:
        return read_excel(source)
    if suffix == ".csv":
        return read_csv(source)
    if suffix == ".pdf":
        return read_pdf(source)
    if suffix == ".docx":
        return read_docx(source)

    raise ValueError(f"Неподдерживаемый формат: {suffix}")
