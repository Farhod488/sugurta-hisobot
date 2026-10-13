from __future__ import annotations

from pathlib import Path

from app.core.models import ReportKind, SourceDocument
from app.readers import load_document
from app.reports import build_report


SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv", ".pdf", ".docx"}


def generate_report(
    input_path: Path | str,
    report_kind: ReportKind | str,
    output_dir: Path | str,
    *,
    template_path: Path | str | None = None,
) -> Path:
    """Read input file, build the selected report, write a DOCX file."""
    source = Path(input_path)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not source.exists():
        raise FileNotFoundError(f"Файл не найден: {source}")

    if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Неподдерживаемый формат: {source.suffix}. "
            f"Допустимо: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
        )

    kind = ReportKind(report_kind) if isinstance(report_kind, str) else report_kind
    document: SourceDocument = load_document(source)
    template = Path(template_path) if template_path else None

    return build_report(
        document=document,
        kind=kind,
        output_dir=out_dir,
        template_path=template,
    )
