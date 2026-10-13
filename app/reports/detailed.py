from __future__ import annotations

from pathlib import Path

from app.core.models import SourceDocument
from app.writers import docx_writer as w


def build_detailed_report(
    document: SourceDocument,
    output_dir: Path,
    *,
    template_path: Path | None = None,
) -> Path:
    title = "Детальный отчёт KPI"
    doc = w.apply_template_if_present(template_path, title, document)
    if doc is None:
        doc = w.new_document(title, document.source_name)

    w.add_section_heading(doc, "Полные данные")
    if not document.tables and not document.text_blocks:
        doc.add_paragraph("В исходном файле не найдено таблиц или текста.")

    for table in document.tables:
        w.add_section_heading(doc, f"Таблица: {table.name}", level=2)
        w.add_key_value(doc, "Строк", table.row_count)
        w.add_key_value(doc, "Колонки", ", ".join(table.columns))
        w.add_table(doc, table)
        doc.add_paragraph()

    if document.text_blocks:
        w.add_section_heading(doc, "Текст документа")
        for index, block in enumerate(document.text_blocks, start=1):
            w.add_section_heading(doc, f"Блок {index}", level=3)
            doc.add_paragraph(block)

    filename = f"detailed_{document.path.stem}.docx"
    return w.save_document(doc, output_dir, filename)
