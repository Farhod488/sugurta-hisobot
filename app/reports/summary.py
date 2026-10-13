from __future__ import annotations

from pathlib import Path

from app.core.models import SourceDocument
from app.writers import docx_writer as w


def build_summary_report(
    document: SourceDocument,
    output_dir: Path,
    *,
    template_path: Path | None = None,
) -> Path:
    title = "Сводный отчёт KPI"
    doc = w.apply_template_if_present(template_path, title, document)
    if doc is None:
        doc = w.new_document(title, document.source_name)

    w.add_section_heading(doc, "Обзор данных")
    total_rows = sum(table.row_count for table in document.tables)
    w.add_key_value(doc, "Файл", document.source_name)
    w.add_key_value(doc, "Формат", document.metadata.get("format", "—"))
    w.add_key_value(doc, "Таблиц", len(document.tables))
    w.add_key_value(doc, "Всего строк", total_rows)
    w.add_key_value(doc, "Текстовых блоков", len(document.text_blocks))

    if document.tables:
        w.add_section_heading(doc, "Краткое содержание таблиц")
        for table in document.tables:
            w.add_key_value(doc, table.name, f"{table.row_count} строк, {len(table.columns)} колонок")
            preview = min(5, table.row_count)
            if preview:
                doc.add_paragraph(f"Первые {preview} строк:")
                w.add_table(doc, table, max_rows=preview)
                doc.add_paragraph()

    if document.text_blocks and not document.tables:
        w.add_section_heading(doc, "Извлечённый текст (фрагмент)")
        for block in document.text_blocks[:3]:
            doc.add_paragraph(block[:800] + ("…" if len(block) > 800 else ""))

    filename = f"summary_{document.path.stem}.docx"
    return w.save_document(doc, output_dir, filename)
