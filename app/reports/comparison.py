from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Any

from app.core.models import SourceDocument, SourceTable
from app.writers import docx_writer as w


def _numeric_values(table: SourceTable) -> dict[str, list[float]]:
    values: dict[str, list[float]] = defaultdict(list)
    for row in table.rows:
        for column in table.columns:
            raw = row.get(column, "")
            if raw in ("", None):
                continue
            try:
                number = float(str(raw).replace(" ", "").replace(",", "."))
            except ValueError:
                continue
            values[column].append(number)
    return values


def _stats(numbers: list[float]) -> dict[str, Any]:
    return {
        "count": len(numbers),
        "sum": round(sum(numbers), 4),
        "avg": round(sum(numbers) / len(numbers), 4) if numbers else 0,
        "min": round(min(numbers), 4) if numbers else 0,
        "max": round(max(numbers), 4) if numbers else 0,
    }


def build_comparison_report(
    document: SourceDocument,
    output_dir: Path,
    *,
    template_path: Path | None = None,
) -> Path:
    title = "Сравнительный отчёт KPI"
    doc = w.apply_template_if_present(template_path, title, document)
    if doc is None:
        doc = w.new_document(title, document.source_name)

    w.add_section_heading(doc, "Сравнение числовых показателей")
    if not document.tables:
        doc.add_paragraph(
            "Таблицы не найдены — сравнение построено нельзя. "
            "Загрузите Excel/CSV или DOCX/PDF с таблицами."
        )
    else:
        for table in document.tables:
            w.add_section_heading(doc, f"Таблица: {table.name}", level=2)
            numeric = _numeric_values(table)
            if not numeric:
                doc.add_paragraph("Числовые колонки не обнаружены.")
                continue

            stats_table = SourceTable(
                name=f"{table.name}_stats",
                columns=["Показатель", "Кол-во", "Сумма", "Среднее", "Мин", "Макс"],
                rows=[
                    {
                        "Показатель": column,
                        "Кол-во": stats["count"],
                        "Сумма": stats["sum"],
                        "Среднее": stats["avg"],
                        "Мин": stats["min"],
                        "Макс": stats["max"],
                    }
                    for column, numbers in numeric.items()
                    for stats in [_stats(numbers)]
                ],
            )
            w.add_table(doc, stats_table)
            doc.add_paragraph()

        if len(document.tables) >= 2:
            w.add_section_heading(doc, "Сопоставление таблиц")
            for table in document.tables:
                w.add_key_value(doc, table.name, f"{table.row_count} строк")

    filename = f"comparison_{document.path.stem}.docx"
    return w.save_document(doc, output_dir, filename)
