from __future__ import annotations

from pathlib import Path

import pdfplumber

from app.core.models import SourceDocument, SourceTable


def read_pdf(path: Path) -> SourceDocument:
    tables: list[SourceTable] = []
    text_blocks: list[str] = []

    with pdfplumber.open(path) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            text = (page.extract_text() or "").strip()
            if text:
                text_blocks.append(text)

            extracted = page.extract_tables() or []
            for table_index, raw_table in enumerate(extracted, start=1):
                if not raw_table or len(raw_table) < 2:
                    continue
                headers = [
                    str(cell).strip() if cell is not None else f"col_{i+1}"
                    for i, cell in enumerate(raw_table[0])
                ]
                # Ensure unique column names
                seen: dict[str, int] = {}
                columns: list[str] = []
                for header in headers:
                    key = header or "col"
                    count = seen.get(key, 0)
                    columns.append(key if count == 0 else f"{key}_{count+1}")
                    seen[key] = count + 1

                rows: list[dict[str, object]] = []
                for raw_row in raw_table[1:]:
                    values = list(raw_row) + [None] * max(0, len(columns) - len(raw_row))
                    rows.append(
                        {
                            columns[i]: ("" if values[i] is None else str(values[i]).strip())
                            for i in range(len(columns))
                        }
                    )
                tables.append(
                    SourceTable(
                        name=f"page_{page_index}_table_{table_index}",
                        columns=columns,
                        rows=rows,
                    )
                )

    return SourceDocument(
        path=path,
        tables=tables,
        text_blocks=text_blocks,
        metadata={"format": "pdf", "page_count": len(text_blocks) or None},
    )
