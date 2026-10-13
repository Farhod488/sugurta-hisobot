from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.core.models import SourceDocument, SourceTable


def dataframe_to_table(name: str, frame: pd.DataFrame) -> SourceTable:
    cleaned = frame.fillna("")
    columns = [str(col) for col in cleaned.columns.tolist()]
    rows = [
        {columns[i]: row[i] for i in range(len(columns))}
        for row in cleaned.itertuples(index=False, name=None)
    ]
    return SourceTable(name=name, columns=columns, rows=rows)


def read_excel(path: Path) -> SourceDocument:
    sheets = pd.read_excel(path, sheet_name=None, dtype=object)
    tables = [
        dataframe_to_table(sheet_name, frame)
        for sheet_name, frame in sheets.items()
    ]
    return SourceDocument(
        path=path,
        tables=tables,
        metadata={"format": "excel", "sheet_count": len(tables)},
    )
