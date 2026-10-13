from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.core.models import SourceDocument
from app.readers.excel_reader import dataframe_to_table


def read_csv(path: Path) -> SourceDocument:
    frame = None
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "utf-8", "cp1251", "latin-1"):
        try:
            frame = pd.read_csv(path, dtype=object, encoding=encoding)
            break
        except Exception as exc:  # noqa: BLE001 — try next encoding
            last_error = exc
    if frame is None:
        raise ValueError(f"Не удалось прочитать CSV: {last_error}")

    table = dataframe_to_table(path.stem, frame)
    return SourceDocument(
        path=path,
        tables=[table],
        metadata={"format": "csv"},
    )
