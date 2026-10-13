from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any


class ReportKind(str, Enum):
    SUMMARY = "summary"
    DETAILED = "detailed"
    COMPARISON = "comparison"

    @property
    def label_ru(self) -> str:
        return {
            ReportKind.SUMMARY: "Сводный отчёт",
            ReportKind.DETAILED: "Детальный отчёт",
            ReportKind.COMPARISON: "Сравнительный отчёт",
        }[self]


@dataclass
class SourceTable:
    """One tabular dataset extracted from an input file."""

    name: str
    columns: list[str]
    rows: list[dict[str, Any]] = field(default_factory=list)

    @property
    def row_count(self) -> int:
        return len(self.rows)


@dataclass
class SourceDocument:
    """Normalized content from any supported input format."""

    path: Path
    tables: list[SourceTable] = field(default_factory=list)
    text_blocks: list[str] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def source_name(self) -> str:
        return self.path.name
