from __future__ import annotations

from pathlib import Path

from app.core.models import ReportKind, SourceDocument
from app.reports.comparison import build_comparison_report
from app.reports.detailed import build_detailed_report
from app.reports.summary import build_summary_report


def build_report(
    *,
    document: SourceDocument,
    kind: ReportKind,
    output_dir: Path,
    template_path: Path | None = None,
) -> Path:
    builders = {
        ReportKind.SUMMARY: build_summary_report,
        ReportKind.DETAILED: build_detailed_report,
        ReportKind.COMPARISON: build_comparison_report,
    }
    builder = builders[kind]
    return builder(document, output_dir, template_path=template_path)
