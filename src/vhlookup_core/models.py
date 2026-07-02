from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

import pandas as pd


IssueSeverity = Literal["info", "warning", "error"]


@dataclass(frozen=True)
class SheetData:
    name: str
    frame: pd.DataFrame


@dataclass(frozen=True)
class WorkbookSource:
    path: Path
    sheets: tuple[SheetData, ...]


@dataclass(frozen=True)
class SheetProfile:
    sheet_name: str
    row_count: int
    column_count: int
    non_empty_cells: int
    header_confidence: float


@dataclass(frozen=True)
class HeaderDetectionResult:
    sheet_name: str
    header_row_index: int
    header_row_count: int
    data_start_row_index: int
    headers: tuple[str, ...]
    confidence: float
    warnings: tuple[str, ...] = ()

    @property
    def header_row_number(self) -> int:
        return self.header_row_index + 1

    @property
    def data_start_row_number(self) -> int:
        return self.data_start_row_index + 1


@dataclass(frozen=True)
class ColumnMapping:
    source_to_target: dict[str, str]
    target_to_source: dict[str, str]
    confidence_by_target: dict[str, float]
    unmapped_targets: tuple[str, ...] = ()
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True)
class KeySpec:
    reference_key_columns: tuple[str, ...]
    target_key_columns: tuple[str, ...] | None = None
    normalization: Literal["text", "loose_numeric"] = "text"

    def target_columns(self) -> tuple[str, ...]:
        return self.target_key_columns or self.reference_key_columns


@dataclass(frozen=True)
class KeyCandidate:
    reference_columns: tuple[str, ...]
    target_columns: tuple[str, ...]
    score: float
    duplicate_ratio_reference: float
    duplicate_ratio_target: float
    loose_overlap_score: float = 0.0


@dataclass(frozen=True)
class ValidationIssue:
    issue_type: str
    message: str
    severity: IssueSeverity = "warning"
    file_name: str | None = None
    sheet_name: str | None = None
    row_number: int | None = None
    column_name: str | None = None
    details: dict[str, Any] = field(default_factory=dict)


@dataclass
class JobResult:
    result_frame: pd.DataFrame
    issues: list[ValidationIssue] = field(default_factory=list)
    summary: dict[str, Any] = field(default_factory=dict)
    mapping: ColumnMapping | None = None
    mapping_records: list[dict[str, Any]] = field(default_factory=list)
    privacy_records: list[dict[str, Any]] = field(default_factory=list)

    def issue_frame(self) -> pd.DataFrame:
        rows = []
        for issue in self.issues:
            rows.append(
                {
                    "severity": issue.severity,
                    "issue_type": issue.issue_type,
                    "message": issue.message,
                    "file_name": issue.file_name,
                    "sheet_name": issue.sheet_name,
                    "row_number": issue.row_number,
                    "column_name": issue.column_name,
                    "details": issue.details,
                }
            )
        return pd.DataFrame(rows)
