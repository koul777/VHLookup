from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd


MONTH_HEADER_RE = re.compile(r"^(20\d{2}[-./년])?\s?(0?[1-9]|1[0-2])\s?(월|month)?$", re.IGNORECASE)
QUARTER_HEADER_RE = re.compile(r"^(20\d{2}[-./년])?\s?[1-4]\s?(분기|q)$", re.IGNORECASE)


@dataclass(frozen=True)
class HorizontalDetection:
    is_horizontal: bool
    value_columns: tuple[str, ...]
    confidence: float


class HorizontalTableEngine:
    def detect(self, frame: pd.DataFrame) -> HorizontalDetection:
        headers = [str(column).strip() for column in frame.columns]
        value_columns = tuple(column for column in headers if self._looks_horizontal_value_header(column))
        confidence = len(value_columns) / max(len(headers), 1)
        return HorizontalDetection(
            is_horizontal=len(value_columns) >= 2 and confidence >= 0.25,
            value_columns=value_columns,
            confidence=round(min(confidence * 1.5, 1.0), 4),
        )

    def wide_to_long(
        self,
        frame: pd.DataFrame,
        id_columns: list[str] | tuple[str, ...],
        value_columns: list[str] | tuple[str, ...] | None = None,
        variable_name: str = "열 기준",
        value_name: str = "값",
    ) -> pd.DataFrame:
        missing_ids = [column for column in id_columns if column not in frame.columns]
        if missing_ids:
            raise KeyError(f"Missing id columns: {missing_ids}")
        if value_columns is None:
            detection = self.detect(frame)
            value_columns = detection.value_columns
        missing_values = [column for column in value_columns if column not in frame.columns]
        if missing_values:
            raise KeyError(f"Missing value columns: {missing_values}")
        return frame.melt(
            id_vars=list(id_columns),
            value_vars=list(value_columns),
            var_name=variable_name,
            value_name=value_name,
        )

    def _looks_horizontal_value_header(self, value: str) -> bool:
        text = value.strip()
        return bool(MONTH_HEADER_RE.match(text) or QUARTER_HEADER_RE.match(text))
