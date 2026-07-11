from __future__ import annotations

import re
from collections.abc import Iterable
from pathlib import Path
from typing import Any

import pandas as pd

from vhlookup_core.models import JobResult
from vhlookup_core.normalization import normalize_header, normalize_key_value


ORDER_VALUE_SEPARATOR_RE = re.compile(r"[\n\r,;\t]+")

ORDER_COLUMN_HINTS = (
    "부서",
    "기관",
    "소속",
    "직제",
    "조직",
    "본부",
    "센터",
    "국",
    "과",
    "팀",
    "직급",
    "직위",
    "직책",
    "department",
    "dept",
    "division",
    "office",
    "team",
    "position",
    "rank",
)


def parse_order_text(text: str) -> list[str]:
    """Parse pasted organization-order text from lines, commas, tabs, or semicolons."""
    values: list[str] = []
    seen: set[str] = set()
    for raw_value in ORDER_VALUE_SEPARATOR_RE.split(text):
        value = normalize_key_value(raw_value)
        if not value or value in seen:
            continue
        values.append(value)
        seen.add(value)
    return values


class OrganizationOrderSorter:
    """Sort rows by a user-defined value order for one selected column."""

    def infer_order_column(self, frame: pd.DataFrame) -> str:
        if frame.empty or not len(frame.columns):
            raise ValueError("정렬 기준으로 사용할 컬럼을 찾을 수 없습니다.")

        candidates: list[tuple[int, int, int, str]] = []
        row_count = max(len(frame), 1)
        for column in frame.columns:
            values = self.unique_values(frame, str(column), limit=min(row_count, 500))
            if not values:
                continue
            unique_count = len(values)
            if unique_count > max(100, row_count):
                continue
            header = normalize_header(column)
            hint_score = 1 if any(hint in header for hint in ORDER_COLUMN_HINTS) else 0
            candidates.append((hint_score, -unique_count, -len(str(column)), str(column)))

        if candidates:
            candidates.sort(reverse=True)
            return candidates[0][3]
        return str(frame.columns[0])

    def unique_values(self, frame: pd.DataFrame, column: str, limit: int | None = None) -> list[str]:
        if column not in frame.columns:
            raise KeyError(f"정렬 기준 컬럼을 찾을 수 없습니다: {column}")

        values: list[str] = []
        seen: set[str] = set()
        for value in frame[column].tolist():
            normalized = normalize_key_value(value)
            if not normalized or normalized in seen:
                continue
            values.append(normalized)
            seen.add(normalized)
            if limit is not None and len(values) >= limit:
                break
        return values

    def sort_frame(
        self,
        frame: pd.DataFrame,
        column: str,
        order_values: Iterable[object],
        metadata: dict[str, Any] | None = None,
    ) -> JobResult:
        if column not in frame.columns:
            raise KeyError(f"정렬 기준 컬럼을 찾을 수 없습니다: {column}")

        order_map = self._order_map(order_values)
        if not order_map:
            raise ValueError("정렬할 값 순서를 1개 이상 지정하세요.")

        working = frame.reset_index(drop=True).copy()
        normalized_values = working[column].map(normalize_key_value)
        default_rank = len(order_map)
        working["__vhlookup_order_rank__"] = normalized_values.map(lambda value: order_map.get(value, default_rank))
        working["__vhlookup_original_order__"] = range(len(working))
        sorted_frame = (
            working.sort_values(
                by=["__vhlookup_order_rank__", "__vhlookup_original_order__"],
                kind="stable",
                ignore_index=True,
            )
            .drop(columns=["__vhlookup_order_rank__", "__vhlookup_original_order__"])
            .reset_index(drop=True)
        )

        unmatched_values = sorted(
            {
                value
                for value in normalized_values.tolist()
                if value and value not in order_map
            }
        )
        missing_order_values = [value for value in order_map if value not in set(normalized_values.tolist())]

        source_name = str((metadata or {}).get("file_name") or "")
        sheet_name = str((metadata or {}).get("sheet_name") or "")
        summary = {
            "workflow": "사용자 지정 순서 정렬",
            "row_count": len(sorted_frame),
            "sort_column": column,
            "specified_order_count": len(order_map),
            "unmatched_value_count": len(unmatched_values),
            "missing_order_value_count": len(missing_order_values),
        }
        if source_name:
            summary["source_file"] = source_name
        if sheet_name:
            summary["source_sheet"] = sheet_name
        if unmatched_values:
            summary["unmatched_values"] = ", ".join(unmatched_values[:30])
        if missing_order_values:
            summary["missing_order_values"] = ", ".join(missing_order_values[:30])

        mapping_records = [
            {
                "role": "organization_order_sort",
                "file_name": source_name,
                "sheet_name": sheet_name,
                "source_column": column,
                "decision": "사용자가 입력한 값 순서대로 행을 안정 정렬했습니다.",
                "review_note": self._review_note(order_map, unmatched_values, missing_order_values),
            }
        ]
        return JobResult(result_frame=sorted_frame, summary=summary, mapping_records=mapping_records)

    def sort_file(
        self,
        path: str | Path,
        column: str,
        order_values: Iterable[object],
        load_table,
    ) -> JobResult:
        payload = load_table(path)
        if isinstance(payload, tuple) and len(payload) == 2:
            table, metadata = payload
        else:
            table = payload
            metadata = {"file_name": Path(path).name}
        return self.sort_frame(table, column, order_values, metadata=metadata)

    def _order_map(self, order_values: Iterable[object]) -> dict[str, int]:
        order_map: dict[str, int] = {}
        for value in order_values:
            normalized = normalize_key_value(value)
            if normalized and normalized not in order_map:
                order_map[normalized] = len(order_map)
        return order_map

    def _review_note(
        self,
        order_map: dict[str, int],
        unmatched_values: list[str],
        missing_order_values: list[str],
    ) -> str:
        ordered_preview = ", ".join(list(order_map)[:20])
        notes = [f"입력 순서: {ordered_preview}"]
        if len(order_map) > 20:
            notes.append(f"외 {len(order_map) - 20}개")
        if unmatched_values:
            notes.append(f"순서에 없는 값 {len(unmatched_values)}개는 뒤쪽에 유지")
        if missing_order_values:
            notes.append(f"자료에 없는 입력값 {len(missing_order_values)}개")
        return " / ".join(notes)
