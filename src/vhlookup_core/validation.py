from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

from vhlookup_core.models import ValidationIssue
from vhlookup_core.normalization import display_value, is_blank, normalize_key_parts
from vhlookup_core.templates import WorkflowTemplate


_NUMERIC_RE = re.compile(r"^-?\d{1,3}(,\d{3})*(\.\d+)?$|^-?\d+(\.\d+)?$")
_YYYY_MM_RE = re.compile(r"^\d{4}[-./년]\s?(0?[1-9]|1[0-2])\s?(월)?$")
_MONTH_RE = re.compile(r"^(0?[1-9]|1[0-2])\s?월$")


@dataclass(frozen=True)
class ValidationProfile:
    required_columns: tuple[str, ...] = ()
    numeric_columns: tuple[str, ...] = ()
    date_columns: tuple[str, ...] = ()
    duplicate_key_columns: tuple[str, ...] = ()

    @classmethod
    def from_template(cls, template: WorkflowTemplate | None) -> "ValidationProfile":
        if template is None:
            return cls()
        return cls(
            required_columns=template.required_columns,
            numeric_columns=template.numeric_columns,
            date_columns=template.date_columns,
            duplicate_key_columns=template.duplicate_key_columns,
        )


class ValidationEngine:
    def validate_frame(
        self,
        frame: pd.DataFrame,
        profile: ValidationProfile,
        file_name: str | None = None,
        sheet_name: str | None = None,
        row_number_column: str = "원본 행 번호",
    ) -> list[ValidationIssue]:
        issues: list[ValidationIssue] = []
        issues.extend(self._validate_required(frame, profile.required_columns, file_name, sheet_name, row_number_column))
        issues.extend(self._validate_numeric(frame, profile.numeric_columns, file_name, sheet_name, row_number_column))
        issues.extend(self._validate_date(frame, profile.date_columns, file_name, sheet_name, row_number_column))
        issues.extend(
            self._validate_duplicates(
                frame,
                profile.duplicate_key_columns,
                file_name,
                sheet_name,
                row_number_column,
            )
        )
        return issues

    def _validate_required(
        self,
        frame: pd.DataFrame,
        columns: tuple[str, ...],
        file_name: str | None,
        sheet_name: str | None,
        row_number_column: str,
    ) -> list[ValidationIssue]:
        issues = []
        for column in columns:
            if column not in frame.columns:
                issues.append(
                    ValidationIssue(
                        issue_type="required_column_missing",
                        message="필수 컬럼이 없습니다.",
                        severity="error",
                        file_name=file_name,
                        sheet_name=sheet_name,
                        column_name=column,
                    )
                )
                continue
            for index, value in frame[column].items():
                if is_blank(value):
                    issues.append(
                        ValidationIssue(
                            issue_type="required_value_missing",
                            message="필수값이 비어 있습니다.",
                            file_name=file_name or self._value_at(frame, index, "원본 파일명"),
                            sheet_name=sheet_name or self._value_at(frame, index, "원본 시트명"),
                            row_number=self._row_number(frame, index, row_number_column),
                            column_name=column,
                        )
                    )
        return issues

    def _validate_numeric(
        self,
        frame: pd.DataFrame,
        columns: tuple[str, ...],
        file_name: str | None,
        sheet_name: str | None,
        row_number_column: str,
    ) -> list[ValidationIssue]:
        issues = []
        for column in columns:
            if column not in frame.columns:
                continue
            for index, value in frame[column].items():
                if is_blank(value):
                    continue
                text = display_value(value).replace("원", "").strip()
                if not _NUMERIC_RE.match(text):
                    issues.append(
                        ValidationIssue(
                            issue_type="numeric_value_invalid",
                            message="금액/수량 컬럼에 숫자가 아닌 값이 있습니다.",
                            file_name=file_name or self._value_at(frame, index, "원본 파일명"),
                            sheet_name=sheet_name or self._value_at(frame, index, "원본 시트명"),
                            row_number=self._row_number(frame, index, row_number_column),
                            column_name=column,
                        )
                    )
        return issues

    def _validate_date(
        self,
        frame: pd.DataFrame,
        columns: tuple[str, ...],
        file_name: str | None,
        sheet_name: str | None,
        row_number_column: str,
    ) -> list[ValidationIssue]:
        issues = []
        for column in columns:
            if column not in frame.columns:
                continue
            parsed = pd.to_datetime(frame[column], errors="coerce")
            for index, value in frame[column].items():
                if is_blank(value):
                    continue
                text = display_value(value)
                if _YYYY_MM_RE.match(text) or _MONTH_RE.match(text):
                    continue
                if pd.isna(parsed.loc[index]):
                    issues.append(
                        ValidationIssue(
                            issue_type="date_value_invalid",
                            message="날짜로 해석할 수 없는 값입니다.",
                            file_name=file_name or self._value_at(frame, index, "원본 파일명"),
                            sheet_name=sheet_name or self._value_at(frame, index, "원본 시트명"),
                            row_number=self._row_number(frame, index, row_number_column),
                            column_name=column,
                        )
                    )
        return issues

    def _validate_duplicates(
        self,
        frame: pd.DataFrame,
        columns: tuple[str, ...],
        file_name: str | None,
        sheet_name: str | None,
        row_number_column: str,
    ) -> list[ValidationIssue]:
        if not columns or any(column not in frame.columns for column in columns):
            return []
        keys = frame.loc[:, list(columns)].apply(lambda row: normalize_key_parts(row.tolist()), axis=1)
        duplicate_keys = set(keys[keys.duplicated(keep=False) & (keys != "")])
        issues = []
        for index, key in keys.items():
            if key in duplicate_keys:
                issues.append(
                    ValidationIssue(
                        issue_type="duplicate_business_key",
                        message="같은 기준의 제출 행이 2건 이상 있습니다.",
                        file_name=file_name or self._value_at(frame, index, "원본 파일명"),
                        sheet_name=sheet_name or self._value_at(frame, index, "원본 시트명"),
                        row_number=self._row_number(frame, index, row_number_column),
                        column_name=" + ".join(columns),
                    )
                )
        return issues

    def _row_number(self, frame: pd.DataFrame, index: int, column: str) -> int | None:
        if column in frame.columns and not is_blank(frame.loc[index, column]):
            try:
                return int(frame.loc[index, column])
            except (TypeError, ValueError):
                return None
        return int(index) + 2

    def _value_at(self, frame: pd.DataFrame, index: int, column: str) -> str | None:
        if column not in frame.columns:
            return None
        value = frame.loc[index, column]
        return None if is_blank(value) else display_value(value)
