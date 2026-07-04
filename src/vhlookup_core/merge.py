from __future__ import annotations

import pandas as pd

from vhlookup_core.keys import build_key_series
from vhlookup_core.models import JobResult, KeySpec, ValidationIssue
from vhlookup_core.normalization import normalize_key_value


class MergeEngine:
    def merge_lookup(
        self,
        reference: pd.DataFrame,
        target: pd.DataFrame,
        key_spec: KeySpec,
        value_columns: list[str] | tuple[str, ...],
        output_suffix: str = "_from_reference",
        include_unmatched_reference: bool = False,
    ) -> JobResult:
        missing_values = [column for column in value_columns if column not in reference.columns]
        if missing_values:
            raise KeyError(f"Missing reference value columns: {missing_values}")

        ref_key_columns = key_spec.reference_key_columns
        target_key_columns = key_spec.target_columns()
        reference_keys = build_key_series(reference, ref_key_columns, key_spec.normalization)
        target_keys = build_key_series(target, target_key_columns, key_spec.normalization)

        issues: list[ValidationIssue] = []
        reference_with_key = reference.copy()
        target_with_key = target.copy()
        reference_with_key["__vh_key"] = reference_keys
        target_with_key["__vh_key"] = target_keys

        duplicate_reference_keys = set(
            reference_with_key.loc[
                reference_with_key["__vh_key"].duplicated(keep=False) & (reference_with_key["__vh_key"] != ""),
                "__vh_key",
            ]
        )
        duplicate_target_keys = set(
            target_with_key.loc[
                target_with_key["__vh_key"].duplicated(keep=False) & (target_with_key["__vh_key"] != ""),
                "__vh_key",
            ]
        )

        self._append_duplicate_issues(
            issues,
            reference_with_key,
            duplicate_reference_keys,
            "reference_duplicate_key",
            "같은 키가 기준표에 2건 이상 존재합니다.",
        )
        self._append_duplicate_issues(
            issues,
            target_with_key,
            duplicate_target_keys,
            "target_duplicate_key",
            "같은 키가 대상표에 2건 이상 존재합니다.",
        )

        unique_reference = reference_with_key[
            ~reference_with_key["__vh_key"].isin(duplicate_reference_keys)
        ].copy()
        attach_columns = ["__vh_key", *value_columns]
        lookup_frame = unique_reference.loc[:, attach_columns].copy()
        rename_map = {
            column: self._output_column_name(column, target.columns, output_suffix)
            for column in value_columns
        }
        output_value_columns = list(rename_map.values())
        lookup_frame = lookup_frame.rename(columns=rename_map)

        result = target_with_key.merge(lookup_frame, on="__vh_key", how="left", sort=False)
        reference_key_set = set(unique_reference["__vh_key"])
        target_key_set = set(target_with_key["__vh_key"])
        duplicate_reference_key_set = duplicate_reference_keys

        for row_index, key in target_with_key["__vh_key"].items():
            row_number = int(row_index) + 2
            if key == "":
                issues.append(
                    ValidationIssue(
                        issue_type="missing_required_key",
                        message="대상표 키 값이 비어 있습니다.",
                        row_number=row_number,
                        details={"result_columns": output_value_columns},
                    )
                )
            elif key in duplicate_reference_key_set:
                issues.append(
                    ValidationIssue(
                        issue_type="reference_duplicate_key_blocked",
                        message="기준표 중복 키로 인해 자동 병합하지 않았습니다.",
                        row_number=row_number,
                        details={"result_columns": output_value_columns},
                    )
                )
            elif key not in reference_key_set:
                mismatch = self._find_format_mismatch(
                    key,
                    target_with_key.loc[row_index, list(target_key_columns)].tolist(),
                    reference_with_key,
                    ref_key_columns,
                )
                if mismatch:
                    issues.append(
                        ValidationIssue(
                            issue_type="format_mismatch",
                            message="값 형식이 달라 매칭에 실패했습니다.",
                            row_number=row_number,
                            details={**mismatch, "result_columns": output_value_columns},
                        )
                    )
                else:
                    issues.append(
                        ValidationIssue(
                            issue_type="match_failed",
                            message="기준표에 없습니다.",
                            row_number=row_number,
                            details={"result_columns": output_value_columns},
                        )
                    )

        if include_unmatched_reference:
            reference_only_columns = [
                str(column)
                for column in target.columns
                if str(column).strip() and str(column) not in set(target_key_columns)
            ]
            reference_only_rows: list[dict[object, object]] = []
            for row_index, row in unique_reference.iterrows():
                key = row["__vh_key"]
                if key == "" or key in target_key_set or key in duplicate_reference_key_set:
                    continue
                output_row = {column: pd.NA for column in result.columns}
                output_row["__vh_key"] = key
                for reference_key_column, target_key_column in zip(ref_key_columns, target_key_columns):
                    if target_key_column in output_row:
                        output_row[target_key_column] = row[reference_key_column]
                for reference_column, output_column in rename_map.items():
                    output_row[output_column] = row[reference_column]
                reference_only_rows.append(output_row)
                issues.append(
                    ValidationIssue(
                        issue_type="reference_only_unmatched",
                        message="붙일 파일에는 있지만 대상표에 해당 기준값이 없어 행을 추가했습니다.",
                        row_number=len(result) + len(reference_only_rows) + 1,
                        details={
                            "source_row_number": int(row_index) + 2,
                            "source_key": key,
                            "result_columns": reference_only_columns,
                        },
                    )
                )
            if reference_only_rows:
                result = pd.concat(
                    [result, pd.DataFrame(reference_only_rows, columns=result.columns)],
                    ignore_index=True,
                )

        result = result.drop(columns=["__vh_key"])
        summary = {
            "target_rows": len(target),
            "reference_rows": len(reference),
            "matched_rows": int(result[list(rename_map.values())].notna().any(axis=1).sum())
            if rename_map
            else 0,
            "reference_only_rows": sum(1 for issue in issues if issue.issue_type == "reference_only_unmatched"),
            "issue_count": len(issues),
        }
        return JobResult(result_frame=result, issues=issues, summary=summary)

    def _append_duplicate_issues(
        self,
        issues: list[ValidationIssue],
        frame: pd.DataFrame,
        duplicate_keys: set[str],
        issue_type: str,
        message: str,
    ) -> None:
        if not duplicate_keys:
            return
        for row_index, row in frame.iterrows():
            if row["__vh_key"] in duplicate_keys:
                issues.append(
                    ValidationIssue(
                        issue_type=issue_type,
                        message=message,
                        row_number=int(row_index) + 2,
                    )
                )

    def _find_format_mismatch(
        self,
        key: str,
        target_values: list[object],
        reference: pd.DataFrame,
        reference_key_columns: tuple[str, ...],
    ) -> dict[str, str] | None:
        target_loose = "\x1f".join(normalize_key_value(value, "loose_numeric") for value in target_values)
        if target_loose == "":
            return None
        for _, row in reference.iterrows():
            reference_values = row.loc[list(reference_key_columns)].tolist()
            reference_loose = "\x1f".join(
                normalize_key_value(value, "loose_numeric") for value in reference_values
            )
            if reference_loose == target_loose and row["__vh_key"] != key:
                return {"target_key": key, "reference_key": row["__vh_key"]}
        return None

    def _output_column_name(
        self,
        column: str,
        target_columns: pd.Index,
        output_suffix: str,
    ) -> str:
        if column not in target_columns:
            return column
        return f"{column}{output_suffix}"
