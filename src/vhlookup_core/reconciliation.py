from __future__ import annotations

import pandas as pd

from vhlookup_core.keys import build_key_series
from vhlookup_core.models import JobResult, KeySpec, ValidationIssue


class ReconciliationEngine:
    def compare_lists(
        self,
        reference: pd.DataFrame,
        target: pd.DataFrame,
        key_spec: KeySpec,
        reference_label: str = "기준표",
        target_label: str = "대상표",
    ) -> JobResult:
        reference_key_columns = key_spec.reference_key_columns
        target_key_columns = key_spec.target_columns()
        reference_keys = build_key_series(reference, reference_key_columns, key_spec.normalization)
        target_keys = build_key_series(target, target_key_columns, key_spec.normalization)

        ref = reference.copy()
        tgt = target.copy()
        ref["__vh_key"] = reference_keys
        tgt["__vh_key"] = target_keys

        issues: list[ValidationIssue] = []
        self._append_missing_key_issues(issues, ref, "reference_missing_required_key", f"{reference_label} 키 값이 비어 있습니다.")
        self._append_missing_key_issues(issues, tgt, "target_missing_required_key", f"{target_label} 키 값이 비어 있습니다.")
        self._append_duplicate_issues(issues, ref, "reference_duplicate_key", f"{reference_label}에 같은 키가 2건 이상 있습니다.")
        self._append_duplicate_issues(issues, tgt, "target_duplicate_key", f"{target_label}에 같은 키가 2건 이상 있습니다.")

        reference_key_set = {key for key in ref["__vh_key"] if key != ""}
        target_key_set = {key for key in tgt["__vh_key"] if key != ""}
        only_reference = ref[ref["__vh_key"].isin(reference_key_set - target_key_set)].copy()
        only_target = tgt[tgt["__vh_key"].isin(target_key_set - reference_key_set)].copy()
        matched = ref[ref["__vh_key"].isin(reference_key_set & target_key_set)].copy()

        only_reference.insert(0, "대조 상태", f"{reference_label}에만 있음")
        only_target.insert(0, "대조 상태", f"{target_label}에만 있음")
        matched.insert(0, "대조 상태", "양쪽 모두 있음")

        for row_index, row in only_reference.iterrows():
            issues.append(
                ValidationIssue(
                    issue_type="missing_in_target",
                    message=f"{target_label}에 없습니다.",
                    row_number=int(row_index) + 2,
                )
            )
        for row_index, row in only_target.iterrows():
            issues.append(
                ValidationIssue(
                    issue_type="missing_in_reference",
                    message=f"{reference_label}에 없습니다.",
                    row_number=int(row_index) + 2,
                )
            )

        result = pd.concat([only_reference, only_target, matched], ignore_index=True, sort=False)
        if "__vh_key" in result.columns:
            result = result.drop(columns=["__vh_key"])

        summary = {
            "reference_rows": len(reference),
            "target_rows": len(target),
            "matched_rows": len(matched),
            "missing_in_target_rows": len(only_reference),
            "missing_in_reference_rows": len(only_target),
            "issue_count": len(issues),
        }
        return JobResult(result_frame=result, issues=issues, summary=summary)

    def _append_missing_key_issues(
        self,
        issues: list[ValidationIssue],
        frame: pd.DataFrame,
        issue_type: str,
        message: str,
    ) -> None:
        for row_index, key in frame["__vh_key"].items():
            if key == "":
                issues.append(
                    ValidationIssue(
                        issue_type=issue_type,
                        message=message,
                        row_number=int(row_index) + 2,
                    )
                )

    def _append_duplicate_issues(
        self,
        issues: list[ValidationIssue],
        frame: pd.DataFrame,
        issue_type: str,
        message: str,
    ) -> None:
        duplicate_keys = set(frame.loc[frame["__vh_key"].duplicated(keep=False) & (frame["__vh_key"] != ""), "__vh_key"])
        for row_index, row in frame.iterrows():
            if row["__vh_key"] in duplicate_keys:
                issues.append(
                    ValidationIssue(
                        issue_type=issue_type,
                        message=message,
                        row_number=int(row_index) + 2,
                    )
                )
