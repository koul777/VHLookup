from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Literal

import pandas as pd

from vhlookup_core.auto import AutoLookupPlanner
from vhlookup_core.header import HeaderDetector
from vhlookup_core.loader import ExcelLoader
from vhlookup_core.mapper import ColumnMapper
from vhlookup_core.merge import MergeEngine
from vhlookup_core.models import JobResult, KeySpec, ValidationIssue
from vhlookup_core.sheet import SheetDetector
from vhlookup_core.templates import WorkflowTemplate, get_template
from vhlookup_core.validation import ValidationEngine, ValidationProfile


class ConsolidationEngine:
    supported_suffixes = {".xlsx", ".xlsm", ".csv"}

    def __init__(
        self,
        loader: ExcelLoader | None = None,
        sheet_detector: SheetDetector | None = None,
        header_detector: HeaderDetector | None = None,
        mapper: ColumnMapper | None = None,
        validation_engine: ValidationEngine | None = None,
    ) -> None:
        self.loader = loader or ExcelLoader()
        self.header_detector = header_detector or HeaderDetector()
        self.sheet_detector = sheet_detector or SheetDetector(self.header_detector)
        self.mapper = mapper or ColumnMapper()
        self.validation_engine = validation_engine or ValidationEngine()

    def recommend_merge_mode(
        self,
        files: list[str | Path],
        scan_rows: int | str = 30,
    ) -> Literal["rows", "columns"]:
        loaded: list[pd.DataFrame] = []
        for file in files:
            try:
                source = self.loader.load(Path(file))
                sheet = self.sheet_detector.select(source)
                detection = self.header_detector.detect(sheet, scan_rows=scan_rows)
                loaded.append(self.header_detector.apply(sheet, detection))
            except Exception:
                return "rows"
        if len(loaded) < 2:
            return "rows"

        base = loaded[0]
        planner = AutoLookupPlanner(mapper=self.mapper)
        column_like_count = 0
        for table in loaded[1:]:
            schema_mapping = self.mapper.map_columns(list(table.columns), list(base.columns), threshold=0.70)
            schema_ratio = len(schema_mapping.source_to_target) / max(len(table.columns), len(base.columns), 1)
            if schema_ratio >= 0.70:
                return "rows"
            try:
                plan = planner.infer_lookup_plan(table, base)
            except Exception:
                return "rows"
            if plan.confidence < 0.55 or not plan.value_columns:
                return "rows"
            column_like_count += 1

        return "columns" if column_like_count == len(loaded) - 1 else "rows"

    def consolidate_folder(
        self,
        folder: str | Path,
        standard_columns: list[str] | None = None,
        scan_rows: int | str = 30,
        template: WorkflowTemplate | str | None = None,
    ) -> JobResult:
        folder_path = Path(folder)
        if not folder_path.exists():
            raise FileNotFoundError(folder_path)
        files = sorted(
            path for path in folder_path.iterdir() if path.is_file() and path.suffix.lower() in self.supported_suffixes
        )
        return self.consolidate_files(files, standard_columns=standard_columns, scan_rows=scan_rows, template=template)

    def consolidate_files(
        self,
        files: list[str | Path],
        standard_columns: list[str] | None = None,
        scan_rows: int | str = 30,
        template: WorkflowTemplate | str | None = None,
        saved_mappings_by_file: dict[str, dict[str, str]] | None = None,
    ) -> JobResult:
        issues: list[ValidationIssue] = []
        frames: list[pd.DataFrame] = []
        mapping_records: list[dict[str, object]] = []
        produced_columns_by_file: dict[str, set[str]] = {}
        sheet_by_file: dict[str, str] = {}
        workflow_template = get_template(template) if isinstance(template, str) else template
        validation_profile = ValidationProfile.from_template(workflow_template)
        if standard_columns is None and workflow_template is not None and workflow_template.standard_columns:
            standard_columns = list(workflow_template.standard_columns)
        resolved_standard_columns = standard_columns
        mapping_count = 0

        for file in files:
            path = Path(file)
            try:
                source = self.loader.load(path)
                sheet = self.sheet_detector.select(source)
                detection = self.header_detector.detect(sheet, scan_rows=scan_rows)
                table = self.header_detector.apply(sheet, detection)
                if resolved_standard_columns is None:
                    resolved_standard_columns = list(table.columns)
                saved_mappings = self._mapping_for_file(saved_mappings_by_file, path)
                mapping = self.mapper.map_columns(
                    list(table.columns),
                    resolved_standard_columns,
                    saved_mappings=saved_mappings,
                )
                mapping_count += len(mapping.source_to_target)
                mapping_records.append(
                    {
                        "역할": "헤더 자동탐지",
                        "파일명": path.name,
                        "시트명": sheet.name,
                        "원본 컬럼": "",
                        "표준 컬럼": "",
                        "신뢰도": detection.confidence,
                        "추천 방식": f"{detection.header_row_number}행을 헤더로 자동 선택",
                        "검토 메모": "; ".join(detection.warnings),
                    }
                )
                for source_column, target_column in mapping.source_to_target.items():
                    mapping_records.append(
                        {
                            "역할": "컬럼 매칭",
                            "파일명": path.name,
                            "시트명": sheet.name,
                            "원본 컬럼": source_column,
                            "표준 컬럼": target_column,
                            "신뢰도": mapping.confidence_by_target.get(target_column),
                            "추천 방식": "컬럼명 동의어/유사도",
                            "검토 메모": "",
                        }
                    )
                renamed = table.rename(columns=mapping.source_to_target)
                output = pd.DataFrame(index=table.index)
                produced_columns: set[str] = set()
                for column in resolved_standard_columns:
                    if column in renamed.columns:
                        output[column] = renamed[column]
                        produced_columns.add(str(column))
                    else:
                        output[column] = pd.NA
                preserved_columns = self._append_unmapped_source_columns(
                    output,
                    table,
                    mapped_sources=set(mapping.source_to_target),
                )
                produced_columns.update(str(output_column) for _source_column, output_column in preserved_columns)
                for source_column, output_column in preserved_columns:
                    mapping_records.append(
                        {
                            "역할": "원본 컬럼 보존",
                            "파일명": path.name,
                            "시트명": sheet.name,
                            "원본 컬럼": source_column,
                            "표준 컬럼": output_column,
                            "신뢰도": 1.0,
                            "추천 방식": "표준 컬럼에 매칭되지 않은 열은 삭제하지 않고 결과 뒤쪽에 보존",
                            "검토 메모": "",
                        }
                    )
                output.insert(0, "원본 행 번호", range(detection.data_start_row_number, detection.data_start_row_number + len(output)))
                output.insert(0, "원본 시트명", sheet.name)
                output.insert(0, "원본 파일명", path.name)
                produced_columns_by_file[path.name] = produced_columns
                sheet_by_file[path.name] = sheet.name
                frames.append(output)
                for warning in detection.warnings + mapping.warnings:
                    issues.append(
                        ValidationIssue(
                            issue_type="auto_detection_warning",
                            message=warning,
                            severity="info",
                            file_name=path.name,
                            sheet_name=sheet.name,
                        )
                    )
            except Exception as exc:
                issues.append(
                    ValidationIssue(
                        issue_type="file_processing_failed",
                        message="파일 처리 실패",
                        severity="error",
                        file_name=path.name,
                        details={"error": str(exc)},
                    )
                )

        result = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        if not result.empty:
            issues.extend(self.validation_engine.validate_frame(result, validation_profile))
            issues.extend(self._one_sided_column_issues(result, produced_columns_by_file, sheet_by_file))
            issues = self._attach_result_indices(result, issues)
        public_result = self._drop_tracking_columns(result)
        summary = {
            "workflow": workflow_template.name if workflow_template else "사용자 지정 수합",
            "file_count": len(files),
            "successful_file_count": len(frames),
            "row_count": len(public_result),
            "issue_count": len(issues),
            "validation_issue_count": sum(
                1
                for issue in issues
                if issue.issue_type
                in {
                    "required_column_missing",
                    "required_value_missing",
                    "numeric_value_invalid",
                    "date_value_invalid",
                    "duplicate_business_key",
                }
            ),
            "mapped_column_count": mapping_count,
        }
        return JobResult(result_frame=public_result, issues=issues, summary=summary, mapping_records=mapping_records)

    def merge_files_by_columns(
        self,
        files: list[str | Path],
        scan_rows: int | str = 30,
        merge_plans_by_file: dict[str, dict[str, object]] | None = None,
    ) -> JobResult:
        loaded: list[tuple[Path, str, pd.DataFrame]] = []
        issues: list[ValidationIssue] = []
        mapping_records: list[dict[str, object]] = []

        for file in files:
            path = Path(file)
            try:
                source = self.loader.load(path)
                sheet = self.sheet_detector.select(source)
                detection = self.header_detector.detect(sheet, scan_rows=scan_rows)
                table = self.header_detector.apply(sheet, detection)
                loaded.append((path, sheet.name, table))
                mapping_records.append(
                    {
                        "역할": "헤더 자동탐지",
                        "파일명": path.name,
                        "시트명": sheet.name,
                        "원본 컬럼": "",
                        "표준 컬럼": "",
                        "신뢰도": detection.confidence,
                        "추천 방식": f"{detection.header_row_number}행을 헤더로 자동 선택",
                        "검토 메모": "; ".join(detection.warnings),
                    }
                )
            except Exception as exc:
                issues.append(
                    ValidationIssue(
                        issue_type="file_processing_failed",
                        message="파일 처리 실패",
                        severity="error",
                        file_name=path.name,
                        details={"error": str(exc)},
                    )
                )

        if not loaded:
            return JobResult(
                result_frame=pd.DataFrame(),
                issues=issues,
                summary={
                    "workflow": "열 방향 파일 합치기",
                    "file_count": len(files),
                    "successful_file_count": 0,
                    "row_count": 0,
                    "column_count": 0,
                    "issue_count": len(issues),
                },
                mapping_records=mapping_records,
            )

        base_path, base_sheet, base_table = loaded[0]
        result = base_table.reset_index(drop=True).copy()
        mapping_records.append(
            {
                "역할": "열 합치기 기준표",
                "파일명": base_path.name,
                "시트명": base_sheet,
                "원본 컬럼": "",
                "표준 컬럼": "",
                "신뢰도": 1.0,
                "추천 방식": "첫 번째 선택 파일을 기준 표로 사용",
                "검토 메모": "나머지 파일은 자동 키 매칭으로 오른쪽에 붙입니다.",
            }
        )

        planner = AutoLookupPlanner()
        merger = MergeEngine()
        for file_order, (path, sheet_name, table) in enumerate(loaded[1:], start=2):
            reference = table.reset_index(drop=True).copy()
            try:
                manual_plan = self._mapping_for_file(merge_plans_by_file, path)
                if manual_plan:
                    reference_keys = tuple(manual_plan.get("source_key_columns", ()))
                    target_keys = tuple(manual_plan.get("base_key_columns", ()))
                    value_columns = tuple(manual_plan.get("value_columns", ()))
                    if not reference_keys or not target_keys or not value_columns:
                        raise ValueError("열 합치기 수동 설정에 키 또는 가져올 컬럼이 없습니다.")
                    normalization = str(
                        manual_plan.get("normalization")
                        or planner.infer_key_normalization(reference, result, reference_keys, target_keys)
                    )
                    if normalization not in {"text", "loose_numeric"}:
                        normalization = "text"
                    key_spec = KeySpec(
                        reference_key_columns=reference_keys,
                        target_key_columns=target_keys,
                        normalization=normalization,  # type: ignore[arg-type]
                    )
                    merge_result = merger.merge_lookup(
                        reference,
                        result,
                        key_spec,
                        list(value_columns),
                        output_suffix=f"_{path.stem}",
                        include_unmatched_reference=True,
                    )
                    result = merge_result.result_frame
                    issues.extend(
                        self._annotate_column_merge_issues(merge_result.issues, path.name, sheet_name, file_order)
                    )
                    mapping_records.append(
                        {
                            "역할": "사용자 선택 키 컬럼",
                            "파일명": path.name,
                            "시트명": sheet_name,
                            "기준표 컬럼": " + ".join(reference_keys),
                            "대상표 컬럼": " + ".join(target_keys),
                            "신뢰도": 1.0,
                            "추천 방식": "미리보기 화면에서 사용자가 선택",
                            "검토 메모": "앞자리 0 차이는 같은 숫자 키로 맞춰 비교했습니다."
                            if normalization == "loose_numeric"
                            else "",
                        }
                    )
                    for value_column in value_columns:
                        mapping_records.append(
                            {
                                "역할": "사용자 선택 가져올 컬럼",
                                "파일명": path.name,
                                "시트명": sheet_name,
                                "기준표 컬럼": value_column,
                                "대상표 컬럼": "",
                                "신뢰도": 1.0,
                                "추천 방식": "미리보기 화면에서 사용자가 선택",
                                "검토 메모": "",
                            }
                        )
                    continue

                plan = planner.infer_lookup_plan(reference, result)
                merge_result = merger.merge_lookup(
                    reference,
                    result,
                    plan.key_spec,
                    list(plan.value_columns),
                    output_suffix=f"_{path.stem}",
                    include_unmatched_reference=True,
                )
                result = merge_result.result_frame
                issues.extend(
                    self._annotate_column_merge_issues(merge_result.issues, path.name, sheet_name, file_order)
                )
                for warning in plan.warnings:
                    issues.append(
                        ValidationIssue(
                            issue_type="auto_key_warning",
                            message=warning,
                            severity="warning",
                            file_name=path.name,
                            sheet_name=sheet_name,
                        )
                    )
                for row in plan.evidence_rows:
                    evidence = dict(row)
                    evidence["파일명"] = path.name
                    evidence["시트명"] = sheet_name
                    mapping_records.append(evidence)
            except Exception as exc:
                result = self._append_columns_by_position(result, reference, path.stem)
                issues.append(
                    ValidationIssue(
                        issue_type="column_merge_by_position",
                        message="자동 키를 찾지 못해 행 순서대로 오른쪽에 붙였습니다.",
                        severity="warning",
                        file_name=path.name,
                        sheet_name=sheet_name,
                        details={"reason": str(exc)},
                    )
                )
                mapping_records.append(
                    {
                        "역할": "행 순서 기준 열 합치기",
                        "파일명": path.name,
                        "시트명": sheet_name,
                        "원본 컬럼": ", ".join(map(str, reference.columns)),
                        "표준 컬럼": "",
                        "신뢰도": 0.0,
                        "추천 방식": "자동 키 매칭 실패 후 행 번호 순서로 오른쪽에 붙임",
                        "검토 메모": str(exc),
                    }
                )

        summary = {
            "workflow": "열 방향 파일 합치기",
            "file_count": len(files),
            "successful_file_count": len(loaded),
            "row_count": len(result),
            "column_count": len(result.columns),
            "issue_count": len(issues),
        }
        return JobResult(result_frame=result, issues=issues, summary=summary, mapping_records=mapping_records)

    def _annotate_column_merge_issues(
        self,
        issues: list[ValidationIssue],
        file_name: str,
        sheet_name: str,
        file_order: int,
    ) -> list[ValidationIssue]:
        file_label = f"{file_order}번째 파일({file_name})"
        annotated: list[ValidationIssue] = []
        for issue in issues:
            updates = {"file_name": file_name, "sheet_name": sheet_name}
            if issue.issue_type == "match_failed":
                updates["message"] = f"{file_label}에 해당 기준값이 없습니다."
            elif issue.issue_type == "reference_only_unmatched":
                updates["message"] = f"{file_label}에만 있는 기준값입니다. 다른 파일에는 없으니 확인하세요."
            elif issue.issue_type == "format_mismatch":
                updates["message"] = f"{file_label}의 기준값과 표시 형식이 달라 매칭에 실패했습니다."
            elif issue.issue_type == "reference_duplicate_key_blocked":
                updates["message"] = f"{file_label}에 같은 기준값이 여러 건 있어 자동 병합하지 않았습니다."
            annotated.append(replace(issue, **updates))
        return annotated

    def _one_sided_column_issues(
        self,
        result: pd.DataFrame,
        produced_columns_by_file: dict[str, set[str]],
        sheet_by_file: dict[str, str],
    ) -> list[ValidationIssue]:
        if result.empty or len(produced_columns_by_file) < 2 or "원본 파일명" not in result.columns:
            return []
        all_columns: set[str] = set()
        for columns in produced_columns_by_file.values():
            all_columns.update(columns)

        issues: list[ValidationIssue] = []
        file_series = result["원본 파일명"].astype(str)
        for file_name, produced_columns in produced_columns_by_file.items():
            indices = [int(index) for index in result.index[file_series == file_name]]
            if not indices:
                continue
            missing_columns = sorted(column for column in all_columns if column not in produced_columns)
            for column in missing_columns:
                issues.append(
                    ValidationIssue(
                        issue_type="column_present_in_other_file",
                        message=f"{column} 열은 다른 파일에만 있어 이 파일의 행은 빈칸입니다.",
                        severity="warning",
                        file_name=file_name,
                        sheet_name=sheet_by_file.get(file_name),
                        details={"result_columns": [column], "result_indices": indices},
                    )
                )
        return issues

    def _drop_tracking_columns(self, frame: pd.DataFrame) -> pd.DataFrame:
        tracking_columns = ["원본 파일명", "원본 시트명", "원본 행 번호"]
        return frame.drop(columns=[column for column in tracking_columns if column in frame.columns])

    def _attach_result_indices(
        self,
        result: pd.DataFrame,
        issues: list[ValidationIssue],
    ) -> list[ValidationIssue]:
        if result.empty:
            return issues
        annotated: list[ValidationIssue] = []
        for issue in issues:
            if isinstance(issue.details, dict) and "result_indices" in issue.details:
                annotated.append(issue)
                continue
            indices = self._matching_tracking_indices(result, issue)
            if not indices:
                annotated.append(issue)
                continue
            details = dict(issue.details)
            details["result_indices"] = indices
            annotated.append(replace(issue, details=details))
        return annotated

    def _matching_tracking_indices(self, result: pd.DataFrame, issue: ValidationIssue) -> list[int]:
        if issue.row_number is None:
            return []
        mask = pd.Series(True, index=result.index)
        matched = False
        if "원본 행 번호" in result.columns:
            mask &= result["원본 행 번호"].astype(str) == str(issue.row_number)
            matched = True
        if issue.file_name and "원본 파일명" in result.columns:
            mask &= result["원본 파일명"].astype(str) == str(issue.file_name)
            matched = True
        if issue.sheet_name and "원본 시트명" in result.columns:
            mask &= result["원본 시트명"].astype(str) == str(issue.sheet_name)
            matched = True
        if not matched:
            return []
        return [int(index) for index in result.index[mask]]

    def _append_columns_by_position(
        self,
        result: pd.DataFrame,
        source: pd.DataFrame,
        file_stem: str,
    ) -> pd.DataFrame:
        output = result.reset_index(drop=True).copy()
        source_by_position = source.reset_index(drop=True).copy()
        max_rows = max(len(output), len(source_by_position))
        output = output.reindex(range(max_rows))
        source_by_position = source_by_position.reindex(range(max_rows))
        existing_columns = set(map(str, output.columns))
        for column in source_by_position.columns:
            output_column = self._unique_output_column_name(f"{file_stem}_{column}", existing_columns)
            output[output_column] = source_by_position[column]
            existing_columns.add(output_column)
        return output

    def _mapping_for_file(
        self,
        mappings_by_file: dict[str, dict[str, object]] | dict[str, dict[str, str]] | None,
        path: Path,
    ) -> dict[str, object]:
        if not mappings_by_file:
            return {}
        candidates = (str(path), str(path.resolve()), path.name)
        for candidate in candidates:
            mapping = mappings_by_file.get(candidate)
            if mapping:
                return dict(mapping)
        return {}

    def _append_unmapped_source_columns(
        self,
        output: pd.DataFrame,
        source: pd.DataFrame,
        mapped_sources: set[str],
    ) -> list[tuple[str, str]]:
        preserved: list[tuple[str, str]] = []
        for source_column in source.columns:
            if source_column in mapped_sources:
                continue
            output_column = self._unique_output_column_name(str(source_column), set(map(str, output.columns)))
            output[output_column] = source[source_column]
            preserved.append((str(source_column), output_column))
        return preserved

    def _unique_output_column_name(self, column: str, existing_columns: set[str]) -> str:
        if column not in existing_columns:
            return column
        base = f"원본_{column}"
        if base not in existing_columns:
            return base
        suffix = 2
        while f"{base}_{suffix}" in existing_columns:
            suffix += 1
        return f"{base}_{suffix}"
