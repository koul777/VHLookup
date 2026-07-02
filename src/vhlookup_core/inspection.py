from __future__ import annotations

import warnings
from pathlib import Path

import pandas as pd

from vhlookup_core.header import HeaderDetector
from vhlookup_core.loader import ExcelLoader
from vhlookup_core.mapper import ColumnMapper
from vhlookup_core.models import ValidationIssue
from vhlookup_core.normalization import display_value, is_blank, normalize_header
from vhlookup_core.privacy import PrivacyScanner
from vhlookup_core.sheet import SheetDetector
from vhlookup_core.templates import WorkflowTemplate, get_template


class InspectionEngine:
    supported_suffixes = {".xlsx", ".xlsm", ".csv"}

    def __init__(
        self,
        loader: ExcelLoader | None = None,
        header_detector: HeaderDetector | None = None,
        sheet_detector: SheetDetector | None = None,
        mapper: ColumnMapper | None = None,
    ) -> None:
        self.loader = loader or ExcelLoader()
        self.header_detector = header_detector or HeaderDetector()
        self.sheet_detector = sheet_detector or SheetDetector(self.header_detector)
        self.mapper = mapper or ColumnMapper()

    def inspect_path(
        self,
        path: str | Path,
        template: WorkflowTemplate | str | None = None,
        scan_rows: int | str = 30,
    ) -> "InspectionResult":
        source_path = Path(path)
        workflow_template = get_template(template) if isinstance(template, str) else template
        if source_path.is_dir():
            files = sorted(
                file
                for file in source_path.iterdir()
                if file.is_file() and file.suffix.lower() in self.supported_suffixes
            )
        else:
            files = [source_path]

        return self.inspect_files(
            files,
            template=workflow_template,
            scan_rows=scan_rows,
            input_label=str(source_path),
        )

    def inspect_files(
        self,
        files: list[str | Path],
        template: WorkflowTemplate | str | None = None,
        scan_rows: int | str = 30,
        input_label: str | None = None,
    ) -> "InspectionResult":
        workflow_template = get_template(template) if isinstance(template, str) else template
        file_rows: list[dict[str, object]] = []
        sheet_rows: list[dict[str, object]] = []
        column_rows: list[dict[str, object]] = []
        column_profile_rows: list[dict[str, object]] = []
        privacy_rows: list[dict[str, object]] = []
        source_tables: dict[str, pd.DataFrame] = {}
        columns_by_file: dict[str, list[str]] = {}
        issues: list[ValidationIssue] = []
        file_paths = [Path(file) for file in files]
        privacy_scanner = PrivacyScanner()

        for file in file_paths:
            try:
                source = self.loader.load(file)
                selected_sheet = self.sheet_detector.select(source)
                selected_table = pd.DataFrame()
                selected_detection = None
                for sheet in source.sheets:
                    detection = self.header_detector.detect(sheet, scan_rows=scan_rows)
                    is_selected = sheet.name == selected_sheet.name
                    if is_selected:
                        selected_detection = detection
                        selected_table = self.header_detector.apply(sheet, detection)
                        source_tables[file.name] = selected_table.copy()
                        columns_by_file[file.name] = [str(column) for column in selected_table.columns]
                        column_profile_rows.extend(
                            self._column_profile_rows(file.name, sheet.name, selected_table)
                        )
                        for record in privacy_scanner.scan_frame(selected_table):
                            privacy_rows.append({"파일명": file.name, "시트명": sheet.name, **record})
                    sheet_rows.append(
                        {
                            "파일명": file.name,
                            "시트명": sheet.name,
                            "자동 선택": "예" if is_selected else "아니오",
                            "행 수": len(sheet.frame.index),
                            "열 수": len(sheet.frame.columns),
                            "헤더 행": detection.header_row_number,
                            "데이터 시작 행": detection.data_start_row_number,
                            "헤더 신뢰도": detection.confidence,
                            "인식 컬럼": ", ".join(detection.headers),
                            "확인 메모": "; ".join(detection.warnings),
                        }
                    )
                    if is_selected and workflow_template is not None and workflow_template.standard_columns:
                        mapping = self.mapper.map_columns(list(detection.headers), list(workflow_template.standard_columns))
                        for source_column, target_column in mapping.source_to_target.items():
                            column_rows.append(
                                {
                                    "파일명": file.name,
                                    "시트명": sheet.name,
                                    "원본 컬럼": source_column,
                                    "표준 컬럼": target_column,
                                    "신뢰도": mapping.confidence_by_target.get(target_column),
                                }
                            )
                        for target_column in mapping.unmapped_targets:
                            column_rows.append(
                                {
                                    "파일명": file.name,
                                    "시트명": sheet.name,
                                    "원본 컬럼": "",
                                    "표준 컬럼": target_column,
                                    "신뢰도": 0.0,
                                }
                            )
                file_rows.append(
                    {
                        "파일명": file.name,
                        "경로": str(file),
                        "시트 수": len(source.sheets),
                        "자동 선택 시트": selected_sheet.name,
                        "헤더 행": selected_detection.header_row_number if selected_detection else "",
                        "데이터 행 수": len(selected_table),
                        "인식 컬럼 수": len(selected_table.columns),
                        "처리 상태": "성공",
                    }
                )
            except Exception as exc:
                file_rows.append(
                    {
                        "파일명": file.name,
                        "경로": str(file),
                        "시트 수": 0,
                        "자동 선택 시트": "",
                        "처리 상태": "실패",
                    }
                )
                issues.append(
                    ValidationIssue(
                        issue_type="file_processing_failed",
                        message="파일 처리 실패",
                        severity="error",
                        file_name=file.name,
                        details={"error": str(exc)},
                    )
                )

        column_comparison = self._column_comparison_frame(columns_by_file)
        if not column_comparison.empty and column_comparison["누락 파일"].map(lambda value: bool(str(value).strip())).any():
            issues.append(
                ValidationIssue(
                    issue_type="column_schema_mismatch",
                    message="파일별 열 구성이 서로 다릅니다. 열비교 시트를 확인하세요.",
                    severity="warning",
                )
            )

        common_column_count = 0
        if not column_comparison.empty:
            common_column_count = int((column_comparison["누락 파일"].astype(str).str.len() == 0).sum())

        return InspectionResult(
            files=pd.DataFrame(file_rows),
            sheets=pd.DataFrame(sheet_rows),
            column_mappings=pd.DataFrame(column_rows),
            column_comparison=column_comparison,
            column_profiles=pd.DataFrame(column_profile_rows),
            privacy_records=pd.DataFrame(privacy_rows),
            source_tables=source_tables,
            issues=issues,
            summary={
                "input_path": input_label or "선택 파일",
                "file_count": len(file_paths),
                "issue_count": len(issues),
                "template": workflow_template.name if workflow_template else "없음",
                "unique_column_count": len(column_comparison) if not column_comparison.empty else 0,
                "common_column_count": common_column_count,
            },
        )

    def _column_comparison_frame(self, columns_by_file: dict[str, list[str]]) -> pd.DataFrame:
        if not columns_by_file:
            return pd.DataFrame()
        file_names = list(columns_by_file)
        all_columns = sorted({column for columns in columns_by_file.values() for column in columns}, key=normalize_header)
        rows = []
        for column in all_columns:
            present_files = [file_name for file_name in file_names if column in columns_by_file[file_name]]
            missing_files = [file_name for file_name in file_names if column not in columns_by_file[file_name]]
            row: dict[str, object] = {
                "컬럼명": column,
                "정규화 컬럼명": normalize_header(column),
                "등장 파일 수": len(present_files),
                "누락 파일": ", ".join(missing_files),
            }
            for file_name in file_names:
                row[file_name] = "있음" if file_name in present_files else ""
            rows.append(row)
        return pd.DataFrame(rows)

    def _column_profile_rows(
        self,
        file_name: str,
        sheet_name: str,
        table: pd.DataFrame,
    ) -> list[dict[str, object]]:
        rows = []
        row_count = len(table)
        for column in table.columns:
            values = table[column]
            non_empty = int(values.map(lambda value: not is_blank(value)).sum())
            blank = row_count - non_empty
            rows.append(
                {
                    "파일명": file_name,
                    "시트명": sheet_name,
                    "컬럼명": str(column),
                    "행 수": row_count,
                    "값 있음": non_empty,
                    "빈값": blank,
                    "고유값 수": int(values.dropna().map(display_value).nunique()),
                    "추정 유형": self._infer_type(values),
                }
            )
        return rows

    def _infer_type(self, values: pd.Series) -> str:
        text_values = values[values.map(lambda value: not is_blank(value))].map(display_value)
        if text_values.empty:
            return "빈값"
        numeric = pd.to_numeric(text_values.str.replace(",", "", regex=False).str.replace("원", "", regex=False), errors="coerce")
        if numeric.notna().mean() >= 0.8:
            return "숫자"
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            parsed_dates = pd.to_datetime(text_values, errors="coerce")
        if parsed_dates.notna().mean() >= 0.8:
            return "날짜"
        return "문자"


class InspectionResult:
    def __init__(
        self,
        files: pd.DataFrame,
        sheets: pd.DataFrame,
        column_mappings: pd.DataFrame,
        issues: list[ValidationIssue],
        summary: dict[str, object],
        column_comparison: pd.DataFrame | None = None,
        column_profiles: pd.DataFrame | None = None,
        privacy_records: pd.DataFrame | None = None,
        source_tables: dict[str, pd.DataFrame] | None = None,
    ) -> None:
        self.files = files
        self.sheets = sheets
        self.column_mappings = column_mappings
        self.column_comparison = column_comparison if column_comparison is not None else pd.DataFrame()
        self.column_profiles = column_profiles if column_profiles is not None else pd.DataFrame()
        self.privacy_records = privacy_records if privacy_records is not None else pd.DataFrame()
        self.source_tables = source_tables or {}
        self.issues = issues
        self.summary = summary
