from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from vhlookup_core.header import HeaderDetector
from vhlookup_core.keys import KeyRecommender, build_key_series
from vhlookup_core.loader import ExcelLoader
from vhlookup_core.mapper import ColumnMapper
from vhlookup_core.normalization import display_value, is_blank, normalize_header
from vhlookup_core.sheet import SheetDetector


@dataclass
class CellComment:
    row_index: int
    column_name: str
    message: str
    fill: str = "FDE68A"


@dataclass
class WorkbookDiffResult:
    after_frame: pd.DataFrame
    diff_frame: pd.DataFrame
    column_frame: pd.DataFrame
    row_frame: pd.DataFrame
    summary: dict[str, Any]
    comments: list[CellComment] = field(default_factory=list)


class WorkbookDiffEngine:
    def __init__(
        self,
        loader: ExcelLoader | None = None,
        header_detector: HeaderDetector | None = None,
        sheet_detector: SheetDetector | None = None,
        mapper: ColumnMapper | None = None,
        key_recommender: KeyRecommender | None = None,
    ) -> None:
        self.loader = loader or ExcelLoader()
        self.header_detector = header_detector or HeaderDetector()
        self.sheet_detector = sheet_detector or SheetDetector(self.header_detector)
        self.mapper = mapper or ColumnMapper()
        self.key_recommender = key_recommender or KeyRecommender(self.mapper)

    def compare_files(
        self,
        before_file: str | Path,
        after_file: str | Path,
        scan_rows: int | str = 30,
        column_mapping_override: dict[str, str] | None = None,
        key_columns: tuple[str, ...] | None = None,
    ) -> WorkbookDiffResult:
        before_path = Path(before_file)
        after_path = Path(after_file)
        before_frame, before_sheet = self._load_table(before_path, scan_rows)
        after_frame, after_sheet = self._load_table(after_path, scan_rows)

        before_to_after = self.suggest_column_mapping(before_frame, after_frame)
        if column_mapping_override:
            before_to_after.update(
                {
                    before_column: after_column
                    for before_column, after_column in column_mapping_override.items()
                    if before_column in before_frame.columns and after_column in after_frame.columns
                }
            )
        after_aligned = pd.DataFrame(index=after_frame.index)
        for before_column, after_column in before_to_after.items():
            after_aligned[before_column] = after_frame[after_column]

        common_columns = [column for column in before_frame.columns if column in after_aligned.columns]
        before_only_columns = [column for column in before_frame.columns if column not in common_columns]
        after_only_columns = [column for column in after_frame.columns if column not in set(before_to_after.values())]
        resolved_key_columns = tuple(column for column in (key_columns or ()) if column in common_columns)
        if not resolved_key_columns:
            resolved_key_columns = self._choose_key_columns(before_frame, after_aligned, common_columns)
        compare_columns = [column for column in common_columns if column not in resolved_key_columns]

        diff_rows: list[dict[str, Any]] = []
        row_rows: list[dict[str, Any]] = []
        comments: list[CellComment] = []

        if resolved_key_columns:
            self._compare_by_key(
                before_frame,
                after_frame,
                after_aligned,
                before_to_after,
                resolved_key_columns,
                compare_columns,
                diff_rows,
                row_rows,
                comments,
            )
            compare_basis = "키 컬럼: " + ", ".join(resolved_key_columns)
        else:
            self._compare_by_position(
                before_frame,
                after_frame,
                after_aligned,
                before_to_after,
                compare_columns,
                diff_rows,
                row_rows,
                comments,
            )
            compare_basis = "행 순서"

        for row_index in range(len(after_frame)):
            if row_index not in {comment.row_index for comment in comments}:
                continue

        column_rows = []
        for column in common_columns:
            column_rows.append(
                {
                    "상태": "양쪽 모두 있음",
                    "전 파일 컬럼": column,
                    "후 파일 컬럼": before_to_after.get(column, ""),
                }
            )
        for column in before_only_columns:
            column_rows.append({"상태": "후 파일에 컬럼 없음", "전 파일 컬럼": column, "후 파일 컬럼": ""})
        for column in after_only_columns:
            column_rows.append({"상태": "전 파일에 없던 컬럼", "전 파일 컬럼": "", "후 파일 컬럼": column})

        summary = {
            "workflow": "전/후 파일 검증",
            "before_file": str(before_path),
            "after_file": str(after_path),
            "before_sheet": before_sheet,
            "after_sheet": after_sheet,
            "compare_basis": compare_basis,
            "before_rows": len(before_frame),
            "after_rows": len(after_frame),
            "matched_column_count": len(common_columns),
            "changed_cell_count": sum(1 for row in diff_rows if row["상태"] == "값 다름"),
            "missing_row_count": sum(1 for row in row_rows if row["상태"] == "후 파일에 행 없음"),
            "added_row_count": sum(1 for row in row_rows if row["상태"] == "전 파일에 없던 행"),
            "before_only_column_count": len(before_only_columns),
            "after_only_column_count": len(after_only_columns),
        }
        return WorkbookDiffResult(
            after_frame=after_frame,
            diff_frame=pd.DataFrame(diff_rows),
            column_frame=pd.DataFrame(column_rows),
            row_frame=pd.DataFrame(row_rows),
            summary=summary,
            comments=comments,
        )

    def suggest_column_mapping(self, before_frame: pd.DataFrame, after_frame: pd.DataFrame) -> dict[str, str]:
        column_mapping = self.mapper.map_columns(list(after_frame.columns), list(before_frame.columns), threshold=0.72)
        return dict(column_mapping.target_to_source)

    def _load_table(self, path: Path, scan_rows: int | str) -> tuple[pd.DataFrame, str]:
        source = self.loader.load(path)
        sheet = self.sheet_detector.select(source)
        detection = self.header_detector.detect(sheet, scan_rows=scan_rows)
        return self.header_detector.apply(sheet, detection), sheet.name

    def _choose_key_columns(
        self,
        before_frame: pd.DataFrame,
        after_frame: pd.DataFrame,
        common_columns: list[str],
    ) -> tuple[str, ...]:
        for column in common_columns:
            normalized = normalize_header(column)
            if any(keyword in normalized for keyword in ("사번", "직원번호", "기관코드", "학교코드", "코드", "id", "번호")):
                if self._is_good_key(before_frame[column], after_frame[column]):
                    return (column,)
        if not common_columns:
            return ()
        candidates = self.key_recommender.recommend(
            before_frame.loc[:, common_columns],
            after_frame.loc[:, common_columns],
            limit=1,
        )
        if candidates and candidates[0].score >= 0.62:
            return candidates[0].reference_columns
        return ()

    def _is_good_key(self, before: pd.Series, after: pd.Series) -> bool:
        before_values = before.map(display_value)
        after_values = after.map(display_value)
        before_non_blank = before_values[before_values != ""]
        after_non_blank = after_values[after_values != ""]
        if before_non_blank.empty or after_non_blank.empty:
            return False
        if before_non_blank.duplicated().any() or after_non_blank.duplicated().any():
            return False
        overlap = len(set(before_non_blank).intersection(set(after_non_blank)))
        return overlap / max(min(before_non_blank.nunique(), after_non_blank.nunique()), 1) >= 0.5

    def _compare_by_key(
        self,
        before_frame: pd.DataFrame,
        after_frame: pd.DataFrame,
        after_aligned: pd.DataFrame,
        before_to_after: dict[str, str],
        key_columns: tuple[str, ...],
        compare_columns: list[str],
        diff_rows: list[dict[str, Any]],
        row_rows: list[dict[str, Any]],
        comments: list[CellComment],
    ) -> None:
        before_keys = build_key_series(before_frame, key_columns)
        after_keys = build_key_series(after_aligned, key_columns)
        before_map = {key: index for index, key in before_keys.items() if key}
        after_map = {key: index for index, key in after_keys.items() if key}
        all_keys = list(dict.fromkeys([*before_map.keys(), *after_map.keys()]))

        for key in all_keys:
            before_index = before_map.get(key)
            after_index = after_map.get(key)
            if before_index is None:
                row_rows.append({"상태": "전 파일에 없던 행", "비교 기준": key, "전 파일 행": "", "후 파일 행": int(after_index) + 2})
                comments.append(
                    CellComment(
                        row_index=int(after_index),
                        column_name=str(after_frame.columns[0]),
                        message="전 파일에 없던 행입니다.",
                        fill="BFDBFE",
                    )
                )
                continue
            if after_index is None:
                row_rows.append({"상태": "후 파일에 행 없음", "비교 기준": key, "전 파일 행": int(before_index) + 2, "후 파일 행": ""})
                diff_rows.append(
                    {
                        "상태": "후 파일에 행 없음",
                        "비교 기준": key,
                        "컬럼명": "",
                        "전 값": "",
                        "후 값": "",
                        "후 파일 셀": "",
                    }
                )
                continue
            row_rows.append(
                {"상태": "양쪽 모두 있음", "비교 기준": key, "전 파일 행": int(before_index) + 2, "후 파일 행": int(after_index) + 2}
            )
            self._compare_cells(
                before_frame,
                after_aligned,
                before_to_after,
                int(before_index),
                int(after_index),
                key,
                compare_columns,
                diff_rows,
                comments,
            )

    def _compare_by_position(
        self,
        before_frame: pd.DataFrame,
        after_frame: pd.DataFrame,
        after_aligned: pd.DataFrame,
        before_to_after: dict[str, str],
        compare_columns: list[str],
        diff_rows: list[dict[str, Any]],
        row_rows: list[dict[str, Any]],
        comments: list[CellComment],
    ) -> None:
        max_rows = max(len(before_frame), len(after_frame))
        for index in range(max_rows):
            key = f"{index + 2}행"
            if index >= len(before_frame):
                row_rows.append({"상태": "전 파일에 없던 행", "비교 기준": key, "전 파일 행": "", "후 파일 행": index + 2})
                comments.append(
                    CellComment(
                        row_index=index,
                        column_name=str(after_frame.columns[0]),
                        message="전 파일에 없던 행입니다.",
                        fill="BFDBFE",
                    )
                )
                continue
            if index >= len(after_frame):
                row_rows.append({"상태": "후 파일에 행 없음", "비교 기준": key, "전 파일 행": index + 2, "후 파일 행": ""})
                continue
            row_rows.append({"상태": "양쪽 모두 있음", "비교 기준": key, "전 파일 행": index + 2, "후 파일 행": index + 2})
            self._compare_cells(
                before_frame,
                after_aligned,
                before_to_after,
                index,
                index,
                key,
                compare_columns,
                diff_rows,
                comments,
            )

    def _compare_cells(
        self,
        before_frame: pd.DataFrame,
        after_aligned: pd.DataFrame,
        before_to_after: dict[str, str],
        before_index: int,
        after_index: int,
        key: str,
        compare_columns: list[str],
        diff_rows: list[dict[str, Any]],
        comments: list[CellComment],
    ) -> None:
        for column in compare_columns:
            before_value = before_frame.at[before_index, column]
            after_value = after_aligned.at[after_index, column]
            if self._same_value(before_value, after_value):
                continue
            after_column = before_to_after.get(column, column)
            cell_address = f"{after_column}{after_index + 2}"
            before_text = display_value(before_value)
            after_text = display_value(after_value)
            diff_rows.append(
                {
                    "상태": "값 다름",
                    "비교 기준": key,
                    "컬럼명": column,
                    "전 값": before_text,
                    "후 값": after_text,
                    "후 파일 셀": cell_address,
                }
            )
            comments.append(
                CellComment(
                    row_index=after_index,
                    column_name=after_column,
                    message=f"전 값: {before_text}\n후 값: {after_text}",
                )
            )

    def _same_value(self, before_value: object, after_value: object) -> bool:
        if is_blank(before_value) and is_blank(after_value):
            return True
        return display_value(before_value) == display_value(after_value)


class WorkbookDiffReportWriter:
    def write_xlsx(self, result: WorkbookDiffResult, path: str | Path) -> Path:
        output_path = Path(path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            self._guide_frame(result).to_excel(writer, sheet_name="먼저확인", index=False)
            result.after_frame.to_excel(writer, sheet_name="후파일_메모", index=False)
            self._frame_or_empty(result.diff_frame, ["상태", "비교 기준", "컬럼명", "전 값", "후 값", "후 파일 셀"]).to_excel(
                writer, sheet_name="차이목록", index=False
            )
            self._frame_or_empty(result.column_frame, ["상태", "전 파일 컬럼", "후 파일 컬럼"]).to_excel(
                writer, sheet_name="컬럼비교", index=False
            )
            self._frame_or_empty(result.row_frame, ["상태", "비교 기준", "전 파일 행", "후 파일 행"]).to_excel(
                writer, sheet_name="행비교", index=False
            )
            changed_rows = self._changed_rows_frame(result)
            if not changed_rows.empty:
                changed_rows.to_excel(writer, sheet_name="차이행만", index=False)
            pd.DataFrame([{"항목": key, "값": value} for key, value in result.summary.items()]).to_excel(
                writer, sheet_name="점검요약", index=False
            )
            self._apply_comments(writer.book, result)
            self._style_workbook(writer.book)
        return output_path

    def _guide_frame(self, result: WorkbookDiffResult) -> pd.DataFrame:
        return pd.DataFrame(
            [
                {
                    "확인 항목": "비교 기준",
                    "값": result.summary.get("compare_basis", ""),
                    "보는 방법": "키 컬럼이 잡히면 키 기준, 아니면 행 순서 기준으로 비교합니다.",
                },
                {
                    "확인 항목": "값이 다른 셀",
                    "값": result.summary.get("changed_cell_count", 0),
                    "보는 방법": "후파일_메모 시트에서 노란색 셀의 메모를 확인하세요.",
                },
                {
                    "확인 항목": "행/컬럼 차이",
                    "값": result.summary.get("missing_row_count", 0) + result.summary.get("added_row_count", 0),
                    "보는 방법": "행비교와 컬럼비교 시트에서 추가/누락을 확인하세요.",
                },
            ]
        )

    def _frame_or_empty(self, frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
        if frame.empty:
            return pd.DataFrame(columns=columns)
        return frame

    def _changed_rows_frame(self, result: WorkbookDiffResult) -> pd.DataFrame:
        if result.after_frame.empty:
            return pd.DataFrame()

        row_status: dict[int, set[str]] = {}
        for comment in result.comments:
            row_status.setdefault(comment.row_index, set()).add("셀 변경/추가")
        if not result.diff_frame.empty:
            for _, diff_row in result.diff_frame.iterrows():
                cell_ref = str(diff_row.get("후 파일 셀", ""))
                row_number = "".join(ch for ch in cell_ref if ch.isdigit())
                if row_number:
                    row_status.setdefault(int(row_number) - 2, set()).add(str(diff_row.get("상태", "")))
        if not row_status:
            return pd.DataFrame()

        rows = []
        for index in sorted(row_status):
            if 0 <= index < len(result.after_frame):
                row = {"차이 유형": ", ".join(sorted(row_status[index])), "후 파일 행": index + 2}
                row.update(result.after_frame.loc[index].to_dict())
                rows.append(row)
        return pd.DataFrame(rows)

    def _apply_comments(self, workbook, result: WorkbookDiffResult) -> None:
        if "후파일_메모" not in workbook.sheetnames:
            return
        sheet = workbook["후파일_메모"]
        column_positions = {cell.value: cell.column for cell in sheet[1]}
        for comment in result.comments:
            column = column_positions.get(comment.column_name)
            if not column:
                continue
            cell = sheet.cell(row=comment.row_index + 2, column=column)
            cell.comment = Comment(comment.message, "VHLookup")
            cell.fill = PatternFill("solid", fgColor=comment.fill)

    def _style_workbook(self, workbook) -> None:
        header_fill = PatternFill("solid", fgColor="1F6F5F")
        header_font = Font(color="FFFFFF", bold=True)
        for sheet in workbook.worksheets:
            sheet.freeze_panes = "A2"
            if sheet.max_row > 1 and sheet.max_column > 0:
                sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for column_cells in sheet.columns:
                max_length = 0
                column_letter = get_column_letter(column_cells[0].column)
                for cell in column_cells:
                    value = "" if cell.value is None else str(cell.value)
                    max_length = max(max_length, min(len(value), 48))
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                sheet.column_dimensions[column_letter].width = max(10, min(max_length + 2, 50))
