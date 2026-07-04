from zipfile import ZipFile

import pandas as pd
from openpyxl import load_workbook

from vhlookup_core.workbook_diff import WorkbookDiffEngine, WorkbookDiffReportWriter


def test_workbook_diff_writes_comments_on_changed_after_cells(tmp_path):
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"
    output = tmp_path / "diff.xlsx"

    pd.DataFrame(
        {
            "사번": ["001", "002"],
            "성명": ["홍길동", "김영희"],
            "금액": [100, 200],
        }
    ).to_excel(before, index=False)
    pd.DataFrame(
        {
            "사번": ["001", "002"],
            "성명": ["홍길동", "김영희"],
            "금액": [150, 200],
        }
    ).to_excel(after, index=False)

    result = WorkbookDiffEngine().compare_files(before, after)
    WorkbookDiffReportWriter().write_xlsx(result, output)

    sheets = pd.read_excel(output, sheet_name=None)
    workbook = load_workbook(output)
    memo_sheet = workbook["후파일_메모"]

    assert result.summary["changed_cell_count"] == 1
    assert workbook.sheetnames[0] == "후파일_메모"
    assert sheets["차이목록"].loc[0, "컬럼명"] == "금액"
    assert "차이행만" in sheets
    assert sheets["차이행만"].loc[0, "후 파일 행"] == 2
    assert memo_sheet["C2"].comment is not None
    assert "전 값: 100" in memo_sheet["C2"].comment.text
    assert "후 값: 150" in memo_sheet["C2"].comment.text
    with ZipFile(output) as archive:
        drawing_xml = "\n".join(
            archive.read(name).decode("utf-8", errors="ignore")
            for name in archive.namelist()
            if "commentsDrawing" in name
        )
    assert "width:420px" in drawing_xml
    assert "height:160px" in drawing_xml


def test_workbook_diff_accepts_user_corrected_column_mapping(tmp_path):
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"

    pd.DataFrame({"사번": ["001"], "지급액": [100]}).to_excel(before, index=False)
    pd.DataFrame({"직원번호": ["001"], "최종금액": [120]}).to_excel(after, index=False)

    result = WorkbookDiffEngine().compare_files(
        before,
        after,
        column_mapping_override={"사번": "직원번호", "지급액": "최종금액"},
        key_columns=("사번",),
    )

    assert result.summary["changed_cell_count"] == 1
    assert result.diff_frame.loc[0, "전 값"] == "100"
    assert result.diff_frame.loc[0, "후 값"] == "120"


def test_workbook_diff_reports_missing_rows_with_original_values(tmp_path):
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"
    output = tmp_path / "diff.xlsx"

    pd.DataFrame(
        {
            "사번": ["E001", "E002", "E003"],
            "성명": ["홍길동", "김영희", "박철수"],
            "금액": [100000, 150000, 120000],
        }
    ).to_excel(before, index=False)
    pd.DataFrame(
        {
            "사번": ["E001", "E002", "E004"],
            "성명": ["홍길동", "김영희", "최민수"],
            "금액": [100000, 170000, 130000],
        }
    ).to_excel(after, index=False)

    result = WorkbookDiffEngine().compare_files(before, after)
    WorkbookDiffReportWriter().write_xlsx(result, output)
    sheets = pd.read_excel(output, sheet_name=None)

    assert result.summary["missing_row_count"] == 1
    assert result.summary["added_row_count"] == 1
    assert "빠진행" in sheets
    assert sheets["빠진행"].loc[0, "비교 기준"] == "E003"
    assert sheets["빠진행"].loc[0, "성명"] == "박철수"
    assert "후 파일에 행 없음" in set(sheets["차이목록"]["상태"])
    assert "전 파일에 없던 행" in set(sheets["차이목록"]["상태"])


def test_workbook_diff_auto_maps_renamed_columns_by_data_overlap_and_loose_keys(tmp_path):
    before = tmp_path / "before.xlsx"
    after = tmp_path / "after.xlsx"

    pd.DataFrame(
        {
            "관리번호": ["001", "002", "003"],
            "지급액": [100, 200, 300],
            "담당부서": ["총무", "회계", "민원"],
        }
    ).to_excel(before, index=False)
    pd.DataFrame(
        {
            "접수ID": ["1", "002", "004"],
            "최종지급액": [100, 250, 400],
            "소속부서": ["총무", "회계", "복지"],
        }
    ).to_excel(after, index=False)

    result = WorkbookDiffEngine().compare_files(before, after)

    assert result.summary["compare_basis"] == "키 컬럼: 관리번호"
    assert result.summary["missing_row_count"] == 1
    assert result.summary["added_row_count"] == 1
    assert result.summary["changed_cell_count"] == 1
    assert result.diff_frame.loc[result.diff_frame["상태"] == "값 다름", "컬럼명"].tolist() == ["지급액"]
    assert result.row_frame.loc[result.row_frame["비교 기준"] == "3", "상태"].item() == "후 파일에 행 없음"
