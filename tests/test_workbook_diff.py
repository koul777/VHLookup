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
