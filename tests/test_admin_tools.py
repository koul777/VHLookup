import pandas as pd

from vhlookup_core.admin_tools import AdminWorkbookTools
from vhlookup_core.report import ReportWriter


def test_admin_clean_file_preserves_columns_and_adds_review_metadata(tmp_path):
    source = tmp_path / "dirty.xlsx"
    pd.DataFrame(
        [
            ["보고서", None, None, None],
            ["Employee_Name", "Department", "Salary", "Memo"],
            [" 홍길동 ", "IT", 100, " 확인 "],
            [None, None, None, None],
            ["김영희", "Sales", 200, None],
        ]
    ).to_excel(source, header=False, index=False)

    result = AdminWorkbookTools().clean_file(source)

    assert list(result.result_frame.columns) == [
        "원본 파일명",
        "원본 시트명",
        "원본 행 번호",
        "Employee_Name",
        "Department",
        "Salary",
        "Memo",
    ]
    assert len(result.result_frame) == 2
    assert result.result_frame.loc[0, "Employee_Name"] == "홍길동"
    assert result.result_frame.loc[0, "Memo"] == "확인"
    assert result.summary["row_count"] == 2


def test_admin_split_workbook_infers_department_column(tmp_path):
    source = tmp_path / "hr.xlsx"
    output = tmp_path / "split.xlsx"
    pd.DataFrame(
        {
            "Employee_Name": ["A", "B", "C"],
            "Department": ["IT", "Sales", "IT"],
            "Salary": [100, 200, 300],
        }
    ).to_excel(source, index=False)

    tools = AdminWorkbookTools()
    frame, _metadata = tools.load_table(source)
    assert tools.infer_split_column(frame) == "Department"

    split_result = tools.write_split_workbook(source, output)
    sheets = pd.read_excel(output, sheet_name=None)

    assert split_result.split_column == "Department"
    assert {"먼저확인", "전체", "IT", "Sales"} <= set(sheets)
    assert len(sheets["IT"]) == 2


def test_clean_file_report_can_be_written(tmp_path):
    source = tmp_path / "dirty.xlsx"
    output = tmp_path / "cleaned.xlsx"
    pd.DataFrame({"성명": ["홍길동"], "부서": ["총무"]}).to_excel(source, index=False)

    result = AdminWorkbookTools().clean_file(source)
    ReportWriter().write_xlsx(result, output)
    sheets = pd.read_excel(output, sheet_name=None)

    assert {"먼저확인", "결과", "처리요약", "자동추천근거", "개인정보점검"} <= set(sheets)


def test_pivot_summary_builds_cross_tab_and_workbook(tmp_path):
    source = tmp_path / "budget.xlsx"
    output = tmp_path / "pivot.xlsx"
    pd.DataFrame(
        {
            "부서": ["총무", "총무", "예산", "예산"],
            "월": ["1월", "2월", "1월", "2월"],
            "사업명": ["교육", "교육", "홍보", "홍보"],
            "금액": [100, 150, 200, 50],
        }
    ).to_excel(source, index=False)

    tools = AdminWorkbookTools()
    frame, _metadata = tools.prepare_pivot_source(source)
    defaults = tools.infer_pivot_defaults(frame)
    assert defaults["row_column"] == "부서"
    assert defaults["value_column"] == "금액"

    summary = tools.build_pivot_summary(frame, row_column="부서", column_column="월", value_column="금액", aggregation="합계")
    totals = summary.pivot_frame.set_index("부서")

    assert totals.loc["총무", "1월"] == 100
    assert totals.loc["총무", "2월"] == 150
    assert totals.loc["총무", "합계"] == 250
    assert totals.loc["합계", "합계"] == 500

    result = tools.write_pivot_workbook(source, output, row_column="부서", column_column="월", value_column="금액", aggregation="합계")
    sheets = pd.read_excel(output, sheet_name=None)

    assert result.row_column == "부서"
    assert {"먼저확인", "피벗요약", "상위목록", "기준설명", "원본"} <= set(sheets)
    assert sheets["피벗요약"].set_index("부서").loc["예산", "합계"] == 250


def test_pivot_summary_count_and_invalid_numeric_rows(tmp_path):
    frame = pd.DataFrame(
        {
            "기관명": ["A기관", "A기관", "B기관"],
            "상태": ["제출", "미제출", "제출"],
            "금액": ["1,000원", "오류", "50"],
        }
    )
    tools = AdminWorkbookTools()

    count_summary = tools.build_pivot_summary(frame, row_column="기관명", column_column="상태", aggregation="건수")
    count_totals = count_summary.pivot_frame.set_index("기관명")
    assert count_totals.loc["A기관", "합계"] == 2
    assert count_totals.loc["합계", "합계"] == 3

    sum_summary = tools.build_pivot_summary(frame, row_column="기관명", value_column="금액", aggregation="합계")
    sum_totals = sum_summary.pivot_frame.set_index("기관명")
    assert sum_totals.loc["A기관", "금액_합계"] == 1000
    assert len(sum_summary.invalid_rows) == 1
    assert sum_summary.invalid_rows.iloc[0]["기관명"] == "A기관"
