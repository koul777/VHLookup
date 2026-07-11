import pandas as pd

from vhlookup_core.organization_order import OrganizationOrderSorter, parse_order_text


def test_parse_order_text_accepts_lines_commas_and_tabs():
    assert parse_order_text("Ready\nDone,Waiting\tReady") == ["Ready", "Done", "Waiting"]


def test_custom_order_sort_uses_selected_column_values_and_keeps_row_order_within_value():
    frame = pd.DataFrame(
        {
            "Ticket": ["A1", "A2", "A3", "A4", "A5"],
            "Status": ["Waiting", "Done", "Ready", "Waiting", "Ready"],
            "Amount": [10, 20, 30, 40, 50],
        }
    )

    result = OrganizationOrderSorter().sort_frame(frame, "Status", ["Ready", "Waiting", "Done"])

    assert result.result_frame["Ticket"].tolist() == ["A3", "A5", "A1", "A4", "A2"]
    assert result.summary["workflow"] == "사용자 지정 순서 정렬"
    assert result.summary["sort_column"] == "Status"
    assert result.summary["unmatched_value_count"] == 0


def test_custom_order_sort_places_values_not_in_order_at_the_end():
    frame = pd.DataFrame(
        {
            "Name": ["one", "two", "three", "four"],
            "Group": ["C", "A", "B", "Z"],
        }
    )

    result = OrganizationOrderSorter().sort_frame(frame, "Group", ["B", "A"])

    assert result.result_frame["Name"].tolist() == ["three", "two", "one", "four"]
    assert result.summary["unmatched_value_count"] == 2
    assert result.summary["unmatched_values"] == "C, Z"


def test_unique_values_are_pulled_from_the_loaded_data_order():
    frame = pd.DataFrame({"Department": ["Planning", "Budget", "Planning", "HR"]})

    values = OrganizationOrderSorter().unique_values(frame, "Department")

    assert values == ["Planning", "Budget", "HR"]
