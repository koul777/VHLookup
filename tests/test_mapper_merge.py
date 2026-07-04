import pandas as pd

from vhlookup_core.mapper import ColumnMapper
from vhlookup_core.merge import MergeEngine
from vhlookup_core.models import KeySpec


def test_column_mapper_uses_korean_synonyms():
    mapping = ColumnMapper().map_columns(
        source_headers=["직원번호", "이름", "소속"],
        target_headers=["사번", "성명", "부서"],
    )

    assert mapping.target_to_source == {"사번": "직원번호", "성명": "이름", "부서": "소속"}
    assert not mapping.unmapped_targets


def test_merge_lookup_reports_format_mismatch_and_missing_match():
    reference = pd.DataFrame(
        {
            "사번": ["123", "124"],
            "부서": ["총무", "인사"],
        }
    )
    target = pd.DataFrame(
        {
            "사번": ["123", "00124", "999"],
            "성명": ["홍길동", "김영희", "이철수"],
        }
    )

    result = MergeEngine().merge_lookup(
        reference=reference,
        target=target,
        key_spec=KeySpec(reference_key_columns=("사번",)),
        value_columns=["부서"],
    )

    assert result.result_frame.loc[0, "부서"] == "총무"
    assert pd.isna(result.result_frame.loc[1, "부서"])
    assert {issue.issue_type for issue in result.issues} >= {"format_mismatch", "match_failed"}
    missing_issue = next(issue for issue in result.issues if issue.issue_type == "match_failed")
    assert missing_issue.details["result_columns"] == ["부서"]


def test_merge_lookup_can_use_loose_numeric_keys():
    reference = pd.DataFrame(
        {
            "사번": ["00123", "00124"],
            "부서": ["총무", "인사"],
        }
    )
    target = pd.DataFrame(
        {
            "직원번호": ["123", "00124"],
            "성명": ["홍길동", "김영희"],
        }
    )

    result = MergeEngine().merge_lookup(
        reference=reference,
        target=target,
        key_spec=KeySpec(
            reference_key_columns=("사번",),
            target_key_columns=("직원번호",),
            normalization="loose_numeric",
        ),
        value_columns=["부서"],
    )

    assert result.result_frame["부서"].tolist() == ["총무", "인사"]
    assert not result.issues


def test_merge_lookup_blocks_duplicate_reference_keys():
    reference = pd.DataFrame(
        {
            "사번": ["123", "123"],
            "부서": ["총무", "인사"],
        }
    )
    target = pd.DataFrame({"사번": ["123"], "성명": ["홍길동"]})

    result = MergeEngine().merge_lookup(
        reference=reference,
        target=target,
        key_spec=KeySpec(reference_key_columns=("사번",)),
        value_columns=["부서"],
    )

    assert pd.isna(result.result_frame.loc[0, "부서"])
    assert "reference_duplicate_key" in {issue.issue_type for issue in result.issues}
    assert "reference_duplicate_key_blocked" in {issue.issue_type for issue in result.issues}
