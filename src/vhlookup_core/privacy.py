from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

import pandas as pd

from vhlookup_core.normalization import display_value, is_blank, normalize_header


@dataclass(frozen=True)
class PrivacyRule:
    category: str
    severity: str
    header_keywords: tuple[str, ...]
    action: str


PRIVACY_RULES = (
    PrivacyRule(
        category="개인 식별 이름",
        severity="주의",
        header_keywords=("성명", "이름", "직원명", "담당자", "담당자명", "대상자명", "신청자명"),
        action="결과 공유 전 이름 컬럼이 필요한 대상에게만 전달되는지 확인하세요.",
    ),
    PrivacyRule(
        category="연락처",
        severity="주의",
        header_keywords=("연락처", "전화번호", "휴대폰", "핸드폰", "내선번호", "phone", "mobile"),
        action="외부 공유 전 연락처 컬럼 필요 여부를 확인하세요.",
    ),
    PrivacyRule(
        category="이메일",
        severity="주의",
        header_keywords=("이메일", "메일", "email", "e-mail"),
        action="이메일 주소가 포함된 결과는 배포 범위를 제한하세요.",
    ),
    PrivacyRule(
        category="주소",
        severity="주의",
        header_keywords=("주소", "거주지", "소재지", "address"),
        action="주소 정보가 포함된 경우 비공개 저장 위치에 보관하세요.",
    ),
    PrivacyRule(
        category="계좌 정보",
        severity="높음",
        header_keywords=("계좌", "계좌번호", "은행", "예금주", "bank", "account"),
        action="계좌 정보는 결과 공유 전 마스킹 또는 별도 보관을 검토하세요.",
    ),
    PrivacyRule(
        category="고유식별정보",
        severity="높음",
        header_keywords=("주민등록", "주민번호", "외국인등록", "여권", "rrn", "passport"),
        action="고유식별정보는 최소 처리 원칙에 따라 포함 여부를 재검토하세요.",
    ),
)

PATTERN_RULES = (
    ("이메일 패턴", "주의", re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+"), "이메일 주소 패턴이 감지되었습니다."),
    ("주민등록번호 패턴", "높음", re.compile(r"\b\d{6}-?[1-4]\d{6}\b"), "주민등록번호로 보이는 패턴이 감지되었습니다."),
    ("전화번호 패턴", "주의", re.compile(r"\b(?:0\d{1,2}-?\d{3,4}-?\d{4})\b"), "전화번호로 보이는 패턴이 감지되었습니다."),
)


class PrivacyScanner:
    def scan_frame(self, frame: pd.DataFrame) -> list[dict[str, Any]]:
        records: list[dict[str, Any]] = []
        if frame.empty:
            return records
        records.extend(self._scan_headers(frame))
        records.extend(self._scan_patterns(frame))
        return dedupe_records(records)

    def _scan_headers(self, frame: pd.DataFrame) -> list[dict[str, Any]]:
        records = []
        for column in frame.columns:
            normalized = normalize_header(column)
            for rule in PRIVACY_RULES:
                keywords = {normalize_header(keyword) for keyword in rule.header_keywords}
                if any(keyword and keyword in normalized for keyword in keywords):
                    records.append(
                        {
                            "점검 유형": rule.category,
                            "위험도": rule.severity,
                            "컬럼명": str(column),
                            "발견 기준": "컬럼명",
                            "감지 건수": int(frame[column].map(lambda value: not is_blank(value)).sum()),
                            "조치 안내": rule.action,
                        }
                    )
        return records

    def _scan_patterns(self, frame: pd.DataFrame) -> list[dict[str, Any]]:
        records = []
        for column in frame.columns:
            values = frame[column].map(display_value)
            for category, severity, pattern, message in PATTERN_RULES:
                count = int(values.map(lambda value: bool(value and pattern.search(value))).sum())
                if count:
                    records.append(
                        {
                            "점검 유형": category,
                            "위험도": severity,
                            "컬럼명": str(column),
                            "발견 기준": message,
                            "감지 건수": count,
                            "조치 안내": "개별 값은 리포트에 저장하지 않았습니다. 원본 파일에서 필요 여부를 확인하세요.",
                        }
                    )
        return records


def dedupe_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen = set()
    deduped = []
    for record in records:
        key = (record.get("점검 유형"), record.get("컬럼명"), record.get("발견 기준"))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(record)
    return deduped
