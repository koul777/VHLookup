from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

from vhlookup_core import (
    AutoLookupPlanner,
    ExcelLoader,
    HeaderDetector,
    ReportWriter,
    SheetDetector,
    TemplateCatalogWriter,
    get_template,
    template_catalog_frame,
)
from vhlookup_core.consolidation import ConsolidationEngine
from vhlookup_core.horizontal import HorizontalTableEngine
from vhlookup_core.inspection import InspectionEngine
from vhlookup_core.inspection_report import InspectionReportWriter
from vhlookup_core.merge import MergeEngine
from vhlookup_core.models import JobResult
from vhlookup_core.reconciliation import ReconciliationEngine


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    args.handler(args)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="vhlookup-cli",
        description="Local-only Excel/CSV consolidation, lookup, reconciliation, and horizontal-table conversion.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    templates = subparsers.add_parser("templates", help="내장 업무 템플릿 목록을 보여주거나 엑셀로 저장합니다.")
    templates.add_argument("--out", default="", help="템플릿 목록 xlsx 저장 경로. 비우면 콘솔 출력")
    templates.set_defaults(handler=run_templates)

    inspect = subparsers.add_parser("inspect", help="파일/폴더의 시트, 헤더, 컬럼 매칭을 사전 점검합니다.")
    inspect.add_argument("--path", required=True, help="점검할 파일 또는 폴더")
    inspect.add_argument("--out", required=True, help="사전점검 xlsx 경로")
    inspect.add_argument("--template", default="", help="업무 템플릿 ID. 지정하면 표준 컬럼 매칭까지 점검")
    inspect.set_defaults(handler=run_inspect)

    consolidate = subparsers.add_parser("consolidate", help="폴더 안 제출자료를 하나로 수합합니다.")
    consolidate.add_argument("--folder", required=True, help="수합할 폴더")
    consolidate.add_argument("--out", required=True, help="결과 xlsx 경로")
    consolidate.add_argument("--template", default="school_submission_consolidation", help="업무 템플릿 ID")
    consolidate.add_argument("--columns", default="", help="표준 컬럼. 쉼표로 구분하며 비우면 템플릿 기준")
    consolidate.set_defaults(handler=run_consolidate)

    lookup = subparsers.add_parser("lookup", help="기준표 값을 대상표에 붙입니다.")
    lookup.add_argument("--reference", required=True, help="기준 파일")
    lookup.add_argument("--target", required=True, help="대상 파일")
    lookup.add_argument("--out", required=True, help="결과 xlsx 경로")
    lookup.add_argument("--template", default="", help="업무 템플릿 ID")
    lookup.add_argument("--key", default="", help="기준표 키 컬럼. 쉼표로 구분. 비우면 자동 추천")
    lookup.add_argument("--target-key", default="", help="대상표 키 컬럼. 쉼표로 구분. 비우면 자동 추천")
    lookup.add_argument("--values", default="", help="가져올 기준표 컬럼. 쉼표로 구분. 비우면 자동 추천")
    lookup.set_defaults(handler=run_lookup)

    reconcile = subparsers.add_parser("reconcile", help="두 명단의 누락/추가 대상을 찾습니다.")
    reconcile.add_argument("--reference", required=True, help="기준 명단 파일")
    reconcile.add_argument("--target", required=True, help="대조 명단 파일")
    reconcile.add_argument("--out", required=True, help="결과 xlsx 경로")
    reconcile.add_argument("--template", default="", help="업무 템플릿 ID")
    reconcile.add_argument("--key", default="", help="기준표 키 컬럼. 쉼표로 구분. 비우면 자동 추천")
    reconcile.add_argument("--target-key", default="", help="대상표 키 컬럼. 쉼표로 구분. 비우면 자동 추천")
    reconcile.add_argument("--reference-label", default="기준표", help="리포트에 표시할 기준표 이름")
    reconcile.add_argument("--target-label", default="대상표", help="리포트에 표시할 대상표 이름")
    reconcile.set_defaults(handler=run_reconcile)

    horizontal = subparsers.add_parser("horizontal", help="월별/분기별 가로표를 세로형 자료로 바꿉니다.")
    horizontal.add_argument("--file", required=True, help="가로표 파일")
    horizontal.add_argument("--out", required=True, help="결과 xlsx 경로")
    horizontal.add_argument("--id-columns", default="", help="행 기준 컬럼. 쉼표로 구분. 비우면 월/분기 컬럼 외 나머지")
    horizontal.set_defaults(handler=run_horizontal)

    return parser


def run_templates(args: argparse.Namespace) -> None:
    if args.out:
        TemplateCatalogWriter().write_xlsx(args.out)
        print(Path(args.out).resolve())
        return
    frame = template_catalog_frame()
    for _, row in frame.iterrows():
        print(f"{row['템플릿 ID']} | {row['모드']} | {row['업무명']}")


def run_inspect(args: argparse.Namespace) -> None:
    template = get_template(args.template) if args.template else None
    result = InspectionEngine().inspect_path(args.path, template=template)
    InspectionReportWriter().write_xlsx(result, args.out)
    print(Path(args.out).resolve())
    print(f"files={result.summary.get('file_count', 0)} issues={len(result.issues)}")


def run_consolidate(args: argparse.Namespace) -> None:
    template = get_template(args.template) if args.template else None
    columns = split_columns(args.columns) or None
    result = ConsolidationEngine().consolidate_folder(args.folder, standard_columns=columns, template=template)
    ReportWriter().write_xlsx(result, args.out, mark_result_cells=False, include_privacy_scan=False)
    print_result(args.out, result)


def run_lookup(args: argparse.Namespace) -> None:
    reference = load_table(Path(args.reference))
    target = load_table(Path(args.target))
    template = get_template(args.template) if args.template else None
    preferred_key = tuple(split_columns(args.key) or (template.key_columns if template else ()))
    preferred_target_key = tuple(split_columns(args.target_key) or (template.target_key_columns if template else ()))
    preferred_values = tuple(split_columns(args.values) or (template.value_columns if template else ()))
    plan = AutoLookupPlanner().infer_lookup_plan(
        reference,
        target,
        preferred_reference_key_columns=preferred_key,
        preferred_target_key_columns=preferred_target_key,
        preferred_value_columns=preferred_values,
    )
    result = MergeEngine().merge_lookup(reference, target, plan.key_spec, list(plan.value_columns))
    attach_auto_summary(result, plan, template.name if template else "기준표에서 값 붙이기")
    ReportWriter().write_xlsx(result, args.out)
    print_result(args.out, result)


def run_reconcile(args: argparse.Namespace) -> None:
    reference = load_table(Path(args.reference))
    target = load_table(Path(args.target))
    template = get_template(args.template) if args.template else None
    preferred_key = tuple(split_columns(args.key) or (template.key_columns if template else ()))
    preferred_target_key = tuple(split_columns(args.target_key) or (template.target_key_columns if template else ()))
    plan = AutoLookupPlanner().infer_reconciliation_key_spec(
        reference,
        target,
        preferred_reference_key_columns=preferred_key,
        preferred_target_key_columns=preferred_target_key,
    )
    result = ReconciliationEngine().compare_lists(
        reference,
        target,
        plan.key_spec,
        reference_label=args.reference_label,
        target_label=args.target_label,
    )
    attach_auto_summary(result, plan, template.name if template else "빠진 사람/누락자료 찾기")
    ReportWriter().write_xlsx(result, args.out)
    print_result(args.out, result)


def run_horizontal(args: argparse.Namespace) -> None:
    table = load_table(Path(args.file))
    engine = HorizontalTableEngine()
    detection = engine.detect(table)
    id_columns = split_columns(args.id_columns) or [column for column in table.columns if column not in detection.value_columns]
    converted = engine.wide_to_long(table, id_columns=id_columns)
    result = JobResult(
        result_frame=converted,
        summary={
            "workflow": "월별 가로표 세로 변환",
            "row_count": len(converted),
            "auto_value_columns": ", ".join(detection.value_columns),
        },
    )
    ReportWriter().write_xlsx(result, args.out)
    print_result(args.out, result)


def load_table(path: Path):
    loader = ExcelLoader()
    header_detector = HeaderDetector()
    sheet_detector = SheetDetector(header_detector)
    source = loader.load(path)
    sheet = sheet_detector.select(source)
    detection = header_detector.detect(sheet)
    return header_detector.apply(sheet, detection)


def attach_auto_summary(result: JobResult, plan, workflow: str) -> JobResult:
    result.mapping_records.extend(plan.evidence_rows)
    result.summary["workflow"] = workflow
    result.summary["auto_key_columns"] = " + ".join(plan.key_spec.reference_key_columns)
    result.summary["auto_target_key_columns"] = " + ".join(plan.key_spec.target_columns())
    if plan.value_columns:
        result.summary["auto_value_columns"] = ", ".join(plan.value_columns)
    return result


def split_columns(value: str | Iterable[str]) -> list[str]:
    if isinstance(value, str):
        return [part.strip() for part in value.split(",") if part.strip()]
    return [str(part).strip() for part in value if str(part).strip()]


def print_result(path: str | Path, result: JobResult) -> None:
    print(Path(path).resolve())
    print(f"rows={len(result.result_frame)} issues={len(result.issues)}")


if __name__ == "__main__":
    raise SystemExit(main())
