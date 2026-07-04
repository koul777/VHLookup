from __future__ import annotations

import sys
from pathlib import Path

from vhlookup_core import (
    APP_DISPLAY_NAME,
    ConsolidationEngine,
    ExcelLoader,
    HeaderDetector,
    KeySpec,
    AutoLookupPlanner,
    ColumnMapper,
    MergeEngine,
    ReconciliationEngine,
    ReportWriter,
    SheetDetector,
    ValidationIssue,
    get_template,
    templates_for_mode,
)
from vhlookup_core.horizontal import HorizontalTableEngine


def _load_qt():
    try:
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import (
            QApplication,
            QComboBox,
            QFileDialog,
            QFormLayout,
            QGroupBox,
            QHBoxLayout,
            QLabel,
            QLineEdit,
            QListWidget,
            QMainWindow,
            QMessageBox,
            QPushButton,
            QPlainTextEdit,
            QStackedWidget,
            QTableWidget,
            QTableWidgetItem,
            QVBoxLayout,
            QWidget,
        )
    except ModuleNotFoundError as exc:
        raise SystemExit("PySide6 is not installed. Run: python -m pip install PySide6") from exc
    return locals()


qt = _load_qt()
Qt = qt["Qt"]
QApplication = qt["QApplication"]
QComboBox = qt["QComboBox"]
QFileDialog = qt["QFileDialog"]
QFormLayout = qt["QFormLayout"]
QGroupBox = qt["QGroupBox"]
QHBoxLayout = qt["QHBoxLayout"]
QLabel = qt["QLabel"]
QLineEdit = qt["QLineEdit"]
QListWidget = qt["QListWidget"]
QMainWindow = qt["QMainWindow"]
QMessageBox = qt["QMessageBox"]
QPushButton = qt["QPushButton"]
QPlainTextEdit = qt["QPlainTextEdit"]
QStackedWidget = qt["QStackedWidget"]
QTableWidget = qt["QTableWidget"]
QTableWidgetItem = qt["QTableWidgetItem"]
QVBoxLayout = qt["QVBoxLayout"]
QWidget = qt["QWidget"]


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_DISPLAY_NAME)
        self.resize(1180, 760)
        self.loader = ExcelLoader()
        self.header_detector = HeaderDetector()
        self.sheet_detector = SheetDetector(self.header_detector)
        self.report_writer = ReportWriter()

        self.mode_select = QComboBox()
        self.mode_select.addItems(
            [
                "학교/부서 제출자료 수합",
                "기준표에서 값 붙이기",
                "빠진 사람/누락자료 찾기",
                "월별 가로표 세로 변환",
            ]
        )
        self.mode_select.currentIndexChanged.connect(self._switch_mode)

        self.stack = QStackedWidget()
        self.consolidation_page = self._build_consolidation_page()
        self.lookup_page = self._build_lookup_page()
        self.reconciliation_page = self._build_reconciliation_page()
        self.horizontal_page = self._build_horizontal_page()
        self.stack.addWidget(self.consolidation_page)
        self.stack.addWidget(self.lookup_page)
        self.stack.addWidget(self.reconciliation_page)
        self.stack.addWidget(self.horizontal_page)

        self.preview = QTableWidget()
        self.preview.setAlternatingRowColors(True)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)

        root = QWidget()
        layout = QVBoxLayout(root)
        header = QLabel(f"{APP_DISPLAY_NAME}\n공공기관 엑셀 수합, 명단 대조, 누락 확인을 내 PC에서 끝냅니다.")
        header.setObjectName("heroLabel")
        layout.addWidget(header)
        layout.addWidget(self.mode_select)
        layout.addWidget(self.stack)
        layout.addWidget(QLabel("미리보기"))
        layout.addWidget(self.preview, stretch=1)
        layout.addWidget(QLabel("확인 필요한 항목"))
        layout.addWidget(self.log, stretch=1)
        self.setCentralWidget(root)
        self._apply_style()

    def _build_consolidation_page(self) -> QWidget:
        page = QWidget()
        layout = QFormLayout(page)
        self.consolidation_template_select = self._template_combo("consolidation")
        self.consolidation_template_select.currentIndexChanged.connect(self._apply_consolidation_template)
        self.folder_input = QLineEdit()
        folder_button = QPushButton("폴더 선택")
        folder_button.clicked.connect(self._choose_folder)
        folder_row = QHBoxLayout()
        folder_row.addWidget(self.folder_input)
        folder_row.addWidget(folder_button)

        self.consolidation_output = QLineEdit(str(Path.cwd() / "vhlookup_consolidated.xlsx"))
        self.standard_columns_input = QLineEdit()
        output_button = QPushButton("저장 위치")
        output_button.clicked.connect(lambda: self._choose_output(self.consolidation_output))
        output_row = QHBoxLayout()
        output_row.addWidget(self.consolidation_output)
        output_row.addWidget(output_button)

        run_button = QPushButton("수합 실행")
        run_button.clicked.connect(self._run_consolidation)
        layout.addRow("업무 템플릿", self.consolidation_template_select)
        layout.addRow("수합 폴더", folder_row)
        layout.addRow("표준 컬럼", self.standard_columns_input)
        layout.addRow("결과 파일", output_row)
        layout.addRow(run_button)
        self._apply_consolidation_template()
        return page

    def _build_lookup_page(self) -> QWidget:
        page = QWidget()
        layout = QFormLayout(page)
        self.lookup_template_select = self._template_combo("lookup")
        self.lookup_template_select.currentIndexChanged.connect(self._apply_lookup_template)
        self.reference_file_input = QLineEdit()
        self.target_file_input = QLineEdit()
        self.lookup_output = QLineEdit(str(Path.cwd() / "vhlookup_lookup.xlsx"))
        self.key_columns_input = QLineEdit()
        self.value_columns_input = QLineEdit()

        layout.addRow("업무 템플릿", self.lookup_template_select)
        layout.addRow("기준 파일", self._file_row(self.reference_file_input))
        layout.addRow("대상 파일", self._file_row(self.target_file_input))
        layout.addRow("키 컬럼", self.key_columns_input)
        layout.addRow("가져올 컬럼", self.value_columns_input)
        layout.addRow("결과 파일", self._output_row(self.lookup_output))
        run_button = QPushButton("값 붙이기 실행")
        run_button.clicked.connect(self._run_lookup)
        layout.addRow(run_button)
        self._apply_lookup_template()
        return page

    def _build_reconciliation_page(self) -> QWidget:
        page = QWidget()
        layout = QFormLayout(page)
        self.reconciliation_template_select = self._template_combo("reconciliation")
        self.reconciliation_template_select.currentIndexChanged.connect(self._apply_reconciliation_template)
        self.reconciliation_reference_input = QLineEdit()
        self.reconciliation_target_input = QLineEdit()
        self.reconciliation_key_input = QLineEdit()
        self.reconciliation_output = QLineEdit(str(Path.cwd() / "vhlookup_reconciliation.xlsx"))
        layout.addRow("업무 템플릿", self.reconciliation_template_select)
        layout.addRow("기준 명단", self._file_row(self.reconciliation_reference_input))
        layout.addRow("대조 명단", self._file_row(self.reconciliation_target_input))
        layout.addRow("키 컬럼", self.reconciliation_key_input)
        layout.addRow("결과 파일", self._output_row(self.reconciliation_output))
        run_button = QPushButton("누락자료 찾기 실행")
        run_button.clicked.connect(self._run_reconciliation)
        layout.addRow(run_button)
        self._apply_reconciliation_template()
        return page

    def _build_horizontal_page(self) -> QWidget:
        page = QWidget()
        layout = QFormLayout(page)
        self.horizontal_template_select = self._template_combo("horizontal")
        self.horizontal_template_select.currentIndexChanged.connect(self._apply_horizontal_template)
        self.horizontal_file_input = QLineEdit()
        self.horizontal_id_columns_input = QLineEdit()
        self.horizontal_output = QLineEdit(str(Path.cwd() / "vhlookup_horizontal.xlsx"))
        layout.addRow("업무 템플릿", self.horizontal_template_select)
        layout.addRow("가로표 파일", self._file_row(self.horizontal_file_input))
        layout.addRow("행 기준 컬럼", self.horizontal_id_columns_input)
        layout.addRow("결과 파일", self._output_row(self.horizontal_output))
        run_button = QPushButton("세로 변환 실행")
        run_button.clicked.connect(self._run_horizontal)
        layout.addRow(run_button)
        self._apply_horizontal_template()
        return page

    def _template_combo(self, mode: str) -> QComboBox:
        combo = QComboBox()
        for template in templates_for_mode(mode):
            combo.addItem(template.name, template.id)
        return combo

    def _selected_template(self, combo: QComboBox):
        template_id = combo.currentData()
        return get_template(template_id) if template_id else None

    def _apply_consolidation_template(self) -> None:
        template = self._selected_template(self.consolidation_template_select)
        if not template:
            return
        self.standard_columns_input.setText(", ".join(template.standard_columns))
        self.consolidation_output.setText(str(Path.cwd() / template.output_name))

    def _apply_lookup_template(self) -> None:
        template = self._selected_template(self.lookup_template_select)
        if not template:
            return
        self.key_columns_input.setText(", ".join(template.key_columns))
        self.value_columns_input.setText(", ".join(template.value_columns))
        self.lookup_output.setText(str(Path.cwd() / template.output_name))

    def _apply_reconciliation_template(self) -> None:
        template = self._selected_template(self.reconciliation_template_select)
        if not template:
            return
        self.reconciliation_key_input.setText(", ".join(template.key_columns))
        self.reconciliation_output.setText(str(Path.cwd() / template.output_name))

    def _apply_horizontal_template(self) -> None:
        template = self._selected_template(self.horizontal_template_select)
        if not template:
            return
        self.horizontal_id_columns_input.setText(", ".join(template.key_columns))
        self.horizontal_output.setText(str(Path.cwd() / template.output_name))

    def _file_row(self, input_widget: QLineEdit) -> QHBoxLayout:
        row = QHBoxLayout()
        button = QPushButton("파일 선택")
        button.clicked.connect(lambda: self._choose_file(input_widget))
        row.addWidget(input_widget)
        row.addWidget(button)
        return row

    def _output_row(self, input_widget: QLineEdit) -> QHBoxLayout:
        row = QHBoxLayout()
        button = QPushButton("저장 위치")
        button.clicked.connect(lambda: self._choose_output(input_widget))
        row.addWidget(input_widget)
        row.addWidget(button)
        return row

    def _switch_mode(self, index: int) -> None:
        self.stack.setCurrentIndex(index)
        self.preview.clear()
        self.log.clear()

    def _choose_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "수합할 폴더 선택")
        if folder:
            self.folder_input.setText(folder)

    def _choose_file(self, input_widget: QLineEdit) -> None:
        file, _ = QFileDialog.getOpenFileName(self, "파일 선택", "", "Excel/CSV (*.xlsx *.xlsm *.csv)")
        if file:
            input_widget.setText(file)

    def _choose_output(self, input_widget: QLineEdit) -> None:
        file, _ = QFileDialog.getSaveFileName(self, "결과 파일 저장", input_widget.text(), "Excel (*.xlsx)")
        if file:
            input_widget.setText(file if file.lower().endswith(".xlsx") else f"{file}.xlsx")

    def _run_consolidation(self) -> None:
        try:
            template = self._selected_template(self.consolidation_template_select)
            standard_columns = self._split_columns(self.standard_columns_input.text()) or None
            result = ConsolidationEngine().consolidate_folder(
                self.folder_input.text(),
                standard_columns=standard_columns,
                template=template,
            )
            self.report_writer.write_xlsx(
                result,
                self.consolidation_output.text(),
                mark_result_cells=False,
                include_privacy_scan=False,
            )
            self._show_result(result)
        except Exception as exc:
            QMessageBox.critical(self, "수합 실패", str(exc))

    def _run_lookup(self) -> None:
        try:
            reference = self._load_first_table(self.reference_file_input.text())
            target = self._load_first_table(self.target_file_input.text())
            template = self._selected_template(self.lookup_template_select)
            key_columns = tuple(self._split_columns(self.key_columns_input.text()) or (template.key_columns if template else ()))
            value_columns = tuple(self._split_columns(self.value_columns_input.text()) or (template.value_columns if template else ()))
            plan = AutoLookupPlanner().infer_lookup_plan(
                reference,
                target,
                preferred_reference_key_columns=key_columns,
                preferred_target_key_columns=template.target_key_columns if template else (),
                preferred_value_columns=value_columns,
            )
            result = MergeEngine().merge_lookup(
                reference,
                target,
                plan.key_spec,
                value_columns=list(plan.value_columns),
            )
            result.mapping_records.extend(plan.evidence_rows)
            self._append_auto_plan_warnings(result, plan.warnings)
            if template:
                result.summary["workflow"] = template.name
            result.summary["auto_key_columns"] = " + ".join(plan.key_spec.reference_key_columns)
            result.summary["auto_target_key_columns"] = " + ".join(plan.key_spec.target_columns())
            result.summary["auto_value_columns"] = ", ".join(plan.value_columns)
            self.report_writer.write_xlsx(result, self.lookup_output.text())
            self._show_result(result)
        except Exception as exc:
            QMessageBox.critical(self, "값 붙이기 실패", str(exc))

    def _run_reconciliation(self) -> None:
        try:
            reference = self._load_first_table(self.reconciliation_reference_input.text())
            target = self._load_first_table(self.reconciliation_target_input.text())
            template = self._selected_template(self.reconciliation_template_select)
            key_columns = tuple(self._split_columns(self.reconciliation_key_input.text()) or (template.key_columns if template else ()))
            plan = AutoLookupPlanner().infer_reconciliation_key_spec(
                reference,
                target,
                preferred_reference_key_columns=key_columns,
                preferred_target_key_columns=template.target_key_columns if template else (),
            )
            result = ReconciliationEngine().compare_lists(
                reference,
                target,
                plan.key_spec,
            )
            result.mapping_records.extend(plan.evidence_rows)
            self._append_auto_plan_warnings(result, plan.warnings)
            if template:
                result.summary["workflow"] = template.name
            result.summary["auto_key_columns"] = " + ".join(plan.key_spec.reference_key_columns)
            result.summary["auto_target_key_columns"] = " + ".join(plan.key_spec.target_columns())
            self.report_writer.write_xlsx(result, self.reconciliation_output.text())
            self._show_result(result)
        except Exception as exc:
            QMessageBox.critical(self, "누락자료 찾기 실패", str(exc))

    def _run_horizontal(self) -> None:
        try:
            table = self._load_first_table(self.horizontal_file_input.text())
            id_columns = self._split_columns(self.horizontal_id_columns_input.text())
            template = self._selected_template(self.horizontal_template_select)
            horizontal_engine = HorizontalTableEngine()
            if not id_columns:
                detection = horizontal_engine.detect(table)
                id_columns = [column for column in table.columns if column not in detection.value_columns]
            converted = horizontal_engine.wide_to_long(table, id_columns=id_columns)
            from vhlookup_core.models import JobResult

            result = JobResult(
                result_frame=converted,
                summary={"workflow": template.name if template else "가로표 세로 변환", "row_count": len(converted)},
            )
            self.report_writer.write_xlsx(result, self.horizontal_output.text())
            self._show_result(result)
        except Exception as exc:
            QMessageBox.critical(self, "세로 변환 실패", str(exc))

    def _load_first_table(self, path: str):
        source = self.loader.load(path)
        sheet = self.sheet_detector.select(source)
        detection = self.header_detector.detect(sheet)
        return self.header_detector.apply(sheet, detection)

    def _split_columns(self, text: str) -> list[str]:
        return [part.strip() for part in text.split(",") if part.strip()]

    def _resolve_key_columns(
        self,
        target_frame,
        reference_key_columns: list[str],
        template_target_columns: tuple[str, ...],
    ) -> tuple[str, ...]:
        if all(column in target_frame.columns for column in reference_key_columns):
            return tuple(reference_key_columns)
        if template_target_columns and all(column in target_frame.columns for column in template_target_columns):
            return tuple(template_target_columns)
        mapping = ColumnMapper().map_columns(list(target_frame.columns), reference_key_columns, threshold=0.70)
        return tuple(mapping.target_to_source.get(column, column) for column in reference_key_columns)

    def _append_auto_plan_warnings(self, result, warnings: tuple[str, ...]) -> None:
        for warning in warnings:
            result.issues.append(
                ValidationIssue(
                    issue_type="auto_inference_warning",
                    message=warning,
                    severity="info",
                )
            )

    def _show_result(self, result) -> None:
        frame = result.result_frame.head(200)
        self.preview.setRowCount(len(frame))
        self.preview.setColumnCount(len(frame.columns))
        self.preview.setHorizontalHeaderLabels([str(column) for column in frame.columns])
        for row_index, (_, row) in enumerate(frame.iterrows()):
            for column_index, value in enumerate(row.tolist()):
                item = QTableWidgetItem("" if value is None else str(value))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.preview.setItem(row_index, column_index, item)
        self.preview.resizeColumnsToContents()
        issue_lines = [
            f"[{issue.severity}] {issue.issue_type}: {issue.message}"
            for issue in result.issues[:300]
        ]
        self.log.setPlainText("\n".join(issue_lines) if issue_lines else "확인 필요한 항목이 없습니다.")

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                font-family: 'Segoe UI', 'Malgun Gothic', sans-serif;
                font-size: 13px;
            }
            QMainWindow {
                background: #f7f7f4;
            }
            QLabel#heroLabel {
                font-size: 22px;
                font-weight: 700;
                color: #19324a;
                padding: 8px 0;
            }
            QLineEdit, QComboBox, QPlainTextEdit, QTableWidget {
                background: #ffffff;
                border: 1px solid #c7c9c4;
                border-radius: 4px;
                padding: 6px;
            }
            QPushButton {
                background: #1f6f5f;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 7px 12px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #195c50;
            }
            """
        )


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
