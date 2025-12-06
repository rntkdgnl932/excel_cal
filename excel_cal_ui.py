# excel_cal_ui.py
# 하비 브라운 전용 부가세·할인 계산 UI + 3종 엑셀 자동 생성

import sys
import re
from pathlib import Path
from typing import List

from PyQt5 import QtWidgets, QtCore

from vat_excel_tool import (
    TradeInfo,
    LineItemInput,
    compute_items_with_vat,
    fill_quote_template,
    fill_delivery_template,
    fill_statement_template,
)


class ExcelCalWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("하비 브라운 부가세·할인 엑셀 생성기")
        self.resize(1100, 750)

        self._last_items_computed = None  # (더 이상 사용 안 함)

        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)
        main_layout = QtWidgets.QVBoxLayout(central)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        base_dir = Path(__file__).resolve().parent

        # ------------------------------------------------------------------
        # 거래처 정보
        # ------------------------------------------------------------------
        grp_info = QtWidgets.QGroupBox("거래처 정보")
        layout_info = QtWidgets.QGridLayout(grp_info)
        layout_info.setColumnStretch(1, 1)
        layout_info.setColumnStretch(3, 1)

        self.le_customer = QtWidgets.QLineEdit()
        self.le_customer.setPlaceholderText("예: 셀릭스, 전라제주시설단 등")

        self.date_supply = QtWidgets.QDateEdit()
        self.date_supply.setDisplayFormat("yyyy-MM-dd")
        self.date_supply.setCalendarPopup(True)
        self.date_supply.setDate(QtCore.QDate.currentDate())

        self.le_bizno = QtWidgets.QLineEdit("849-63-00642")
        self.le_contact = QtWidgets.QLineEdit("010-4874-8419")

        self.le_vat = QtWidgets.QLineEdit("10")

        layout_info.addWidget(QtWidgets.QLabel("거래처명"), 0, 0)
        layout_info.addWidget(self.le_customer, 0, 1, 1, 3)

        layout_info.addWidget(QtWidgets.QLabel("공급일자"), 1, 0)
        layout_info.addWidget(self.date_supply, 1, 1)

        layout_info.addWidget(QtWidgets.QLabel("사업자등록번호"), 1, 2)
        layout_info.addWidget(self.le_bizno, 1, 3)

        layout_info.addWidget(QtWidgets.QLabel("연락처"), 2, 0)
        layout_info.addWidget(self.le_contact, 2, 1)

        layout_info.addWidget(QtWidgets.QLabel("부가세율(%)"), 2, 2)
        layout_info.addWidget(self.le_vat, 2, 3)

        main_layout.addWidget(grp_info)

        # ------------------------------------------------------------------
        # ① 고객 요청 총액 기준 계산 (화면용)
        # ------------------------------------------------------------------
        grp_total = QtWidgets.QGroupBox("① 고객 요청 총액 기준 계산 (화면용)")
        layout_total = QtWidgets.QGridLayout(grp_total)

        self.le_total_amount = QtWidgets.QLineEdit()
        self.le_total_amount.setPlaceholderText("예: 200000 (부가세 포함 총액)")

        self.le_total_qty = QtWidgets.QLineEdit()
        self.le_total_qty.setPlaceholderText("예: 8 (총 수량)")

        self.le_total_vat = QtWidgets.QLineEdit()
        self.le_total_vat.setPlaceholderText("기본은 위의 부가세율 사용")

        self.btn_total_calc = QtWidgets.QPushButton("계산하기")

        layout_total.addWidget(QtWidgets.QLabel("총 금액(부가세 포함)"), 0, 0)
        layout_total.addWidget(self.le_total_amount, 0, 1)

        layout_total.addWidget(QtWidgets.QLabel("총 수량"), 0, 2)
        layout_total.addWidget(self.le_total_qty, 0, 3)

        layout_total.addWidget(QtWidgets.QLabel("부가세율(%)"), 1, 0)
        layout_total.addWidget(self.le_total_vat, 1, 1)
        layout_total.addWidget(self.btn_total_calc, 1, 3)

        self.lbl_total_result = QtWidgets.QLabel(
            "⇒ 총액/수량 기준으로 계산된 단가·공급가액·세액·합계를 여기 표시합니다."
        )
        self.lbl_total_result.setStyleSheet("color: #444;")
        layout_total.addWidget(self.lbl_total_result, 2, 0, 1, 4)

        self.btn_total_calc.clicked.connect(self.on_total_calc)

        main_layout.addWidget(grp_total)

        # ------------------------------------------------------------------
        # 엑셀 템플릿 경로
        # ------------------------------------------------------------------
        grp_tpl = QtWidgets.QGroupBox("엑셀 템플릿 파일 (기존 양식)")
        layout_tpl = QtWidgets.QGridLayout(grp_tpl)

        def make_tpl_row(row: int, label_text: str, default_name: str):
            label_widget = QtWidgets.QLabel(label_text)
            line_edit = QtWidgets.QLineEdit(str(base_dir / default_name))
            btn = QtWidgets.QPushButton("찾기...")

            def browse():
                path, _ = QtWidgets.QFileDialog.getOpenFileName(
                    self,
                    f"{label_text} 템플릿 선택",
                    str(base_dir),
                    "Excel Files (*.xlsx *.xlsm *.xltx *.xltm)",
                )
                if path:
                    line_edit.setText(path)

            btn.clicked.connect(browse)
            layout_tpl.addWidget(label_widget, row, 0)
            layout_tpl.addWidget(line_edit, row, 1)
            layout_tpl.addWidget(btn, row, 2)
            return line_edit

        # ex 폴더 기준
        self.le_tpl_quote = make_tpl_row(0, "견적서 템플릿", "ex/견적서.xlsx")
        self.le_tpl_delivery = make_tpl_row(1, "납품서 템플릿", "ex/납품서.xlsx")
        self.le_tpl_statement = make_tpl_row(2, "거래명세표 템플릿", "ex/거래명세표.xlsx")

        main_layout.addWidget(grp_tpl)

        # ------------------------------------------------------------------
        # ② 품목별 정가 & 임의 할인율 기준 계산
        # ------------------------------------------------------------------
        grp_items = QtWidgets.QGroupBox("② 품목별 정가 & 임의 할인율 기준 계산 (엑셀 내보내기용)")
        vbox_items = QtWidgets.QVBoxLayout(grp_items)

        self.table = QtWidgets.QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(
            ["품목명", "규격", "수량", "정가 단가(부가세 포함)", "할인율(%)"]
        )
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        vbox_items.addWidget(self.table)

        btn_row_layout = QtWidgets.QHBoxLayout()
        btn_add = QtWidgets.QPushButton("행 추가")
        btn_del = QtWidgets.QPushButton("선택 행 삭제")
        btn_cal_items = QtWidgets.QPushButton("품목 계산하기")
        btn_row_layout.addWidget(btn_add)
        btn_row_layout.addWidget(btn_del)
        btn_row_layout.addWidget(btn_cal_items)
        btn_row_layout.addStretch(1)
        vbox_items.addLayout(btn_row_layout)

        btn_add.clicked.connect(self.add_row)
        btn_del.clicked.connect(self.delete_selected_rows)
        btn_cal_items.clicked.connect(self.on_calc_items)

        # 품목 합계 표시
        summary_layout = QtWidgets.QHBoxLayout()
        self.lbl_sum_supply = QtWidgets.QLabel("공급가 합계: -")
        self.lbl_sum_vat = QtWidgets.QLabel("부가세 합계: -")
        self.lbl_sum_gross = QtWidgets.QLabel("합계(부가세 포함): -")
        for lbl in (self.lbl_sum_supply, self.lbl_sum_vat, self.lbl_sum_gross):
            lbl.setStyleSheet("font-weight: bold;")
            summary_layout.addWidget(lbl)
        summary_layout.addStretch(1)
        vbox_items.addLayout(summary_layout)

        main_layout.addWidget(grp_items, 1)

        # ------------------------------------------------------------------
        # 실행 버튼들
        # ------------------------------------------------------------------
        bottom_layout = QtWidgets.QHBoxLayout()
        self.btn_make_all = QtWidgets.QPushButton("3종 엑셀 모두 생성")
        self.btn_make_quote = QtWidgets.QPushButton("견적서만 생성")
        self.btn_make_delivery = QtWidgets.QPushButton("납품서만 생성")
        self.btn_make_statement = QtWidgets.QPushButton("거래명세표만 생성")
        bottom_layout.addWidget(self.btn_make_all)
        bottom_layout.addWidget(self.btn_make_quote)
        bottom_layout.addWidget(self.btn_make_delivery)
        bottom_layout.addWidget(self.btn_make_statement)
        bottom_layout.addStretch(1)

        main_layout.addLayout(bottom_layout)

        self.status = self.statusBar()

        self.btn_make_all.clicked.connect(self.on_make_all)
        self.btn_make_quote.clicked.connect(self.on_make_quote)
        self.btn_make_delivery.clicked.connect(self.on_make_delivery)
        self.btn_make_statement.clicked.connect(self.on_make_statement)

        # [수정] 강제로 0줄로 초기화 후 1줄 추가 (무조건 1줄 시작)
        self.table.setRowCount(0)
        self.add_row()

    # ------------------------------------------------------------------
    # ① 고객 요청 총액 기준 계산
    # ------------------------------------------------------------------
    def on_total_calc(self):
        try:
            total_txt = self.le_total_amount.text().replace(",", "").strip()
            qty_txt = self.le_total_qty.text().strip()

            if not total_txt or not qty_txt:
                raise ValueError("총 금액과 총 수량을 입력해 주세요.")

            total = int(total_txt)
            qty = int(qty_txt)
            if qty <= 0:
                raise ValueError("총 수량은 1 이상이어야 합니다.")

            vat_txt = self.le_total_vat.text().strip()
            if vat_txt:
                vat_rate = float(vat_txt)
            else:
                vat_rate = float(self.le_vat.text().strip() or "10")

            rate = vat_rate / 100.0

            gross_per = total / qty  # 1개당 (부가세 포함) 단가
            gross_per_rounded = round(gross_per)
            supply_per = round(gross_per / (1 + rate))
            vat_per = gross_per_rounded - supply_per

            supply_total = supply_per * qty
            vat_total = total - supply_total

            self.lbl_total_result.setText(
                f"⇒ 1개당 세전 단가: {supply_per:,}원, "
                f"1개당 부가세: {vat_per:,}원, "
                f"1개당 합계(부가세 포함): {gross_per_rounded:,}원 / "
                f"공급가액 합계: {supply_total:,}원, "
                f"부가세 합계: {vat_total:,}원, "
                f"합계(부가세 포함): {total:,}원"
            )

        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "계산 오류", str(e))

    # ------------------------------------------------------------------
    # 품목 테이블 조작
    # ------------------------------------------------------------------
    def add_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 2, QtWidgets.QTableWidgetItem("0"))
        self.table.setItem(row, 3, QtWidgets.QTableWidgetItem("0"))
        self.table.setItem(row, 4, QtWidgets.QTableWidgetItem("0"))

    def delete_selected_rows(self):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.table.removeRow(r)

    # ------------------------------------------------------------------
    # 거래정보/품목 수집 + 품목 계산하기
    # ------------------------------------------------------------------
    def collect_trade_info(self) -> TradeInfo:
        customer = self.le_customer.text().strip() or "미정"
        supply_date = self.date_supply.date().toString("yyyy-MM-dd")
        biz_no = self.le_bizno.text().strip()
        contact = self.le_contact.text().strip()
        vat_text = self.le_vat.text().strip() or "10"
        try:
            vat_rate = float(vat_text)
        except ValueError:
            raise ValueError("부가세율(%)을 숫자로 입력해 주세요.")
        return TradeInfo(
            customer_name=customer,
            supply_date=supply_date,
            biz_no=biz_no,
            contact=contact,
            vat_rate=vat_rate,
        )

    def collect_items(self) -> List[LineItemInput]:
        items: List[LineItemInput] = []
        for row in range(self.table.rowCount()):
            name_item = self.table.item(row, 0)
            spec_item = self.table.item(row, 1)
            qty_item = self.table.item(row, 2)
            unit_item = self.table.item(row, 3)
            disc_item = self.table.item(row, 4)

            name = (name_item.text().strip() if name_item else "")
            if not name:
                continue

            spec = (spec_item.text().strip() if spec_item else "")
            qty_text = (qty_item.text().strip() if qty_item else "0")
            price_text = (unit_item.text().strip() if unit_item else "0")
            disc_text = (disc_item.text().strip() if disc_item else "0")

            try:
                qty = int(qty_text)
            except ValueError:
                raise ValueError(f"{row + 1}행 수량이 숫자가 아닙니다.")
            try:
                unit_gross = int(price_text)
            except ValueError:
                raise ValueError(f"{row + 1}행 정가 단가가 숫자가 아닙니다.")
            try:
                disc = float(disc_text)
            except ValueError:
                raise ValueError(f"{row + 1}행 할인율이 숫자가 아닙니다.")

            items.append(
                LineItemInput(
                    name=name,
                    spec=spec,
                    qty=qty,
                    unit_gross=unit_gross,
                    discount_rate=disc,
                )
            )

        if not items:
            raise ValueError("최소 1개 이상의 품목을 입력해 주세요.")

        return items

    def on_calc_items(self):
        """품목 계산하기: 합계 화면 표시 (단순 조회용)"""
        try:
            info = self.collect_trade_info()
            items_input = self.collect_items()
            items_computed = compute_items_with_vat(items_input, info.vat_rate)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "입력/계산 오류", str(e))
            return

        total_supply = sum(it.supply_total for it in items_computed)
        total_vat = sum(it.vat_total for it in items_computed)
        total_gross = sum(it.gross_total for it in items_computed)

        self.lbl_sum_supply.setText(f"공급가 합계: {total_supply:,}원")
        self.lbl_sum_vat.setText(f"부가세 합계: {total_vat:,}원")
        self.lbl_sum_gross.setText(f"합계(부가세 포함): {total_gross:,}원")

    # ------------------------------------------------------------------
    # 엑셀 생성 공통 (거래처명/날짜별 폴더 저장)
    # ------------------------------------------------------------------

    def _ensure_items_computed(self):
        # [수정완료] 캐시 기능 완전 삭제. 무조건 화면대로 계산.
        info = self.collect_trade_info()
        items_input = self.collect_items()
        items_computed = compute_items_with_vat(items_input, info.vat_rate)

        # [디버깅용 팝업] 사장님 확인용: 계산된 줄 수를 보여줍니다.
        count = len(items_computed)
        QtWidgets.QMessageBox.information(self, "계산 확인", f"현재 입력된 {count}개의 줄로 엑셀을 만듭니다.")

        return info, items_computed

    def _run_export(self, do_quote: bool, do_delivery: bool, do_statement: bool):
        try:
            info, items_computed = self._ensure_items_computed()

            base_dir = Path(__file__).resolve().parent
            quote_tpl = Path(self.le_tpl_quote.text().strip())
            delivery_tpl = Path(self.le_tpl_delivery.text().strip())
            statement_tpl = Path(self.le_tpl_statement.text().strip())

            # --- 거래처명/날짜 기준 폴더 생성 ---
            safe_customer = re.sub(r'[\\/:*?"<>|]', "_", info.customer_name or "미정")
            safe_date = info.supply_date or "no_date"
            out_dir = base_dir / "out" / f"{safe_date}_{safe_customer}"
            out_dir.mkdir(parents=True, exist_ok=True)

            messages = []

            if do_quote:
                if quote_tpl.is_file():
                    out_path = out_dir / "견적서_자동생성.xlsx"
                    fill_quote_template(quote_tpl, out_path, info, items_computed)
                    messages.append(f"견적서: {out_path}")
                else:
                    messages.append(f"[경고] 견적서 템플릿 없음: {quote_tpl}")

            if do_delivery:
                if delivery_tpl.is_file():
                    out_path = out_dir / "납품서_자동생성.xlsx"
                    fill_delivery_template(delivery_tpl, out_path, info, items_computed)
                    messages.append(f"납품서: {out_path}")
                else:
                    messages.append(f"[경고] 납품서 템플릿 없음: {delivery_tpl}")

            if do_statement:
                if statement_tpl.is_file():
                    out_path = out_dir / "거래명세표_자동생성.xlsx"
                    fill_statement_template(statement_tpl, out_path, info, items_computed)
                    messages.append(f"거래명세표: {out_path}")
                else:
                    messages.append(f"[경고] 거래명세표 템플릿 없음: {statement_tpl}")

            if not messages:
                QtWidgets.QMessageBox.information(self, "완료", "생성된 파일이 없습니다.")
            else:
                self.status.showMessage(" / ".join(messages), 15000)
                QtWidgets.QMessageBox.information(
                    self,
                    "완료",
                    "다음 위치에 파일이 생성되었습니다.\n\n"
                    + "\n".join(str(m) for m in messages)
                    + "\n\n※ 기본 저장 위치: out/공급일자_거래처명 폴더",
                )

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "엑셀 생성 중 오류", str(e))

    # ------------------------------------------------------------------
    # 버튼 핸들러
    # ------------------------------------------------------------------
    def on_make_all(self):
        self._run_export(True, True, True)

    def on_make_quote(self):
        self._run_export(True, False, False)

    def on_make_delivery(self):
        self._run_export(False, True, False)

    def on_make_statement(self):
        self._run_export(False, False, True)


def main():
    app = QtWidgets.QApplication(sys.argv)
    win = ExcelCalWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()