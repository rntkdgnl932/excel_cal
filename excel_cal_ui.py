# excel_cal_ui.py
# 하비 브라운 전용 부가세·할인 계산 UI + 3종 엑셀 자동 생성

import os
import pandas as pd  # pandas 추가
from datetime import datetime # 날짜용 추가
import msoffcrypto

import io
from PyQt5.QtWidgets import QFileDialog # 파일 탐색기
from PyQt5.QtTest import QTest # 대기(wait)용

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


# 풀버젼
        # git config --global --add safe.directory C:/my_games/excel_cal
        # auto_blog
        # data_basic
        # pyinstaller --hidden-import PyQt5 --hidden-import pyserial --hidden-import OpenAI --hidden-import feedparser --hidden-import requests --hidden-import chardet --hidden-import google.generativeai --name excel_cal -i="icon.ico" --add-data="icon.ico;./" --icon="icon.ico" --paths "C:\my_games\excel_cal\.venv\Scripts\python.exe" main.py
        # 업데이트버젼
        # pyinstaller --hidden-import PyQt5 --hidden-import pyserial --hidden-import requests --hidden-import chardet --add-data="C:\\my_games\\game_folder\\data_game;./data_game" --name game_folder -i="game_folder_macro.ico" --add-data="game_folder_macro.ico;./" --icon="game_folder_macro.ico" --paths "C:\Users\1_S_3\AppData\Local\Programs\Python\Python311\Lib\site-packages\cv2" main.py


class ExcelCalWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("하비 브라운 부가세·할인 엑셀 생성기")
        self.resize(1180, 820)

        self.df_list = None

        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)

        root = QtWidgets.QVBoxLayout(central)
        # ✅ 전체 여백/간격은 기존 유지 (필요 시 여기서도 조절 가능)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(10)

        base_dir = Path(__file__).resolve().parent

        splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, 1)

        # --------------------------
        # 상단: 스크롤(작게)
        # --------------------------
        top_scroll = QtWidgets.QScrollArea()
        top_scroll.setWidgetResizable(True)
        top_scroll.setFrameShape(QtWidgets.QFrame.NoFrame)

        top_host = QtWidgets.QWidget()
        top_lay = QtWidgets.QVBoxLayout(top_host)

        # ✅ 여기 spacing을 늘려서 카드(거래처정보/총액/템플릿) 사이가 숨쉬게 함
        top_lay.setContentsMargins(0, 0, 0, 0)
        top_lay.setSpacing(12)

        top_scroll.setWidget(top_host)
        splitter.addWidget(top_scroll)

        def _mk_label(text: str, w: int = 110) -> QtWidgets.QLabel:
            lb = QtWidgets.QLabel(text)
            lb.setFixedWidth(w)
            lb.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            return lb

        # --------------------------
        # 거래처 정보 (여백/간격 개선)
        # --------------------------
        grp_info = QtWidgets.QGroupBox("거래처 정보")
        grp_info.setProperty("card", True)



        layout_info = QtWidgets.QGridLayout(grp_info)

        # ✅ 핵심: 안쪽 여백/행 간격을 확실히 늘림
        layout_info.setContentsMargins(14, 14, 14, 14)
        layout_info.setHorizontalSpacing(14)
        layout_info.setVerticalSpacing(12)

        layout_info.setColumnStretch(1, 1)
        layout_info.setColumnStretch(3, 1)

        # ✅ 행 자체의 최소 높이를 확보 (LineEdit 높이만 늘려도 라벨/행이 빽빽하면 답답해 보임)
        layout_info.setRowMinimumHeight(0, 34)
        layout_info.setRowMinimumHeight(1, 34)
        layout_info.setRowMinimumHeight(2, 34)

        self.le_customer = QtWidgets.QLineEdit()
        self.le_customer.setPlaceholderText("예: 셀릭스, 전라제주시설단 등")
        self.le_customer.setMinimumHeight(32)

        self.date_supply = QtWidgets.QDateEdit()
        self.date_supply.setDisplayFormat("yyyy-MM-dd")
        self.date_supply.setCalendarPopup(True)
        self.date_supply.setDate(QtCore.QDate.currentDate())
        self.date_supply.setMinimumHeight(32)

        self.le_bizno = QtWidgets.QLineEdit("849-63-00642")
        self.le_bizno.setMinimumHeight(32)

        self.le_contact = QtWidgets.QLineEdit("010-4874-8419")
        self.le_contact.setMinimumHeight(32)

        self.le_vat = QtWidgets.QLineEdit("10")
        self.le_vat.setMinimumHeight(32)

        layout_info.addWidget(_mk_label("거래처명"), 0, 0)
        layout_info.addWidget(self.le_customer, 0, 1, 1, 3)

        layout_info.addWidget(_mk_label("공급일자"), 1, 0)
        layout_info.addWidget(self.date_supply, 1, 1)

        layout_info.addWidget(_mk_label("사업자번호"), 1, 2)
        layout_info.addWidget(self.le_bizno, 1, 3)

        layout_info.addWidget(_mk_label("연락처"), 2, 0)
        layout_info.addWidget(self.le_contact, 2, 1)

        layout_info.addWidget(_mk_label("부가세율(%)"), 2, 2)
        layout_info.addWidget(self.le_vat, 2, 3)

        top_lay.addWidget(grp_info)

        # --------------------------
        # 총액 기준 계산 (기존 유지)
        # --------------------------
        grp_total = QtWidgets.QGroupBox("① 고객 요청 총액 기준 계산 (화면용)")
        grp_total.setProperty("card", True)

        layout_total = QtWidgets.QGridLayout(grp_total)
        layout_total.setContentsMargins(10, 10, 10, 10)
        layout_total.setHorizontalSpacing(10)
        layout_total.setVerticalSpacing(8)
        layout_total.setColumnStretch(1, 1)
        layout_total.setColumnStretch(3, 1)

        self.le_total_amount = QtWidgets.QLineEdit()
        self.le_total_amount.setPlaceholderText("예: 200000 (부가세 포함 총액)")
        self.le_total_amount.setMinimumHeight(30)

        self.le_total_qty = QtWidgets.QLineEdit()
        self.le_total_qty.setPlaceholderText("예: 8 (총 수량)")
        self.le_total_qty.setMinimumHeight(30)

        self.le_total_vat = QtWidgets.QLineEdit()
        self.le_total_vat.setPlaceholderText("기본은 위의 부가세율 사용")
        self.le_total_vat.setMinimumHeight(30)

        self.btn_total_calc = QtWidgets.QPushButton("계산하기")
        self.btn_total_calc.setProperty("accent", True)
        self.btn_total_calc.setMinimumHeight(32)

        layout_total.addWidget(_mk_label("총 금액"), 0, 0)
        layout_total.addWidget(self.le_total_amount, 0, 1)

        layout_total.addWidget(_mk_label("총 수량"), 0, 2)
        layout_total.addWidget(self.le_total_qty, 0, 3)

        layout_total.addWidget(_mk_label("부가세율(%)"), 1, 0)
        layout_total.addWidget(self.le_total_vat, 1, 1)
        layout_total.addWidget(self.btn_total_calc, 1, 3)

        self.lbl_total_result = QtWidgets.QLabel("⇒ 계산 결과가 여기에 표시됩니다.")
        self.lbl_total_result.setWordWrap(True)
        self.lbl_total_result.setProperty("muted", True)
        layout_total.addWidget(self.lbl_total_result, 2, 0, 1, 4)

        self.btn_total_calc.clicked.connect(self.on_total_calc)

        top_lay.addWidget(grp_total)

        # --------------------------
        # 템플릿 (기존 유지)
        # --------------------------
        grp_tpl = QtWidgets.QGroupBox("엑셀 템플릿 파일 (기존 양식)")
        grp_tpl.setProperty("card", True)

        # ✅ 겹침 원인 제거: 높이 강제 금지 (3줄 입력칸 + 마진/타이틀이 140px에 못 들어감)
        grp_tpl.setMaximumHeight(16777215)  # 사실상 무제한
        grp_tpl.setMinimumHeight(0)

        layout_tpl = QtWidgets.QGridLayout(grp_tpl)
        layout_tpl.setContentsMargins(12, 16, 12, 12)  # 타이틀 아래 공간 확보(겹침 방지)
        layout_tpl.setHorizontalSpacing(12)
        layout_tpl.setVerticalSpacing(10)
        layout_tpl.setColumnStretch(1, 1)
        layout_tpl.setColumnMinimumWidth(0, 110)  # 라벨 폭 고정 (줄바꿈/찌그러짐 방지)
        layout_tpl.setColumnMinimumWidth(2, 88)  # '찾기' 버튼 칼럼 폭 확보

        self.TEMPLATE_DIR = base_dir / "ex"

        def make_tpl_row(row: int, label_text: str, filename: str):
            lb = _mk_label(label_text, 110)
            abs_path = self.TEMPLATE_DIR / filename
            line_edit = QtWidgets.QLineEdit(str(abs_path))
            line_edit.setMinimumHeight(30)

            btn = QtWidgets.QPushButton("찾기...")
            btn.setProperty("ghost", True)
            btn.setMinimumHeight(30)

            def browse():
                path, _ = QtWidgets.QFileDialog.getOpenFileName(
                    self,
                    f"{label_text} 템플릿 선택",
                    str(self.TEMPLATE_DIR),
                    "Excel Files (*.xlsx *.xlsm *.xltx *.xltm)",
                )
                if path:
                    line_edit.setText(path)

            btn.clicked.connect(browse)

            layout_tpl.addWidget(lb, row, 0)
            layout_tpl.addWidget(line_edit, row, 1)
            layout_tpl.addWidget(btn, row, 2)
            return line_edit

        self.le_tpl_quote = make_tpl_row(0, "견적서", "견적서.xlsx")
        self.le_tpl_delivery = make_tpl_row(1, "납품서", "납품서.xlsx")
        self.le_tpl_statement = make_tpl_row(2, "거래명세표", "거래명세표.xlsx")

        top_lay.addWidget(grp_tpl)
        top_lay.addStretch(1)

        # --------------------------
        # 하단: 품목 테이블 크게 (기존 유지)
        # --------------------------
        grp_items = QtWidgets.QGroupBox("② 품목별 정가 & 임의 할인율 기준 계산 (엑셀 내보내기용)")
        grp_items.setProperty("card", True)
        vbox_items = QtWidgets.QVBoxLayout(grp_items)
        vbox_items.setContentsMargins(10, 10, 10, 10)
        vbox_items.setSpacing(8)

        self.table = QtWidgets.QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(
            ["품목명", "규격", "수량", "정가 단가(부가세 포함)", "할인율(%)"]
        )
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.EditKeyPressed
        )
        self.table.setMinimumHeight(360)

        header = self.table.horizontalHeader()
        header.setDefaultAlignment(QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft)

        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Interactive)  # 품목명
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Interactive)  # 규격
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)  # 수량
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.Interactive)  # 정가
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeToContents)  # 할인율

        self.table.setColumnWidth(0, 220)
        self.table.setColumnWidth(1, 170)
        self.table.setColumnWidth(2, 80)
        self.table.setColumnWidth(3, 240)
        self.table.setColumnWidth(4, 110)

        self.table.setRowCount(0)
        vbox_items.addWidget(self.table, 1)

        btn_row_layout = QtWidgets.QHBoxLayout()
        btn_add = QtWidgets.QPushButton("행 추가")
        btn_del = QtWidgets.QPushButton("선택 행 삭제")
        btn_cal_items = QtWidgets.QPushButton("품목 계산하기")
        btn_cal_items.setProperty("accent", True)

        btn_row_layout.addWidget(btn_add)
        btn_row_layout.addWidget(btn_del)
        btn_row_layout.addWidget(btn_cal_items)
        btn_row_layout.addStretch(1)
        vbox_items.addLayout(btn_row_layout)

        btn_add.clicked.connect(self.add_row)
        btn_del.clicked.connect(self.delete_selected_rows)
        btn_cal_items.clicked.connect(self.on_calc_items)

        summary_layout = QtWidgets.QHBoxLayout()
        self.lbl_sum_supply = QtWidgets.QLabel("공급가 합계: -")
        self.lbl_sum_vat = QtWidgets.QLabel("부가세 합계: -")
        self.lbl_sum_gross = QtWidgets.QLabel("합계(부가세 포함): -")
        for lbl in (self.lbl_sum_supply, self.lbl_sum_vat, self.lbl_sum_gross):
            lbl.setProperty("sum", True)
            summary_layout.addWidget(lbl)
        summary_layout.addStretch(1)
        vbox_items.addLayout(summary_layout)

        splitter.addWidget(grp_items)

        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([240, 560])

        # --------------------------
        # 맨 아래 실행 버튼 줄 (기존 유지)
        # --------------------------
        bottom_layout = QtWidgets.QHBoxLayout()
        self.btn_naver = QtWidgets.QPushButton("네이버 송장")
        self.btn_coopang = QtWidgets.QPushButton("쿠팡 송장")
        self.btn_make_all = QtWidgets.QPushButton("3종 엑셀 모두 생성")
        self.btn_make_quote = QtWidgets.QPushButton("견적서만 생성")
        self.btn_make_delivery = QtWidgets.QPushButton("납품서만 생성")
        self.btn_make_statement = QtWidgets.QPushButton("거래명세표만 생성")
        self.btn_make_all.setProperty("accent", True)

        bottom_layout.addWidget(self.btn_naver)
        bottom_layout.addWidget(self.btn_coopang)
        bottom_layout.addWidget(self.btn_make_all)
        bottom_layout.addWidget(self.btn_make_quote)
        bottom_layout.addWidget(self.btn_make_delivery)
        bottom_layout.addWidget(self.btn_make_statement)
        bottom_layout.addStretch(1)
        root.addLayout(bottom_layout)

        self.status = self.statusBar()

        self.btn_naver.clicked.connect(self.my_naver)
        self.btn_coopang.clicked.connect(self.my_coopang)
        self.btn_make_all.clicked.connect(self.on_make_all)
        self.btn_make_quote.clicked.connect(self.on_make_quote)
        self.btn_make_delivery.clicked.connect(self.on_make_delivery)
        self.btn_make_statement.clicked.connect(self.on_make_statement)

        # 시작은 1줄
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

            # 템플릿 경로는 UI에 적힌 값(기본은 C:\my_games\excel_cal\ex\... )
            quote_tpl = Path(self.le_tpl_quote.text().strip())
            delivery_tpl = Path(self.le_tpl_delivery.text().strip())
            statement_tpl = Path(self.le_tpl_statement.text().strip())

            # --- [변경] 결과 저장 루트: C:\my_games\excel_cal\out ---
            out_root = Path(r"C:\my_games\excel_cal\out")
            out_root.mkdir(parents=True, exist_ok=True)

            # 날짜(yyyy-MM-dd) + 거래처명으로 서브 폴더 생성
            safe_customer = re.sub(r'[\\/:*?"<>|]', "_", info.customer_name or "미정")
            date_str = info.supply_date or datetime.now().strftime("%Y-%m-%d")

            # 예: C:\my_games\excel_cal\out\2025-12-08_하비브라운
            folder_name = f"{date_str}_{safe_customer}"
            out_dir = out_root / folder_name
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
                    + "\n\n※ 기본 루트: C:\\my_games\\excel_cal\\out\\날짜_거래처명",
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

    # 네이버, 쿠팡

    def my_naver(self):


        try:
            # x = Test_check(self)
            # # self.mytestin.setText("GootEvening")
            # # self.mytestin.setDisabled(True)
            # x.start()

            a = '수취인명'
            aa = "받으시는 분"

            b = '수취인연락처1'
            bb = '받으시는 분 전화'
            c = '수취인연락처2'
            cc = '받는분핸드폰'
            d = '상품명'
            dd = '품목명'
            e = '수량'
            ee = '수량'
            f = '배송메세지'
            ff = '특기사항'
            g = '옵션정보'
            gg = '메모1'
            h = '기본배송지'
            hh = '기본배송지'
            i = '상세배송지'
            ii = '상세배송지'
            j = '우편번호'
            jj = '받는분우편번호'
            k = '구매자명'
            kk = '구매자명'
            l = '구매자연락처'
            ll = '구매자연락처'
            m = '주문번호'
            mm = '출고번호'
            n = '상품주문번호'
            nn = '상품주문번호'
            buy_time = '주문일시'

            ###################
            o = ''
            oo = '운임Type'
            p = ''
            pp = '지불조건'
            ###################

            q = '1년 주문건수'
            qq = '1년 주문건수'

            ##################

            r = '배송방법'
            s = '택배사'
            t = '송장번호'

            ############ 추가

            gagin = "각인"

            work_status = "작업유무"  # 컬럼 이름 정의

            ##############엑셀
            file_path, ext = QFileDialog.getOpenFileName(self, '파일 열기', os.getcwd(), 'excel file (*.xls *.xlsx)')
            if file_path:

                print("file_path", file_path)

                self.df_list = self.get_df_from_password_excel(file_path, "1111")

                print("결과???", self.df_list)

                # self.df_list = self.loadData(file_path)[0]

                print(self.df_list)

                # self.df_list.insert(2, c, "s")
                data_o = "s"
                self.df_list[str(oo)] = data_o
                data_p = "신용"
                self.df_list[str(pp)] = data_p

                data_s = "한진택배"
                self.df_list[str(s)] = data_s

                print(self.df_list)

                self.df_list.rename(
                    columns={a: aa, b: bb, c: cc, d: dd, f: ff, q: qq, g: gg, j: jj, k: kk, l: ll, m: mm, n: nn},
                    inplace=True)
                print("재정비")

                self.df_list[gagin] = ""

                print(self.df_list)

                print("원하는 열 나열하기1")
                columns_list = self.df_list[[aa, bb, cc, dd, ee, ff, qq, gg, hh, ii, jj, kk, ll, mm, nn, oo, pp, gagin,
                                             buy_time]].values.tolist()

                print("원하는 열 나열하기2")
                print(columns_list)
                #

                # 출고번호(m)에서 동일한거 묶고

                # 메모1(g)과 수량(e)를 품목명(d)에 다시 수정한다.

                print("특정 열 나열하기")
                columns_list_m = self.df_list[mm].values.tolist()
                print(columns_list_m)
                print("[[특정 열 중복제거]]")
                result_set_1 = set(columns_list_m)
                print("특정 열 중복제거 준비", result_set_1)
                result_set_1_len = len(result_set_1)

                result_set_2 = list(result_set_1)
                print("특정 열 중복제거 완료", result_set_2)
                result_set_2_len = len(result_set_2)
                #

                ############
                # # data_list의 데이터를 데이터프레임으로 변환
                # data = [row.split(":") for row in new_list]
                # df = pd.DataFrame(data, columns=["사용자", "게임", "서버", "다이아", "골드", "컴퓨터번호"])
                #############

                # 엑셀 저장 날짜 및 시간을 엑셀 파일명으로...
                nowDay_ = datetime.today().strftime("%Y%m%d%H%M%S")
                year = datetime.today().strftime("%Y")
                month = datetime.today().strftime("%m")
                day = datetime.today().strftime("%d")
                hour = datetime.today().strftime("%H")
                minute = datetime.today().strftime("%M")
                last = str(day) + "d_" + str(hour) + "h" + str(minute) + "m"
                nowDay = str(nowDay_)

                ################ 네이버 송장 발부 #############
                df = pd.DataFrame(columns_list,
                                  columns=[aa, bb, cc, dd, ee, ff, qq, gg, hh, ii, jj, kk, ll, mm, nn, oo, pp, gagin,
                                           buy_time])
                dir_path = "C:/my_games/excel_result/" + str(year) + "/" + str(month) + "/"
                if not os.path.isdir(dir_path):
                    os.makedirs(dir_path)

                temporary_data = pd.DataFrame(
                    columns=[aa, bb, cc, dd, ee, ff, qq, gg, hh, ii, jj, kk, ll, mm, nn, oo, pp])

                # 네이버 송장발부용
                new_data = pd.DataFrame(
                    columns=[aa, bb, cc, dd, ee, ff, qq, gg, hh, ii, jj, kk, ll, mm, nn, oo, pp, gagin, buy_time])

                # dd 품목명, ee 수량, gg 메모1
                print("new_data 1", new_data)
                for set_2 in range(result_set_2_len):

                    # 출고번호로 검색한 것...
                    result = df[df[mm] == result_set_2[set_2]]
                    print("???", result)
                    print("???", result.index)
                    # print("???", result.index[0])
                    print("???gg", result[gg])
                    # print("???gg[0]", result[gg][result.index[0]])
                    # print("???", len(result.index))
                    QTest.qWait(100)
                    # 검색된 것을 임시로 저장하기

                    print("뽑기전 temporary_data", temporary_data)

                    if len(temporary_data) == 0:
                        print("전체저장")
                        temporary_data = pd.concat([temporary_data, result])
                    else:
                        print("삭제 후 저장")
                        temporary_data = temporary_data.drop(temporary_data.index[:len(temporary_data)])
                        temporary_data = pd.concat([temporary_data, result])
                    print("뽑아낸 temporary_data", temporary_data)

                    total_ = 0

                    # 뽑아낸 주문번호의 중복갯수로 다시 반복문 돌려서 작업하기
                    for set_index in range(len(result.index)):
                        print("여기인가 1")
                        ###############

                        # 해당문구 추출하여 갯수파악하고 바로 수량 적어놓기
                        # 해당문구가 품목명에 있을 경우 pass 하기기
                        # df['col_name'].str.contains("여기에문구...")

                        ################
                        # 'gg' 문구 갯수 파악
                        # many = temporary_data[temporary_data[gg] == temporary_data[gg][result.index[set_index]]]
                        # print("how many???", len(many.index))

                        # print("set_index", set_index)
                        # print("set_index?", result.index[set_index])
                        add_write = "=> " + str(result[ee][result.index[set_index]]) + " ea"

                        write_num = set_index + 1
                        result_write_num = str(write_num) + ". "

                        if set_index == 0:
                            # print("여기인가 2")
                            # print("여기인가 2", df.loc[result.index[set_index], gg])
                            # print("여기인가 2", str(add_write))
                            # test = df.loc[result.index[set_index], gg] + str(add_write)
                            # print("add_write", add_write)
                            # print("test", test)

                            # print("df.loc[result.index[set_index], gg]", df.loc[result.index[set_index], gg])
                            if pd.isnull(df.loc[result.index[set_index], gg]) == True:
                                # if len(df.loc[result.index[set_index], gg]) == 0:
                                df.loc[result.index[set_index], dd] = result_write_num + df.loc[
                                    result.index[set_index], dd] + str(add_write)

                                df.loc[result.index[set_index], gagin] = result_write_num + df.loc[
                                    result.index[set_index], dd] + str(add_write)
                            else:

                                write_memo = str(df.loc[result.index[set_index], gg]).replace("여기에 문구:", "")
                                write_memo = str(write_memo).replace("여기에 각인 문구:", "")
                                df.loc[result.index[set_index], dd] = result_write_num + write_memo + str(add_write)

                                df.loc[result.index[set_index], gagin] = result_write_num + write_memo + str(add_write)

                            df_2 = df.iloc[result.index[0]]
                            print("set_index == 0", df_2)
                            new_data.loc[len(new_data)] = df_2

                            # new_data.loc[set_2, dd] = new_data.loc[set_2, gg] + "\n" + df.loc[result.index[set_index], gg]

                        else:

                            print("수정 및 추가하기1", new_data)
                            print("수정 및 추가하기2", result.index[set_index])

                            # [송장발부]마지막으로 같은 문구 갯수 파악하기
                            result_memo = new_data[new_data[dd] == result.loc[result.index[set_index], gg]] + add_write

                            print("수정 및 추가하기3", result_memo)
                            print("수정 및 추가하기4", df.loc[result.index[set_index], gg])

                            if pd.isnull(df.loc[result.index[set_index], gg]) == True:
                                print("this is nan")

                            if len(result_memo.index) == 0:
                                print("없당.", result_memo[gg])
                                if pd.isnull(df.loc[result.index[set_index], gg]) == True:
                                    new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + result_write_num + \
                                                              df.loc[result.index[set_index], dd] + add_write
                                    new_data.loc[set_2, gagin] = new_data.loc[set_2, gagin] + "\n" + result_write_num + \
                                                                 df.loc[result.index[set_index], dd] + add_write
                                else:

                                    write_memo = str(df.loc[result.index[set_index], gg]).replace("여기에 문구:", "")
                                    write_memo = str(write_memo).replace("여기에 각인 문구:", "")

                                    new_data.loc[set_2, dd] = new_data.loc[
                                                                  set_2, dd] + "\n" + result_write_num + write_memo + add_write
                                    new_data.loc[set_2, gagin] = new_data.loc[
                                                                     set_2, gagin] + "\n" + result_write_num + write_memo + add_write
                            else:
                                print("송장발부는 이미 있다..", result_memo[gg])

                            # # 마지막으로 같은 문구 갯수 파악하기
                            # result_memo = df[df[gg] == df.loc[result.index[set_index], gg]]
                            #
                            # if len(result_memo.index) == 1:
                            #     new_data.loc[set_2, dd] = new_data.loc[set_2, gg] + "\n" + df.loc[result.index[set_index], gg] + "=> " + str(len(result_memo.index)) +"개"
                            # else:
                            #     print("g")
                            # new_data.loc[set_2, gg] = new_data.loc[set_2, gg] + "\n" + df.loc[result.index[set_index], gg]

                        # print("new_data for", new_data)
                        # print("new_data for2", new_data[dd])

                        # total 갯수 구하기 및 1년 주문건수

                        title_display_count = 11
                        print("total_total_total_total_total_total_", result[ee][result.index[set_index]])
                        total_ += int(result[ee][result.index[set_index]])
                        print("total_", total_)
                        if len(result.index) - 1 == set_index:

                            # 몇개 표시할지...

                            # new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + "total => " + str(total_) + " ea"
                            new_data.loc[set_2, gagin] = "[hobby brown] total => " + str(total_) + " ea" + "\n\n" + \
                                                         new_data.loc[set_2, dd]

                            # 1년 주문건수
                            # if new_data.loc[set_2, qq] > 0:
                            new_data.loc[set_2, gagin] = "1년 주문건수 : " + str(new_data.loc[set_2, qq]) + "건" + "\n" + \
                                                         new_data.loc[set_2, gagin]

                            if '\n' in new_data.loc[set_2, dd]:

                                result_split = new_data.loc[set_2, dd].split('\n')
                                print("new_data.loc[set_2, dd].split('\n')", result_split)

                                if len(result_split) < title_display_count + 1:
                                    new_data.loc[set_2, dd] = new_data.loc[set_2, dd]
                                    new_data.loc[set_2, dd] = "[hobby brown] total => " + str(total_) + " ea" + "\n\n" + \
                                                              new_data.loc[
                                                                  set_2, dd]
                                else:

                                    for w in range(title_display_count):
                                        if w == 0:

                                            new_data.loc[set_2, dd] = result_split[w]
                                        elif w < title_display_count - 3:
                                            new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + result_split[w]
                                        elif w == title_display_count - 1:
                                            new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + str("^_~")
                                        else:
                                            new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + str(".")
                                    # new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + "total => " + str(total_) + " ea"
                                    new_data.loc[set_2, dd] = "[hobby brown] total => " + str(total_) + " ea" + "\n\n" + \
                                                              new_data.loc[set_2, dd]
                            else:
                                # new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + "total => " + str(total_) + " ea"
                                new_data.loc[set_2, dd] = "[hobby brown] total => " + str(total_) + " ea" + "\n\n" + \
                                                          new_data.loc[set_2, dd]
                                # print("new_data 2", new_data)

                # print("new_data 3", len(new_data))

                # 수량 1로 바꾸기
                for many in range(len(new_data)):
                    new_data.loc[many, ee] = 1
                # 메모1 빼고 '1년 구매건수'로 바꾸기
                new_data = pd.DataFrame(new_data,
                                        columns=[aa, bb, cc, dd, ee, ff, qq, hh, ii, jj, kk, ll, mm, nn, oo, pp, gagin,
                                                 buy_time])

                #################################################
                #################################################
                #################################################
                print("발송처리 부분")

                # data_s = "한진택배"
                # self.df_list[str(s)] = data_s

                # get_name = "받으시는 분"
                # billing_number = "운송장번호"
                # df_send = pd.DataFrame(columns_list, columns=[nn, r, s, t, a, "비고", mm, get_name, billing_number])
                self.df_list.rename(
                    columns={aa: a},
                    inplace=True)
                print("재정비")
                print(self.df_list)

                print("발송처리 : 원하는 열 나열하기1")
                columns_list = self.df_list[
                    [nn, r, s, t, a, mm]].values.tolist()
                print("발송처리 : 원하는 열 나열하기2")
                print(columns_list)
                #

                # 출고번호(m)에서 동일한거 묶고

                # 메모1(g)과 수량(e)를 품목명(d)에 다시 수정한다.

                print("특정 열 나열하기")
                columns_list_m = self.df_list[mm].values.tolist()
                print(columns_list_m)
                print("[[특정 열 중복제거]]")
                result_set_1 = set(columns_list_m)
                print("특정 열 중복제거 준비", result_set_1)
                result_set_1_len = len(result_set_1)

                result_set_2 = list(result_set_1)
                print("특정 열 중복제거 완료", result_set_2)
                result_set_2_len = len(result_set_2)

                ################ 발송 처리 #############

                df_send = pd.DataFrame(columns_list, columns=[nn, r, s, t, a, mm])

                # 발송처리용

                get_name = "출고번호넣기"
                billing_number = "운송장번호"

                send_data = pd.DataFrame(columns=[nn, r, s, t, a, "비고", mm, get_name, billing_number])

                # dd 품목명, ee 수량, gg 메모1
                for set_2 in range(result_set_2_len):

                    # 출고번호로 검색한 것...
                    result = df_send[df_send[mm] == result_set_2[set_2]]
                    QTest.qWait(100)
                    # 검색된 것을 임시로 저장하기

                    # 뽑아낸 주문번호의 중복갯수로 다시 반복문 돌려서 작업하기
                    for set_index in range(len(result.index)):
                        df_send_result = df_send.iloc[result.index[set_index]]
                        send_data.loc[len(send_data)] = df_send_result

                ###################################################
                ###################################################
                ##################################################
                new_data.astype(str)
                send_data.astype(str)

                new_data[work_status] = ""

                # 데이터프레임을 엑셀 파일로 저장
                excel_file_name = dir_path + last + "네이버_송장발부.xlsx"
                # df.to_excel(excel_file_name, index=False)
                # writer_1 = pd.ExcelWriter(excel_file_name, options={'strings_to_urls': False}
                new_data[dd] = new_data[dd].astype(str)
                new_data[mm] = new_data[mm].astype(str)
                new_data[nn] = new_data[nn].astype(str)
                new_data[buy_time] = new_data[buy_time].astype(str)
                new_data.to_excel(excel_file_name, index=False, engine="openpyxl")
                # writer_1.save()

                # === 추가: 품목 전체 표기 버전 엑셀도 함께 저장 ===
                # new_data에는 dd(품목명)가 잘린 버전, gagin(각인)에 전체 문구가 들어있다.
                new_data_full = new_data.copy()
                if gagin in new_data_full.columns:
                    # '품목명' 컬럼에 '각인'에 들어있는 전체 내용을 그대로 넣어준다.
                    new_data_full[dd] = new_data_full[gagin]

                excel_file_name_full = dir_path + last + "네이버_송장발부_전체.xlsx"
                new_data_full.to_excel(excel_file_name_full, index=False, engine="openpyxl")


                excel_file_name = dir_path + last + "네이버_발송처리.xlsx"
                # df.to_excel(excel_file_name, index=False)
                # writer_2 = pd.ExcelWriter(excel_file_name, options={'strings_to_urls': False})
                send_data[nn] = send_data[nn].astype(str)
                send_data[mm] = send_data[mm].astype(str)
                send_data.to_excel(excel_file_name, index=False, sheet_name='발송처리', engine='openpyxl')
                # writer_2.save()

                # self는 현재 클래스(ExcelCalWindow)를 의미합니다.
                QtWidgets.QMessageBox.information(self, '엑셀로 저장', '엑셀 파일로 저장했습니다. 꼬꼬님')

        except Exception as e:
            print(e)
            return 0

    def my_coopang(self):

        try:
            # x = Test_check(self)
            # # self.mytestin.setText("GootEvening")
            # # self.mytestin.setDisabled(True)
            # x.start()

            a = '수취인이름'
            aa = "받으시는 분"

            b = '수취인전화번호'
            bb = '받으시는 분 전화'
            c = '수취인전화번호'
            cc = '받는분핸드폰'
            d = '등록옵션명'
            dd = '품목명'
            e = '구매수(수량)'
            ee = '수량'
            f = '배송메세지'
            ff = '특기사항'
            g = '주문자 추가메시지'
            gg = '메모1'
            h = '수취인 주소'
            hh = '기본배송지'
            i = '상세배송지'  # ?? 비워놓음
            ii = '상세배송지'
            j = '우편번호'
            jj = '받는분우편번호'
            k = '구매자'
            kk = '구매자명'
            l = '구매자전화번호'
            ll = '구매자연락처'
            m = '묶음배송번호'
            mm = '출고번호'
            n = '주문번호'
            nn = '상품주문번호'

            ###################
            o = ''
            oo = '운임Type'
            p = ''
            pp = '지불조건'
            ###################

            q = '1년 주문건수'
            qq = '1년 주문건수'

            ##################

            r = '배송방법'
            s = '택배사'
            t = '송장번호'

            ############ 추가

            gagin = "각인"

            work_status = "작업유무"

            ##############쿠팡 추가

            is_option = "최초등록등록상품명/옵션명"
            option = "상품옵션명"

            ##############엑셀
            file_path, ext = QFileDialog.getOpenFileName(self, '파일 열기', os.getcwd(), 'excel file (*.xls *.xlsx)')
            if file_path:

                print("file_path", file_path)

                self.df_list = self.get_df_from_non_password_excel(file_path)

                # self.df_list = self.loadData(file_path)[0]

                print("결과???", self.df_list)

                # self.df_list.insert(2, c, "s")
                data_o = "s"
                self.df_list[str(oo)] = data_o
                data_p = "신용"
                self.df_list[str(pp)] = data_p

                # data_s = "한진택배"
                # self.df_list[str(s)] = data_s

                print(self.df_list)

                self.df_list.rename(
                    columns={a: aa, c: cc, d: dd, e: ee, f: ff, h: hh, g: gg, j: jj, k: kk, l: ll, m: mm, n: nn,
                             is_option: option}, inplace=True)
                print("재정비")

                self.df_list[gagin] = ""

                print(self.df_list)

                print("원하는 열 나열하기1")
                columns_list = self.df_list[
                    [aa, cc, dd, ee, ff, gg, hh, jj, kk, ll, mm, nn, oo, pp, gagin, option]].values.tolist()

                print("원하는 열 나열하기2")
                print(columns_list)
                #

                # 출고번호(m)에서 동일한거 묶고

                # 메모1(g)과 수량(e)를 품목명(d)에 다시 수정한다.

                print("특정 열 나열하기")
                columns_list_m = self.df_list[mm].values.tolist()
                print(columns_list_m)
                print("[[특정 열 중복제거]]")
                result_set_1 = set(columns_list_m)
                print("특정 열 중복제거 준비", result_set_1)
                result_set_1_len = len(result_set_1)

                result_set_2 = list(result_set_1)
                print("특정 열 중복제거 완료", result_set_2)
                result_set_2_len = len(result_set_2)
                #

                ############
                # # data_list의 데이터를 데이터프레임으로 변환
                # data = [row.split(":") for row in new_list]
                # df = pd.DataFrame(data, columns=["사용자", "게임", "서버", "다이아", "골드", "컴퓨터번호"])
                #############

                # 엑셀 저장 날짜 및 시간을 엑셀 파일명으로...
                nowDay_ = datetime.today().strftime("%Y%m%d%H%M%S")
                year = datetime.today().strftime("%Y")
                month = datetime.today().strftime("%m")
                day = datetime.today().strftime("%d")
                hour = datetime.today().strftime("%H")
                minute = datetime.today().strftime("%M")
                last = str(day) + "d_" + str(hour) + "h" + str(minute) + "m"
                nowDay = str(nowDay_)

                ################ 쿠팡 송장 발부 #############
                df = pd.DataFrame(columns_list,
                                  columns=[aa, cc, dd, ee, ff, gg, hh, jj, kk, ll, mm, nn, oo, pp, gagin, option])
                dir_path = "C:/my_games/excel_result/" + str(year) + "/" + str(month) + "/"
                if not os.path.isdir(dir_path):
                    os.makedirs(dir_path)

                temporary_data = pd.DataFrame(
                    columns=[aa, cc, dd, ee, ff, gg, hh, jj, kk, ll, mm, nn, oo, pp, gagin, option])

                # 쿠팡 송장발부용
                new_data = pd.DataFrame(columns=[aa, cc, dd, ee, ff, gg, hh, jj, kk, ll, mm, nn, oo, pp, gagin, option])

                # dd 품목명, ee 수량, gg 메모1
                print("new_data 1", new_data)
                for set_2 in range(result_set_2_len):

                    # 출고번호로 검색한 것...
                    result = df[df[mm] == result_set_2[set_2]]
                    print("???", result)
                    print("???", result.index)
                    # print("???", result.index[0])
                    print("???gg", result[gg])
                    # print("???gg[0]", result[gg][result.index[0]])
                    # print("???", len(result.index))
                    QTest.qWait(100)
                    # 검색된 것을 임시로 저장하기

                    print("뽑기전 temporary_data", temporary_data)

                    if len(temporary_data) == 0:
                        print("전체저장")
                        temporary_data = pd.concat([temporary_data, result])
                    else:
                        print("삭제 후 저장")
                        temporary_data = temporary_data.drop(temporary_data.index[:len(temporary_data)])
                        temporary_data = pd.concat([temporary_data, result])
                    print("뽑아낸 temporary_data", temporary_data)

                    total_ = 0

                    # 뽑아낸 주문번호의 중복갯수로 다시 반복문 돌려서 작업하기
                    for set_index in range(len(result.index)):
                        print("여기인가 1")
                        ###############

                        # 해당문구 추출하여 갯수파악하고 바로 수량 적어놓기
                        # 해당문구가 품목명에 있을 경우 pass 하기기
                        # df['col_name'].str.contains("여기에문구...")

                        ################
                        # 'gg' 문구 갯수 파악
                        # many = temporary_data[temporary_data[gg] == temporary_data[gg][result.index[set_index]]]
                        # print("how many???", len(many.index))

                        # print("set_index", set_index)
                        # print("set_index?", result.index[set_index])
                        add_write = "=> " + str(result[ee][result.index[set_index]]) + " ea"

                        write_num = set_index + 1
                        result_write_num = str(write_num) + ". "

                        if set_index == 0:
                            # print("여기인가 2")
                            # print("여기인가 2", df.loc[result.index[set_index], gg])
                            # print("여기인가 2", str(add_write))
                            # test = df.loc[result.index[set_index], gg] + str(add_write)
                            # print("add_write", add_write)
                            # print("test", test)

                            # print("df.loc[result.index[set_index], gg]", df.loc[result.index[set_index], gg])
                            if pd.isnull(df.loc[result.index[set_index], gg]) == True:
                                # if len(df.loc[result.index[set_index], gg]) == 0:
                                df.loc[result.index[set_index], dd] = result_write_num + df.loc[
                                    result.index[set_index], option] + ":" + df.loc[result.index[set_index], dd] + str(
                                    add_write)

                                df.loc[result.index[set_index], gagin] = result_write_num + df.loc[
                                    result.index[set_index], option] + ":" + df.loc[result.index[set_index], dd] + str(
                                    add_write)
                            else:

                                write_memo = str(df.loc[result.index[set_index], gg]).replace("여기에 문구:", "")
                                write_memo = str(write_memo).replace("여기에 각인 문구:", "")
                                df.loc[result.index[set_index], dd] = result_write_num + df.loc[
                                    result.index[set_index], option] + ":" + write_memo + str(add_write)

                                df.loc[result.index[set_index], gagin] = result_write_num + df.loc[
                                    result.index[set_index], option] + ":" + write_memo + str(add_write)

                            df_2 = df.iloc[result.index[0]]
                            print("set_index == 0", df_2)
                            new_data.loc[len(new_data)] = df_2

                            # new_data.loc[set_2, dd] = new_data.loc[set_2, gg] + "\n" + df.loc[result.index[set_index], gg]

                        else:

                            print("수정 및 추가하기1", new_data)
                            print("수정 및 추가하기2", result.index[set_index])

                            # [송장발부]마지막으로 같은 문구 갯수 파악하기
                            result_memo = new_data[new_data[dd] == result.loc[result.index[set_index], gg]] + add_write

                            print("수정 및 추가하기3", result_memo)
                            print("수정 및 추가하기4", df.loc[result.index[set_index], gg])

                            if pd.isnull(df.loc[result.index[set_index], gg]) == True:
                                print("this is nan")

                            if len(result_memo.index) == 0:
                                print("없당.", result_memo[gg])
                                if pd.isnull(df.loc[result.index[set_index], gg]) == True:
                                    new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + result_write_num + \
                                                              df.loc[result.index[set_index], option] + ":" + df.loc[
                                                                  result.index[set_index], dd] + add_write
                                    new_data.loc[set_2, gagin] = new_data.loc[set_2, gagin] + "\n" + result_write_num + \
                                                                 df.loc[result.index[set_index], option] + ":" + df.loc[
                                                                     result.index[set_index], dd] + add_write
                                else:

                                    write_memo = str(df.loc[result.index[set_index], gg]).replace("여기에 문구:", "")
                                    write_memo = str(write_memo).replace("여기에 각인 문구:", "")

                                    new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + result_write_num + \
                                                              df.loc[result.index[
                                                                  set_index], option] + ":" + write_memo + add_write
                                    new_data.loc[set_2, gagin] = new_data.loc[set_2, gagin] + "\n" + result_write_num + \
                                                                 df.loc[result.index[
                                                                     set_index], option] + ":" + write_memo + add_write
                            else:
                                print("송장발부는 이미 있다..", result_memo[gg])

                            # # 마지막으로 같은 문구 갯수 파악하기
                            # result_memo = df[df[gg] == df.loc[result.index[set_index], gg]]
                            #
                            # if len(result_memo.index) == 1:
                            #     new_data.loc[set_2, dd] = new_data.loc[set_2, gg] + "\n" + df.loc[result.index[set_index], gg] + "=> " + str(len(result_memo.index)) +"개"
                            # else:
                            #     print("g")
                            # new_data.loc[set_2, gg] = new_data.loc[set_2, gg] + "\n" + df.loc[result.index[set_index], gg]

                        # print("new_data for", new_data)
                        # print("new_data for2", new_data[dd])

                        # total 갯수 구하기 및 1년 주문건수

                        title_display_count = 11
                        print("total_total_total_total_total_total_", result[ee][result.index[set_index]])
                        total_ += int(result[ee][result.index[set_index]])
                        print("total_", total_)
                        if len(result.index) - 1 == set_index:

                            # 몇개 표시할지...

                            # new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + "total => " + str(total_) + " ea"
                            new_data.loc[set_2, gagin] = "[hobby brown] total => " + str(total_) + " ea" + "\n\n" + \
                                                         new_data.loc[set_2, dd]

                            # 1년 주문건수
                            # if new_data.loc[set_2, qq] > 0:
                            # new_data.loc[set_2, gagin] = "1년 주문건수 : " + str(new_data.loc[set_2, qq]) + "건" + "\n" + \
                            #                              new_data.loc[set_2, gagin]

                            if '\n' in new_data.loc[set_2, dd]:

                                result_split = new_data.loc[set_2, dd].split('\n')
                                print("new_data.loc[set_2, dd].split('\n')", result_split)

                                if len(result_split) < title_display_count + 1:
                                    new_data.loc[set_2, dd] = new_data.loc[set_2, dd]
                                    new_data.loc[set_2, dd] = "[hobby brown] total => " + str(total_) + " ea" + "\n\n" + \
                                                              new_data.loc[
                                                                  set_2, dd]
                                else:

                                    for w in range(title_display_count):
                                        if w == 0:

                                            new_data.loc[set_2, dd] = result_split[w]
                                        elif w < title_display_count - 3:
                                            new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + result_split[w]
                                        elif w == title_display_count - 1:
                                            new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + str("^_~")
                                        else:
                                            new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + str(".")
                                    # new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + "total => " + str(total_) + " ea"
                                    new_data.loc[set_2, dd] = "[hobby brown] total => " + str(total_) + " ea" + "\n\n" + \
                                                              new_data.loc[set_2, dd]
                            else:
                                # new_data.loc[set_2, dd] = new_data.loc[set_2, dd] + "\n" + "total => " + str(total_) + " ea"
                                new_data.loc[set_2, dd] = "[hobby brown] total => " + str(total_) + " ea" + "\n\n" + \
                                                          new_data.loc[set_2, dd]
                                # print("new_data 2", new_data)

                # print("new_data 3", len(new_data))

                # 수량 1로 바꾸기
                for many in range(len(new_data)):
                    new_data.loc[many, ee] = 1
                # 메모1 빼고 '1년 구매건수'로 바꾸기
                new_data = pd.DataFrame(new_data,
                                        columns=[aa, bb, cc, dd, ee, ff, qq, hh, ii, jj, kk, ll, mm, nn, oo, pp, gagin])

                #################################################
                #################################################
                #################################################
                print("발송처리 부분")

                ###################################################
                ###################################################
                ##################################################
                new_data.astype(str)

                new_data[work_status] = ""

                # 데이터프레임을 엑셀 파일로 저장
                excel_file_name = dir_path + last + "쿠팡_송장발부.xlsx"
                # df.to_excel(excel_file_name, index=False)
                # writer_1 = pd.ExcelWriter(excel_file_name, options={'strings_to_urls': False}
                new_data[dd] = new_data[dd].astype(str)
                new_data[mm] = new_data[mm].astype(str)
                new_data[nn] = new_data[nn].astype(str)
                new_data.to_excel(excel_file_name, index=False, engine="openpyxl")
                # writer_1.save()

                # === 추가: 품목 전체 표기 버전 엑셀도 함께 저장 ===
                new_data_full = new_data.copy()
                if gagin in new_data_full.columns:
                    new_data_full[dd] = new_data_full[gagin]

                excel_file_name_full = dir_path + last + "쿠팡_송장발부_전체.xlsx"
                new_data_full.to_excel(excel_file_name_full, index=False, engine="openpyxl")


                # self는 현재 클래스(ExcelCalWindow)를 의미합니다.
                QtWidgets.QMessageBox.information(self, '엑셀로 저장', '엑셀 파일로 저장했습니다. 꼬꼬님')

        except Exception as e:
            print(e)
            return 0

    def get_df_from_password_excel(self, excelpath, password):

        df = pd.DataFrame()
        temp = io.BytesIO()
        with open(excelpath, 'rb') as f:
            excel = msoffcrypto.OfficeFile(f)
            excel.load_key(password)
            excel.decrypt(temp)
            # df = pd.read_excel(temp, skiprows=[0], converters={"code": lambda x: str(x)})
            df = pd.read_excel(temp, skiprows=[0])
            del temp
            # df = df.drop(1)
        return df

    def get_df_from_non_password_excel(self, file_name):

        df_list = []
        with pd.ExcelFile(file_name) as wb:
            for i, sn in enumerate(wb.sheet_names):
                try:

                    pd.set_option('display.max_columns', None)
                    pd.set_option('display.max_rows', None)
                    df = pd.read_excel(wb, sheet_name=sn, engine='openpyxl')

                except Exception as e:
                    print('File read error:', e)
                else:
                    df = df.fillna(0)
                    print("dfdfdfdfdfdfdfdf", df)
                    print("snsnsnsnsnsnsnsnsnsnsn", sn)
                    df.name = sn
                    df_list.append(df)

        return df


def main():
    app = QtWidgets.QApplication(sys.argv)
    win = ExcelCalWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()