# read_excel.py
# 네이버·쿠팡 송장 결과 엑셀을 읽어서 화면에 보여주는 탭

import os
from typing import Optional

import pandas as pd
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QComboBox,
    QFileDialog,
    QTableWidget,
    QTableWidgetItem,
    QPlainTextEdit,
)


class ReadInvoiceWidget(QWidget):
    """
    네이버 송장 / 쿠팡 송장으로 나온 엑셀 파일을 읽어서
    테이블로 보여주는 전용 탭 위젯.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # ------------------------------------------------------------------
        # 상단: 타입 선택 + 파일 불러오기 버튼
        # ------------------------------------------------------------------
        top_layout = QHBoxLayout()

        lbl_type = QLabel("송장 타입:")
        self.combo_type = QComboBox()
        self.combo_type.addItems(["네이버 송장", "쿠팡 송장"])

        self.lbl_file = QLabel("선택된 파일: (없음)")
        self.btn_open = QPushButton("엑셀 불러오기")

        top_layout.addWidget(lbl_type)
        top_layout.addWidget(self.combo_type)
        top_layout.addSpacing(20)
        top_layout.addWidget(self.lbl_file, 1)
        top_layout.addSpacing(20)
        top_layout.addWidget(self.btn_open)

        main_layout.addLayout(top_layout)

        # ------------------------------------------------------------------
        # 가운데: 엑셀 내용을 보여줄 테이블
        # ------------------------------------------------------------------
        self.table = QTableWidget()
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        main_layout.addWidget(self.table, 1)

        # ------------------------------------------------------------------
        # 하단: 로그/안내 영역
        # ------------------------------------------------------------------
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setFixedHeight(150)
        main_layout.addWidget(self.log)

        # 시그널 연결
        self.btn_open.clicked.connect(self.on_click_open)

        # 내부 상태
        self.current_df: Optional[pd.DataFrame] = None
        self.current_file: Optional[str] = None

    # ----------------------------------------------------------------------
    # UI 핸들러
    # ----------------------------------------------------------------------
    def on_click_open(self):
        """
        엑셀 파일을 선택하고, 송장 타입에 맞게 읽어와서 테이블에 표시
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "송장 엑셀 파일 선택",
            r"C:\my_games\excel_result",  # ← 기본 폴더 고정
            "Excel Files (*.xlsx *.xlsm *.xltx *.xltm);;All Files (*.*)",
        )
        if not file_path:
            return

        self.current_file = file_path
        self.lbl_file.setText(f"선택된 파일: {os.path.basename(file_path)}")

        invoice_type = self.combo_type.currentText()
        try:
            if invoice_type == "네이버 송장":
                df = self._load_naver_invoice(file_path)
            else:
                df = self._load_coupang_invoice(file_path)

            self.current_df = df
            self._show_df_in_table(df)
            self._log_columns(df, invoice_type, file_path)

        except Exception as e:
            self.log.appendPlainText(f"[오류] 엑셀을 읽는 중 문제가 발생했습니다: {e}")

    # ----------------------------------------------------------------------
    # 엑셀 파싱 로직 (초기 버전: 일단 전체 시트/전체 컬럼을 그대로 보여줌)
    # 필요하면 여기 컬럼 매핑만 손보면 됨.
    # ----------------------------------------------------------------------
    def _load_naver_invoice(self, file_path: str) -> pd.DataFrame:
        """
        네이버 송장 결과 엑셀 읽기.
        - 현재는 '첫 번째 시트 전체'를 읽어서 반환.
        - 실제 포맷에 맞춰 특정 컬럼만 골라내고 싶으면 여기서 처리.
        """
        # 기본: 첫 번째 시트 전체
        df = pd.read_excel(file_path)

        # 예시) 특정 컬럼이 있을 경우 골라내는 형태 (필요하면 사용)
        # wanted_cols = ["상품명", "수량", "단가", "공급가액", "부가세"]
        # existing = [c for c in wanted_cols if c in df.columns]
        # if existing:
        #     df = df[existing]

        return df

    def _load_coupang_invoice(self, file_path: str) -> pd.DataFrame:
        """
        쿠팡 송장 결과 엑셀 읽기.
        - 현재는 '첫 번째 시트 전체'를 읽어서 반환.
        - 실제 포맷에 맞춰 특정 컬럼만 골라내고 싶으면 여기서 처리.
        """
        df = pd.read_excel(file_path)

        # 예시) 특정 컬럼이 있을 경우 골라내는 형태 (필요하면 사용)
        # wanted_cols = ["상품명", "수량", "매입금액", "공급가액", "부가세"]
        # existing = [c for c in wanted_cols if c in df.columns]
        # if existing:
        #     df = df[existing]

        return df

    # ----------------------------------------------------------------------
    # DataFrame → QTableWidget 표시
    # ----------------------------------------------------------------------
    def _show_df_in_table(self, df: pd.DataFrame):
        self.table.clear()

        # 헤더 설정
        self.table.setColumnCount(len(df.columns))
        self.table.setRowCount(len(df))

        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])

        # 데이터 채우기
        for row_idx in range(len(df)):
            for col_idx, col_name in enumerate(df.columns):
                val = df.iloc[row_idx, col_idx]
                if pd.isna(val):
                    text = ""
                else:
                    text = str(val)
                item = QTableWidgetItem(text)
                self.table.setItem(row_idx, col_idx, item)

        self.table.resizeColumnsToContents()

    # ----------------------------------------------------------------------
    # 로그 출력
    # ----------------------------------------------------------------------
    def _log_columns(self, df: pd.DataFrame, invoice_type: str, file_path: str):
        self.log.appendPlainText(
            f"▶ [{invoice_type}] 엑셀 읽기 완료: {os.path.basename(file_path)}"
        )
        self.log.appendPlainText(f"  - 행 수: {len(df)}, 열 수: {len(df.columns)}")
        self.log.appendPlainText("  - 컬럼 목록:")

        for c in df.columns:
            self.log.appendPlainText(f"     · {c}")

        self.log.appendPlainText("-" * 40)
