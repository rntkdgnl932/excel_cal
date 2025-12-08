# read_excel.py
# 네이버·쿠팡 송장 엑셀을 읽어와서 보여주고,
# 품목명 파싱, 복사용 문구 COPY, 사진 첨부/삭제/재사용,
# 문자 전송 UI 뼈대까지 포함한 탭 위젯.

import os
import re
import json
import shutil
from pathlib import Path
from typing import Optional, List, Dict

import pandas as pd
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import Qt, QDateTime
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
    QDialog,
    QLineEdit,
    QApplication,
    QListWidget,
    QListWidgetItem,
    QGroupBox,
    QDateTimeEdit,
    QCheckBox,
)


# ----------------------------------------------------------------------
# COPY 다이얼로그: 파싱된 문구들 + 각 줄별 COPY 버튼
# ----------------------------------------------------------------------
class CopyLinesDialog(QDialog):
    def __init__(self, lines: List[str], parent=None):
        super().__init__(parent)

        self.setWindowTitle("문구 복사")
        self.resize(600, 400)

        self._line_edits: List[QLineEdit] = []
        self._buttons: List[QPushButton] = []

        main_layout = QVBoxLayout(self)

        info = QLabel("복사하려는 문구 옆의 [COPY] 버튼을 클릭하세요.")
        main_layout.addWidget(info)

        # 마지막 복사 상태 표시 라벨
        self.lbl_last = QLabel("마지막 복사: 없음")
        main_layout.addWidget(self.lbl_last)

        for idx, text in enumerate(lines):
            row_layout = QHBoxLayout()

            edit = QLineEdit()
            edit.setReadOnly(True)
            edit.setText(text)

            btn_copy = QPushButton("COPY")

            # 핸들러 연결
            btn_copy.clicked.connect(self._make_copy_handler(idx, text))

            row_layout.addWidget(edit, 1)
            row_layout.addWidget(btn_copy)

            main_layout.addLayout(row_layout)

            self._line_edits.append(edit)
            self._buttons.append(btn_copy)

        btn_close = QPushButton("닫기")
        btn_close.clicked.connect(self.accept)
        main_layout.addWidget(btn_close)

    def _make_copy_handler(self, idx: int, text: str):
        def handler():
            # 1) 클립보드 복사
            QApplication.clipboard().setText(text)
            # 2) 표시 업데이트
            self._mark_copied(idx, text)
        return handler

    def _mark_copied(self, idx: int, text: str):
        """
        idx 번째 줄을 '마지막 복사'로 표시.
        - 해당 줄 버튼: '✔ COPIED'
        - 해당 줄 배경: 연한 노랑
        - 다른 줄은 모두 초기화
        """
        # 전부 초기화
        for btn in self._buttons:
            btn.setText("COPY")
        for edit in self._line_edits:
            edit.setStyleSheet("")

        # 선택된 줄만 강조
        if 0 <= idx < len(self._buttons):
            self._buttons[idx].setText("✔ COPIED")
        if 0 <= idx < len(self._line_edits):
            self._line_edits[idx].setStyleSheet(
                "background-color: #fff8c6;"  # 연한 노랑
            )

        # 라벨도 갱신
        self.lbl_last.setText(f"마지막 복사: {text}")



# ----------------------------------------------------------------------
# 이미지 관리 다이얼로그: 행별 여러 장 추가/삭제/미리보기
# ----------------------------------------------------------------------
class ImageManageDialog(QDialog):
    def __init__(
        self,
        parent,
        row_id: int,
        image_dir: Path,
        current_files: List[str],
    ):
        super().__init__(parent)

        self.setWindowTitle(f"사진 관리 - 행 {row_id}")
        self.resize(600, 400)

        self.row_id = row_id
        self.image_dir = image_dir
        self._images: List[str] = list(current_files)

        main_layout = QVBoxLayout(self)

        top_layout = QHBoxLayout()
        main_layout.addLayout(top_layout)

        # 좌측: 리스트
        self.list_widget = QListWidget()
        self.list_widget.currentRowChanged.connect(self._on_list_selection_changed)
        top_layout.addWidget(self.list_widget, 2)

        # 우측: 미리보기
        right_layout = QVBoxLayout()
        top_layout.addLayout(right_layout, 3)

        self.lbl_preview = QLabel("미리보기 없음")
        self.lbl_preview.setFrameShape(QLabel.Box)
        self.lbl_preview.setAlignment(Qt.AlignCenter)
        self.lbl_preview.setFixedSize(260, 260)
        right_layout.addWidget(self.lbl_preview)

        self.lbl_filename = QLabel("")
        right_layout.addWidget(self.lbl_filename)

        # 하단 버튼들
        btn_layout = QHBoxLayout()
        main_layout.addLayout(btn_layout)

        self.btn_add = QPushButton("+ 추가")
        self.btn_primary = QPushButton("대표로")   # ✅ 대표 이미지로 올리기
        self.btn_del = QPushButton("- 삭제")
        self.btn_close = QPushButton("닫기")

        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_primary)
        btn_layout.addWidget(self.btn_del)
        btn_layout.addStretch(1)
        btn_layout.addWidget(self.btn_close)

        self.btn_add.clicked.connect(self._on_add)
        self.btn_primary.clicked.connect(self._on_set_primary)
        self.btn_del.clicked.connect(self._on_del)
        self.btn_close.clicked.connect(self.accept)

        # 초기 리스트 로드
        self._reload_list()

    # 외부에서 결과 조회용
    def images(self) -> List[str]:
        return list(self._images)

    # 리스트 갱신
    def _reload_list(self):
        self.list_widget.clear()

        for fname in self._images:
            item = QListWidgetItem(fname)
            fpath = self.image_dir / fname
            if fpath.is_file():
                pix = QtGui.QPixmap(str(fpath))
                if not pix.isNull():
                    icon = QtGui.QIcon(
                        pix.scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    )
                    item.setIcon(icon)
            self.list_widget.addItem(item)

        if self._images:
            # 항상 첫 번째 항목을 선택 상태로
            self.list_widget.setCurrentRow(0)
        else:
            self.lbl_preview.setText("미리보기 없음")
            self.lbl_preview.setPixmap(QtGui.QPixmap())
            self.lbl_filename.setText("")

    # 리스트 선택 변경 시 미리보기 갱신
    def _on_list_selection_changed(self, row: int):
        if row < 0 or row >= len(self._images):
            self.lbl_preview.setText("미리보기 없음")
            self.lbl_preview.setPixmap(QtGui.QPixmap())
            self.lbl_filename.setText("")
            return

        fname = self._images[row]
        fpath = self.image_dir / fname
        if not fpath.is_file():
            self.lbl_preview.setText("파일 없음")
            self.lbl_preview.setPixmap(QtGui.QPixmap())
            self.lbl_filename.setText(fname)
            return

        pix = QtGui.QPixmap(str(fpath))
        if pix.isNull():
            self.lbl_preview.setText("이미지 로드 실패")
            self.lbl_preview.setPixmap(QtGui.QPixmap())
            self.lbl_filename.setText(fname)
            return

        scaled = pix.scaled(
            self.lbl_preview.size(),
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation,
        )
        self.lbl_preview.setPixmap(scaled)
        self.lbl_filename.setText(fname)

    # 이미지 추가
    def _on_add(self):
        src_path, _ = QFileDialog.getOpenFileName(
            self,
            "첨부할 이미지 선택",
            str(self.image_dir),
            "Images (*.png *.jpg *.jpeg *.bmp *.gif);;All Files (*.*)",
        )
        if not src_path:
            return

        src = Path(src_path)
        ext = src.suffix.lower() or ".png"

        next_idx = len(self._images) + 1
        new_name = f"row_{self.row_id:04d}_{next_idx}{ext}"

        try:
            self.image_dir.mkdir(parents=True, exist_ok=True)
            (self.image_dir / new_name).write_bytes(src.read_bytes())
        except (OSError, IOError) as e:
            QtWidgets.QMessageBox.critical(self, "복사 실패", str(e))
            return

        self._images.append(new_name)
        self._reload_list()

    # ✅ 선택한 이미지를 "대표"로: 리스트 맨 앞으로 이동
    def _on_set_primary(self):
        row = self.list_widget.currentRow()
        if row <= 0 or row >= len(self._images):
            # 0 이하면 이미 대표거나 선택 없음
            return

        fname = self._images.pop(row)
        self._images.insert(0, fname)
        self._reload_list()
        self.list_widget.setCurrentRow(0)  # 대표 선택 상태로 유지

    # 이미지 삭제
    def _on_del(self):
        row = self.list_widget.currentRow()
        if row < 0 or row >= len(self._images):
            return

        fname = self._images.pop(row)
        fpath = self.image_dir / fname
        try:
            if fpath.is_file():
                fpath.unlink()
        except (OSError, IOError):
            pass

        self._reload_list()



# ----------------------------------------------------------------------
# 메인 탭 위젯
# ----------------------------------------------------------------------
class ReadInvoiceWidget(QWidget):
    """
    네이버 송장 / 쿠팡 송장으로 나온 엑셀 파일을 읽어서
    테이블로 보여주는 전용 탭 위젯.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        self.current_df: Optional[pd.DataFrame] = None
        self.current_file: Optional[str] = None

        # 이미지 매핑: row_id(1부터) -> [파일명, ...]
        self._image_map: Dict[int, List[str]] = {}
        self._image_dir: Optional[Path] = None
        self._meta_path: Optional[Path] = None

        # 테이블 컬럼 이름 → index 매핑
        self._col_index: Dict[str, int] = {}
        self._current_row_idx: Optional[int] = None

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # 상단
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

        # 테이블
        self.table = QTableWidget()
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(True)
        main_layout.addWidget(self.table, 1)

        # 문자 전송 패널
        self._build_sms_panel(main_layout)

        # 로그
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setFixedHeight(150)
        main_layout.addWidget(self.log)

        # 시그널
        self.btn_open.clicked.connect(self.on_click_open)
        self.table.itemDoubleClicked.connect(self.on_item_double_clicked)
        self.table.itemSelectionChanged.connect(self.on_table_selection_changed)

    # ------------------------------------------------------------------
    # 문자 전송 UI
    # ------------------------------------------------------------------
    def _build_sms_panel(self, parent_layout: QVBoxLayout):
        grp = QGroupBox("문자 전송")
        layout = QHBoxLayout(grp)

        # 왼쪽: 선택 고객 정보
        left_box = QVBoxLayout()
        self.lbl_sms_target = QLabel("선택 고객: (없음)")
        self.lbl_sms_count = QLabel("문구개수: -")
        left_box.addWidget(self.lbl_sms_target)
        left_box.addWidget(self.lbl_sms_count)
        left_box.addStretch(1)
        layout.addLayout(left_box, 3)

        # 가운데: 예약 시간 + 버튼들
        mid_box = QVBoxLayout()

        time_box = QHBoxLayout()
        lbl_time = QLabel("보내는 시간:")
        self.dt_send = QDateTimeEdit(QDateTime.currentDateTime())
        self.dt_send.setCalendarPopup(True)
        self.chk_send_now = QCheckBox("지금 보내기")
        self.chk_send_now.setChecked(True)
        self.chk_send_now.toggled.connect(self._on_send_now_toggled)

        time_box.addWidget(lbl_time)
        time_box.addWidget(self.dt_send)
        time_box.addWidget(self.chk_send_now)
        mid_box.addLayout(time_box)

        btn_box = QHBoxLayout()
        self.btn_send_selected = QPushButton("선택 고객에게 보내기")
        self.btn_send_all = QPushButton("표시된 전체 고객에게 일괄 보내기")
        btn_box.addWidget(self.btn_send_selected)
        btn_box.addWidget(self.btn_send_all)
        mid_box.addLayout(btn_box)

        mid_box.addStretch(1)
        layout.addLayout(mid_box, 5)

        # 오른쪽: 대표 이미지 미리보기
        right_box = QVBoxLayout()
        self.lbl_img_preview = QLabel("대표 이미지\n미리보기 없음")
        self.lbl_img_preview.setFrameShape(QLabel.Box)
        # noinspection PyUnresolvedReferences
        self.lbl_img_preview.setAlignment(Qt.AlignCenter)
        self.lbl_img_preview.setFixedSize(160, 160)

        self.lbl_img_name = QLabel("")
        right_box.addWidget(self.lbl_img_preview)
        right_box.addWidget(self.lbl_img_name)
        right_box.addStretch(1)
        layout.addLayout(right_box, 3)

        parent_layout.addWidget(grp)

        self.btn_send_selected.clicked.connect(self.on_send_selected)
        self.btn_send_all.clicked.connect(self.on_send_all)

    def _on_send_now_toggled(self, checked: bool):
        self.dt_send.setEnabled(not checked)

    # ------------------------------------------------------------------
    # UI 핸들러
    # ------------------------------------------------------------------
    def on_click_open(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "송장 엑셀 파일 선택",
            r"C:\my_games\excel_result",
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

            df = self._add_item_count_column(df, invoice_type)
            self.current_df = df

            self._setup_image_store()
            self._show_df_in_table(df)
            self._log_columns(df, invoice_type, file_path)

        except (OSError, IOError, ValueError) as e:
            QtWidgets.QMessageBox.critical(self, "엑셀 읽기 오류", str(e))
            self.log.appendPlainText(f"[오류] 엑셀을 읽는 중 문제가 발생했습니다: {e}")

    def on_item_double_clicked(self, item: QTableWidgetItem):
        if self.current_df is None:
            return

        row = item.row()
        col = item.column()

        header = self.table.horizontalHeaderItem(col)
        col_name = header.text() if header is not None else ""

        if col_name not in ("품목명", "상품명"):
            self.log.appendPlainText(
                "※ 품목명/상품명 컬럼에서만 더블클릭 복사가 동작합니다."
            )
            return

        cell_text = self.table.item(row, col).text() if self.table.item(row, col) else ""
        if not cell_text.strip():
            return

        invoice_type = self.combo_type.currentText()

        if invoice_type == "네이버 송장":
            lines = self._parse_naver_lines(cell_text)
        else:
            lines = self._parse_coupang_lines(cell_text)

        if not lines:
            self.log.appendPlainText("※ 파싱된 문구가 없습니다. 패턴을 한 번 확인해 주세요.")
            return

        dlg = CopyLinesDialog(lines, self)
        dlg.exec_()

    def on_table_selection_changed(self):
        selected = self.table.selectionModel().selectedRows()
        if not selected:
            self._current_row_idx = None
            self.lbl_sms_target.setText("선택 고객: (없음)")
            self.lbl_sms_count.setText("문구개수: -")
            self._clear_preview()
            return

        row_idx = selected[0].row()
        self._current_row_idx = row_idx
        self._update_sms_panel_for_row(row_idx)

    # ------------------------------------------------------------------
    # 문자 전송 버튼 (현재는 로그만 남김)
    # ------------------------------------------------------------------
    def on_send_selected(self):
        if self._current_row_idx is None:
            QtWidgets.QMessageBox.information(self, "알림", "선택된 고객이 없습니다.")
            return

        name = self._get_cell_text(self._current_row_idx, "받으시는 분")
        phone = self._get_cell_text(self._current_row_idx, "받으시는 분 전화")
        invoice_type = self.combo_type.currentText()
        row_id = self._current_row_idx + 1
        img_count = len(self._image_map.get(row_id, []))

        when = (
            "지금"
            if self.chk_send_now.isChecked()
            else self.dt_send.dateTime().toString("yyyy-MM-dd HH:mm")
        )
        self.log.appendPlainText(
            f"[테스트] [{invoice_type}] 선택 고객 문자 전송 예정: "
            f"{name}({phone}), 행={row_id}, 이미지={img_count}장, 시간={when}"
        )

    def on_send_all(self):
        row_count = self.table.rowCount()
        if row_count == 0:
            QtWidgets.QMessageBox.information(self, "알림", "표시된 고객이 없습니다.")
            return

        invoice_type = self.combo_type.currentText()
        when = (
            "지금"
            if self.chk_send_now.isChecked()
            else self.dt_send.dateTime().toString("yyyy-MM-dd HH:mm")
        )

        self.log.appendPlainText(
            f"[테스트] [{invoice_type}] 표시된 전체 고객({row_count}명)에게 문자 전송 예정. 시간={when}"
        )

    # ------------------------------------------------------------------
    # 엑셀 읽기 (static 가능)
    # ------------------------------------------------------------------
    @staticmethod
    def _load_naver_invoice(file_path: str) -> pd.DataFrame:
        return pd.read_excel(file_path)

    @staticmethod
    def _load_coupang_invoice(file_path: str) -> pd.DataFrame:
        return pd.read_excel(file_path)

    # ------------------------------------------------------------------
    # 이미지 저장 위치 / meta.json 세팅
    # ------------------------------------------------------------------
    def _setup_image_store(self):
        if not self.current_file:
            self._image_dir = None
            self._meta_path = None
            self._image_map = {}
            return

        excel_path = Path(self.current_file)
        base_dir = excel_path.parent / excel_path.stem
        image_dir = base_dir / "images"
        meta_path = base_dir / "meta.json"

        self._image_dir = image_dir
        self._meta_path = meta_path
        self._image_map = {}

        if meta_path.is_file():
            try:
                with meta_path.open("r", encoding="utf-8") as f:
                    data = json.load(f)
                rows = data.get("rows", {})
                self._image_map = {int(k): list(v) for k, v in rows.items()}
            except (OSError, IOError, json.JSONDecodeError, ValueError, TypeError):
                self._image_map = {}

    def _save_image_meta(self):
        if not self._meta_path:
            return

        try:
            self._meta_path.parent.mkdir(parents=True, exist_ok=True)
            data = {"rows": self._image_map}
            with self._meta_path.open("w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except (OSError, IOError, TypeError, ValueError) as e:
            self.log.appendPlainText(f"[경고] meta.json 저장 실패: {e}")

    # ------------------------------------------------------------------
    # DataFrame → QTableWidget 표시
    # ------------------------------------------------------------------
    def _show_df_in_table(self, df: pd.DataFrame):
        self.table.clear()

        cols = list(df.columns)
        photo_col_name = "사진"
        cols_with_photo = cols + [photo_col_name]

        row_count = len(df)
        col_count = len(cols_with_photo)

        self.table.setColumnCount(col_count)
        self.table.setRowCount(row_count)
        self.table.setHorizontalHeaderLabels([str(c) for c in cols_with_photo])

        self._col_index = {name: idx for idx, name in enumerate(cols_with_photo)}

        for row_idx in range(row_count):
            for col_idx, col_name in enumerate(cols):
                val = df.iloc[row_idx, col_idx]
                text = "" if pd.isna(val) else str(val)
                item = QTableWidgetItem(text)
                if col_name in ("품목명", "상품명"):
                    item.setToolTip(text)
                self.table.setItem(row_idx, col_idx, item)

            photo_col_idx = self._col_index[photo_col_name]
            btn = QPushButton()
            row_id = row_idx + 1
            count = len(self._image_map.get(row_id, []))
            btn.setText(f"사진({count}장)…")
            btn.clicked.connect(self._make_photo_button_handler(row_idx))
            self.table.setCellWidget(row_idx, photo_col_idx, btn)

        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()

    def _make_photo_button_handler(self, row_idx: int):
        def handler():
            self._open_image_manager(row_idx)
        return handler

    def _open_image_manager(self, row_idx: int):
        if not self._image_dir:
            QtWidgets.QMessageBox.information(self, "알림", "이미지 저장 폴더를 찾을 수 없습니다.")
            return

        row_id = row_idx + 1
        current_files = self._image_map.get(row_id, [])

        dlg = ImageManageDialog(self, row_id, self._image_dir, current_files)
        if dlg.exec_() == QDialog.Accepted:
            files = dlg.images()
            if files:
                self._image_map[row_id] = files
            else:
                self._image_map.pop(row_id, None)

            self._save_image_meta()
            self._refresh_photo_buttons()
            self._update_preview_for_row(row_idx)

    def _refresh_photo_buttons(self):
        """
        테이블에 있는 사진 버튼 텍스트를 image_map 기준으로 다시 그려준다.
        """
        photo_col_idx = self._col_index.get("사진")
        if photo_col_idx is None:
            return

        for row_idx in range(self.table.rowCount()):
            widget = self.table.cellWidget(row_idx, photo_col_idx)
            if not isinstance(widget, QPushButton):
                continue
            row_id = row_idx + 1
            count = len(self._image_map.get(row_id, []))
            widget.setText(f"사진({count}장)…")

    # ------------------------------------------------------------------
    # 로그 출력
    # ------------------------------------------------------------------
    def _log_columns(self, df: pd.DataFrame, invoice_type: str, file_path: str):
        self.log.appendPlainText(
            f"▶ [{invoice_type}] 엑셀 읽기 완료: {os.path.basename(file_path)}"
        )
        self.log.appendPlainText(f"  - 행 수: {len(df)}, 열 수: {len(df.columns)}")
        self.log.appendPlainText("  - 컬럼 목록:")
        for c in df.columns:
            self.log.appendPlainText(f"     · {c}")
        self.log.appendPlainText("-" * 40)

    # ------------------------------------------------------------------
    # 네이버 / 쿠팡 품목명 파싱 (static)
    # ------------------------------------------------------------------
    @staticmethod
    def _parse_naver_lines(cell_text: str) -> List[str]:
        lines: List[str] = []
        for raw in cell_text.splitlines():
            s = raw.strip()
            if not re.match(r"^\d+\.", s):
                continue
            body = s.split(".", 1)[1].lstrip()
            idx = body.find("/ 각인체")
            if idx != -1:
                core = body[:idx].strip()
            else:
                if "=>" in body:
                    core = body.split("=>", 1)[0].strip()
                else:
                    core = body.strip()
            if core:
                lines.append(core)
        return lines

    @staticmethod
    def _parse_coupang_lines(cell_text: str) -> List[str]:
        lines: List[str] = []
        for raw in cell_text.splitlines():
            s = raw.strip()
            if not re.match(r"^\d+\.", s):
                continue
            if ":" in s:
                right = s.split(":", 1)[1]
            else:
                right = s
            if "=>" in right:
                core = right.split("=>", 1)[0].strip()
            else:
                core = right.strip()
            if core:
                lines.append(core)
        return lines

    # ------------------------------------------------------------------
    # "문구개수" 컬럼 자동 추가
    # ------------------------------------------------------------------
    def _add_item_count_column(
        self, df: pd.DataFrame, invoice_type: str
    ) -> pd.DataFrame:
        col_key = None
        for cand in ("품목명", "상품명"):
            if cand in df.columns:
                col_key = cand
                break
        if col_key is None:
            return df

        counts: List[int] = []
        for idx in range(len(df)):
            val = df.iloc[idx][col_key]
            text = "" if pd.isna(val) else str(val)
            if invoice_type == "네이버 송장":
                lines = self._parse_naver_lines(text)
            else:
                lines = self._parse_coupang_lines(text)
            counts.append(len(lines))

        base_name = "문구개수"
        col_name = base_name
        n = 2
        while col_name in df.columns:
            col_name = f"{base_name}{n}"
            n += 1
        df[col_name] = counts
        return df

    # ------------------------------------------------------------------
    # 문자 패널 / 미리보기 갱신
    # ------------------------------------------------------------------
    def _get_cell_text(self, row_idx: int, col_name: str) -> str:
        col_idx = self._col_index.get(col_name)
        if col_idx is None:
            return ""
        item = self.table.item(row_idx, col_idx)
        return "" if item is None else item.text()

    def _update_sms_panel_for_row(self, row_idx: int):
        name = self._get_cell_text(row_idx, "받으시는 분")
        phone = self._get_cell_text(row_idx, "받으시는 분 전화")
        cnt = self._get_cell_text(row_idx, "문구개수") or "-"
        self.lbl_sms_target.setText(f"선택 고객: {name} ({phone})")
        self.lbl_sms_count.setText(f"문구개수: {cnt}")
        self._update_preview_for_row(row_idx)

    def _clear_preview(self):
        self.lbl_img_preview.setPixmap(QtGui.QPixmap())
        self.lbl_img_preview.setText("대표 이미지\n미리보기 없음")
        self.lbl_img_name.setText("")

    def _update_preview_for_row(self, row_idx: int):
        if not self._image_dir:
            self._clear_preview()
            return

        row_id = row_idx + 1
        files = self._image_map.get(row_id) or []
        if not files:
            self._clear_preview()
            return

        fname = files[0]
        fpath = self._image_dir / fname
        if not fpath.is_file():
            self._clear_preview()
            return

        pix = QtGui.QPixmap(str(fpath))
        if pix.isNull():
            self._clear_preview()
            return
        # noinspection PyUnresolvedReferences
        scaled = pix.scaled(
            self.lbl_img_preview.size(),
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation,
        )
        self.lbl_img_preview.setPixmap(scaled)
        self.lbl_img_name.setText(fname)


# ----------------------------------------------------------------------
# 단독 실행 테스트용
# ----------------------------------------------------------------------
if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    w = ReadInvoiceWidget()
    w.resize(1200, 800)
    w.show()
    sys.exit(app.exec_())
