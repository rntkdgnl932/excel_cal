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
from PyQt5.QtCore import Qt, QDateTime, QTimer
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



# (기존 import 문에 QScrollArea가 없다면 추가해야 하지만,
# 아래 코드처럼 QtWidgets.QScrollArea로 쓰면 import 수정 안 해도 됩니다.)

class CopyLinesDialog(QDialog):
    def __init__(self, full_text: str, invoice_type: str, save_callback, parent=None):
        super().__init__(parent)
        self.setWindowTitle("문구 복사 & 수정")
        self.resize(650, 500)

        self.save_callback = save_callback  # 저장 시 호출할 함수 (엑셀 반영용)
        self.raw_lines = full_text.splitlines()  # 원본 줄들 (보존용)
        self.parsed_items = []  # 화면에 표시할 파싱된 데이터

        # ---------------------------------------------------------
        # 1. 문구 파싱 (수정 시 재조립을 위해 앞/뒤 문맥까지 분리)
        # ---------------------------------------------------------
        for idx, line in enumerate(self.raw_lines):
            # 파싱 로직을 여기로 가져옴
            parts = self._parse_structure(line, invoice_type)
            if parts:
                # parts = (prefix, core, suffix)
                self.parsed_items.append({
                    "line_idx": idx,
                    "prefix": parts[0],
                    "core": parts[1],
                    "suffix": parts[2]
                })

        # ---------------------------------------------------------
        # 2. UI 구성
        # ---------------------------------------------------------
        main_layout = QVBoxLayout(self)

        # 상단 안내
        top_layout = QVBoxLayout()
        info = QLabel("문구를 [수정] 후 [저장]하면 엑셀에도 반영됩니다.")
        self.lbl_last = QLabel("마지막 복사: 없음")
        self.lbl_last.setStyleSheet("color: blue; font-weight: bold;")
        top_layout.addWidget(info)
        top_layout.addWidget(self.lbl_last)
        main_layout.addLayout(top_layout)

        # 스크롤 영역
        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QtWidgets.QFrame.NoFrame)

        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)

        self._ui_rows = []  # (edit, btn_edit, btn_copy, item_data) 저장용

        for item in self.parsed_items:
            row_layout = QHBoxLayout()

            # 텍스트 입력창
            edit = QLineEdit()
            edit.setText(item["core"])
            edit.setReadOnly(True)
            edit.setStyleSheet("background-color: #f0f0f0; color: #333;")

            # 수정/저장 버튼
            btn_edit = QPushButton("수정")
            btn_edit.setFixedWidth(60)

            # 복사 버튼
            btn_copy = QPushButton("COPY")
            btn_copy.setFixedWidth(70)

            # 핸들러 연결
            # 주의: 루프 변수 캡처를 위해 별도 메서드로 연결
            self._connect_handlers(edit, btn_edit, btn_copy, item)

            row_layout.addWidget(edit, 1)
            row_layout.addWidget(btn_edit)
            row_layout.addWidget(btn_copy)
            content_layout.addLayout(row_layout)

            self._ui_rows.append({
                "edit": edit,
                "btn_edit": btn_edit,
                "btn_copy": btn_copy
            })

        content_layout.addStretch(1)
        content_widget.setLayout(content_layout)
        scroll_area.setWidget(content_widget)
        main_layout.addWidget(scroll_area, 1)

        # 닫기 버튼
        btn_close = QPushButton("닫기")
        btn_close.clicked.connect(self.accept)
        main_layout.addWidget(btn_close)

    def _parse_structure(self, line: str, invoice_type: str):
        """
        한 줄을 (접두어, 핵심문구, 접미어)로 분리합니다.
        분해 실패 시(패턴 불일치) None 반환.
        """
        s = line.strip()
        # 공통: 숫자+점(1.) 으로 시작하는지 확인
        if not re.match(r"^\d+\.", s):
            return None

        # 핵심 문구 추출 로직 (기존과 동일하되 위치를 찾음)
        core = ""

        # 1) 앞부분(Prefix) 잘라내기
        # 네이버: "1. 품목명" -> "1. " + "품목명"
        # 쿠팡: "1. 옵션: 품목명" -> "1. 옵션: " + "품목명" (대략적)

        # 편의상 기존 로직을 활용해 'core' 텍스트를 먼저 찾습니다.
        temp_core = ""
        if invoice_type == "네이버 송장":
            body = s.split(".", 1)[1].lstrip()  # "1." 떼고 나머지
            idx = body.find("/ 각인체")
            if idx != -1:
                temp_core = body[:idx].strip()
            elif "=>" in body:
                temp_core = body.split("=>", 1)[0].strip()
            else:
                temp_core = body.strip()
        else:  # 쿠팡
            body = s
            if ":" in body:
                body = body.split(":", 1)[1]
            if "=>" in body:
                temp_core = body.split("=>", 1)[0].strip()
            else:
                temp_core = body.strip()

        if not temp_core:
            return None

        # 2) 원본 문자열에서 temp_core의 위치를 찾아서 정확히 3등분
        # (주의: temp_core가 여러 번 나올 수 있으나, 보통 구조상 한 번 나옴. 첫 번째로 처리)
        start_idx = line.find(temp_core)
        if start_idx == -1:
            return None

        prefix = line[:start_idx]
        suffix = line[start_idx + len(temp_core):]

        return (prefix, temp_core, suffix)

    def _connect_handlers(self, edit, btn_edit, btn_copy, item_data):
        # 수정/저장 버튼 로직
        def on_edit_click():
            if btn_edit.text() == "수정":
                # 수정 모드 진입
                edit.setReadOnly(False)
                edit.setFocus()
                edit.setStyleSheet("background-color: #ffffff; color: #000; border: 2px solid #4dabf7;")
                btn_edit.setText("저장")
                btn_edit.setStyleSheet("color: blue; font-weight: bold;")
                # 저장 모드일 땐 복사 비활성화 (선택)
                btn_copy.setEnabled(False)
            else:
                # 저장 로직 수행
                new_text = edit.text()

                # 1. 데이터 업데이트 (메모리)
                item_data["core"] = new_text

                # 2. 원본 라인 재조립
                new_line = item_data["prefix"] + new_text + item_data["suffix"]
                self.raw_lines[item_data["line_idx"]] = new_line

                # 3. 전체 텍스트 합치기
                new_full_text = "\n".join(self.raw_lines)

                # 4. 부모창(Widget)에 저장 요청!
                self.save_callback(new_full_text)

                # UI 복귀
                edit.setReadOnly(True)
                edit.setStyleSheet("background-color: #f0f0f0; color: #333;")
                btn_edit.setText("수정")
                btn_edit.setStyleSheet("")
                btn_copy.setEnabled(True)

        # 복사 버튼 로직
        def on_copy_click():
            text = edit.text()
            QApplication.clipboard().setText(text)
            self._mark_copied(text)

        btn_edit.clicked.connect(on_edit_click)
        btn_copy.clicked.connect(on_copy_click)

    def _mark_copied(self, text: str):
        # 모든 버튼 'COPY'로 초기화
        for row in self._ui_rows:
            row["btn_copy"].setText("COPY")

        # 현재 누른 버튼 찾아서 '완료' 표시 (sender 이용하거나 해서)
        sender = self.sender()
        if sender:
            sender.setText("✔ 완료")

        self.lbl_last.setText(f"마지막 복사: {text}")
# ----------------------------------------------------------------------
# 메모기능
# ----------------------------------------------------------------------
class MemoDialog(QDialog):
    def __init__(self, current_text: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("메모 작성")
        self.resize(400, 300)

        layout = QVBoxLayout(self)

        self.txt_memo = QPlainTextEdit()
        self.txt_memo.setPlaceholderText("여기에 메모를 입력하세요...")
        self.txt_memo.setPlainText(current_text)
        layout.addWidget(self.txt_memo)

        btn_layout = QHBoxLayout()
        btn_save = QPushButton("저장")
        btn_close = QPushButton("닫기")

        btn_save.clicked.connect(self.accept)  # accept -> 결과 OK
        btn_close.clicked.connect(self.reject)  # reject -> 취소

        # 저장 버튼 스타일 (파란색)
        btn_save.setStyleSheet("background-color: #4dabf7; color: white; font-weight: bold;")

        btn_layout.addStretch(1)
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_close)

        layout.addLayout(btn_layout)

    def get_text(self):
        return self.txt_memo.toPlainText()
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
# [추가] 색상 혼합을 위한 전용 페인트공 (Delegate)
# ----------------------------------------------------------------------
class BlendDelegate(QtWidgets.QStyledItemDelegate):
    def paint(self, painter, option, index):
        # 1. 먼저 "선택되지 않은 척"하고 원래 배경(노랑/흰색)과 글자를 그립니다.
        opt = QtWidgets.QStyleOptionViewItem(option)
        opt.state &= ~QtWidgets.QStyle.State_Selected
        super().paint(painter, opt, index)

        # 2. 만약 실제로 "선택된 상태"라면, 그 위에 반투명한 파란색을 덧칠합니다.
        if option.state & QtWidgets.QStyle.State_Selected:
            painter.save()
            # 색상: 하늘색(R, G, B) / 투명도(Alpha: 100) -> 이 숫자를 조절하면 색감이 바뀜
            # 노란색(배경) + 하늘색(덧칠) = 연두색(초록빛)으로 보임
            color = QtGui.QColor(0, 120, 215, 100)
            painter.fillRect(option.rect, color)
            painter.restore()

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

        self._image_map: Dict[int, List[str]] = {}
        self._image_dir: Optional[Path] = None
        self._meta_path: Optional[Path] = None

        self._col_index: Dict[str, int] = {}
        self._current_row_idx: Optional[int] = None

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(12)

        # -------------------------
        # 상단 컨트롤 바 (ship_top)
        # -------------------------
        top_wrap = QWidget(self)
        top_wrap.setObjectName("ship_top")
        top_layout = QHBoxLayout(top_wrap)
        top_layout.setContentsMargins(12, 10, 12, 10)
        top_layout.setSpacing(10)

        lbl_type = QLabel("송장 타입:")
        self.combo_type = QComboBox()
        self.combo_type.addItems(["네이버 송장", "쿠팡 송장"])
        self.lbl_file = QLabel("선택된 파일: (없음)")
        self.btn_open = QPushButton("엑셀 불러오기")

        top_layout.addWidget(lbl_type)
        top_layout.addWidget(self.combo_type)
        top_layout.addSpacing(14)
        top_layout.addWidget(self.lbl_file, 1)
        top_layout.addSpacing(14)
        top_layout.addWidget(self.btn_open)

        main_layout.addWidget(top_wrap)

        # -------------------------
        # 테이블 (ship_table)
        # -------------------------
        self.table = QTableWidget()
        self.table.setObjectName("ship_table")

        self.table.setItemDelegate(BlendDelegate(self.table))
        
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(True)
        main_layout.addWidget(self.table, 1)

        self.btn_complete = QPushButton("▼ 선택된 주문 '작업 완료' 체크 (노란색 표시)")
        self.btn_complete.setFixedHeight(30)
        # 배경색(#FFB6C1 = 연한 핑크), 글자색(black), 굵게
        self.btn_complete.setStyleSheet("""
                    QPushButton {
                        font-weight: bold;
                        background-color: #FFB6C1; 
                        border: 1px solid #ff9eb0;
                        border-radius: 4px;
                    }
                    QPushButton:hover {
                        background-color: #ff99aa;
                    }
                """)
        self.btn_complete.clicked.connect(self.on_click_complete)
        main_layout.addWidget(self.btn_complete)

        # ============================================================
        # [추가됨] 메모 작성 버튼
        # ============================================================
        # self.btn_memo = QPushButton("📝 선택된 주문 '메모' 작성/수정")
        # self.btn_memo.setFixedHeight(30)
        # self.btn_memo.clicked.connect(self.on_click_memo)
        # main_layout.addWidget(self.btn_memo)
        # ============================================================

        # ============================================================
        # [추가됨] 문자 전송 패널 토글 버튼
        # ============================================================
        self.btn_toggle_sms = QPushButton("💬 문자 전송 패널 열기 (클릭)")
        self.btn_toggle_sms.setFixedHeight(30)
        self.btn_toggle_sms.clicked.connect(self.on_toggle_sms)
        main_layout.addWidget(self.btn_toggle_sms)
        # ============================================================

        # -------------------------
        # 문자 전송 패널 (ship_sms)
        # -------------------------
        self._build_sms_panel(main_layout)

        # ============================================================
        # [수정됨] 하단 로그 영역 (3분할: 로그 | 작업현황 | 시간정보)
        # ============================================================

        # 1. 가로 배치를 위한 레이아웃 생성
        bottom_log_layout = QHBoxLayout()
        bottom_log_layout.setSpacing(10)

        # [1구역: 왼쪽] 기존 작업 로그 (self.log)
        self.log = QPlainTextEdit()
        self.log.setObjectName("ship_log")
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("▶ [1] 작업 로그 기록")
        self.log.setFixedHeight(150)

        # [2구역: 가운데] 작업 현황 (1, 2, 3번 항목)
        self.info_dash_counts = QPlainTextEdit()
        self.info_dash_counts.setReadOnly(True)
        self.info_dash_counts.setPlaceholderText("▶ [2] 작업/각인 갯수 현황")
        self.info_dash_counts.setFixedHeight(150)

        # [3구역: 오른쪽] 시간 정보 (4, 5번 항목)
        self.info_dash_time = QPlainTextEdit()
        self.info_dash_time.setReadOnly(True)
        self.info_dash_time.setPlaceholderText("▶ [3] 시작/소요 시간")
        self.info_dash_time.setFixedHeight(150)

        # 레이아웃에 추가 (비율 1:1:1)
        bottom_log_layout.addWidget(self.log, 1)
        bottom_log_layout.addWidget(self.info_dash_counts, 1)
        bottom_log_layout.addWidget(self.info_dash_time, 1)

        # 전체 레이아웃에 하단 영역 추가
        main_layout.addLayout(bottom_log_layout)

        # 시그널
        self.btn_open.clicked.connect(self.on_click_open)
        self.table.itemDoubleClicked.connect(self.on_item_double_clicked)
        self.table.itemSelectionChanged.connect(self.on_table_selection_changed)

        # -------------------------
        # [추가] 대시보드용 변수 및 타이머
        # -------------------------
        self._start_time = None
        self._dashboard_timer = QTimer(self)
        self._dashboard_timer.setInterval(1000)

        self._dashboard_timer.timeout.connect(self._update_dashboard_timer)

        self._target_fonts = ["영문필기체", "영문바탕채", "한문바탕체", "한글바탕체"]

    # ------------------------------------------------------------------
    # 문자 전송 UI
    # ------------------------------------------------------------------
    def _build_sms_panel(self, parent_layout: QVBoxLayout):
        self.grp_sms = QGroupBox("문자 전송")
        self.grp_sms.setObjectName("ship_sms")
        self.grp_sms.setMaximumHeight(220)

        # [핵심] 기본적으로 숨김 처리 (엑셀 넓게 보기 위해)
        self.grp_sms.setVisible(False)

        layout = QHBoxLayout(self.grp_sms)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        grp = self.grp_sms

        # 왼쪽: 선택 고객 정보
        left_box = QVBoxLayout()
        left_box.setSpacing(6)
        self.lbl_sms_target = QLabel("선택 고객: (없음)")
        self.lbl_sms_count = QLabel("문구개수: -")
        left_box.addWidget(self.lbl_sms_target)
        left_box.addWidget(self.lbl_sms_count)
        left_box.addStretch(1)
        layout.addLayout(left_box, 3)

        # 가운데: 예약 시간 + 버튼들
        mid_box = QVBoxLayout()
        mid_box.setSpacing(10)

        time_box = QHBoxLayout()
        time_box.setSpacing(8)
        lbl_time = QLabel("보내는 시간:")
        self.dt_send = QDateTimeEdit(QDateTime.currentDateTime())
        self.dt_send.setCalendarPopup(True)
        self.chk_send_now = QCheckBox("지금 보내기")
        self.chk_send_now.setChecked(True)
        self.chk_send_now.toggled.connect(self._on_send_now_toggled)

        time_box.addWidget(lbl_time)
        time_box.addWidget(self.dt_send, 1)
        time_box.addWidget(self.chk_send_now)
        mid_box.addLayout(time_box)

        btn_box = QHBoxLayout()
        btn_box.setSpacing(10)
        self.btn_send_selected = QPushButton("선택 고객에게 보내기")
        self.btn_send_all = QPushButton("표시된 전체 고객에게 일괄 보내기")
        btn_box.addWidget(self.btn_send_selected)
        btn_box.addWidget(self.btn_send_all)
        mid_box.addLayout(btn_box)

        mid_box.addStretch(1)
        layout.addLayout(mid_box, 5)

        # 오른쪽: 대표 이미지 미리보기
        right_box = QVBoxLayout()
        right_box.setSpacing(8)
        self.lbl_img_preview = QLabel("대표 이미지\n미리보기 없음")
        self.lbl_img_preview.setFrameShape(QLabel.Box)
        self.lbl_img_preview.setAlignment(Qt.AlignCenter)
        self.lbl_img_preview.setFixedSize(150, 150)

        self.lbl_img_name = QLabel("")
        right_box.addWidget(self.lbl_img_preview)
        right_box.addWidget(self.lbl_img_name)
        right_box.addStretch(1)
        layout.addLayout(right_box, 3)

        parent_layout.addWidget(grp)

        self.btn_send_selected.clicked.connect(self.on_send_selected)
        self.btn_send_all.clicked.connect(self.on_send_all)

    def on_toggle_sms(self):
        # 현재 보이는 상태인지 확인
        is_visible = self.grp_sms.isVisible()

        # 반대로 설정 (보이면 숨기고, 숨겨져 있으면 보이고)
        self.grp_sms.setVisible(not is_visible)

        # 버튼 글자도 상태에 맞춰 변경
        if not is_visible:
            self.btn_toggle_sms.setText("💬 문자 전송 패널 닫기 (숨기기)")
        else:
            self.btn_toggle_sms.setText("💬 문자 전송 패널 열기 (클릭)")

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

            # 1. 시작 시간 기록
            self._start_time = QDateTime.currentDateTime()

            # 2. 타이머 시작
            self._dashboard_timer.start()

            # 3. 화면 즉시 갱신 (두 함수 모두 호출)
            self._update_dashboard_counts()  # 가운데 창 (갯수+시작시간)
            self._update_dashboard_timer()  # 오른쪽 창 (소요시간)

        except (OSError, IOError, ValueError) as e:
            QtWidgets.QMessageBox.critical(self, "엑셀 읽기 오류", str(e))
            self.log.appendPlainText(f"[오류] 엑셀을 읽는 중 문제가 발생했습니다: {e}")

    def on_item_double_clicked(self, item: QTableWidgetItem):
        if self.current_df is None:
            return

        visual_row = item.row()
        col = item.column()

        # [★핵심] 진짜 인덱스 꺼내오기
        real_row = self.table.item(visual_row, 0).data(Qt.UserRole)

        header = self.table.horizontalHeaderItem(col)
        col_name = header.text() if header is not None else ""

        # 1. 메모 수정
        if col_name == "메모":
            # 화면갱신용(visual)과 저장용(real) 둘 다 넘깁니다.
            self._open_memo_dialog(visual_row, real_row)
            return

        # 2. 품목명/상품명 수정
        if col_name in ("품목명", "상품명"):
            cell_text = item.text()
            invoice_type = self.combo_type.currentText()

            def save_modified_text(new_full_text: str):
                try:
                    # 화면 업데이트 (보이는 곳)
                    item.setText(new_full_text)

                    # 데이터프레임 업데이트 (진짜 위치)
                    self.current_df.at[real_row, col_name] = new_full_text

                    # 파일 저장
                    if self.current_file:
                        self.current_df.to_excel(self.current_file, index=False, engine="openpyxl")
                        self.log.appendPlainText(f"[수정] 행 {real_row + 1} 문구 업데이트.")
                except Exception as e:
                    QtWidgets.QMessageBox.critical(self, "오류", str(e))

            dlg = CopyLinesDialog(cell_text, invoice_type, save_modified_text, self)
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

    def on_click_complete(self):
        if self.current_df is None or not self.current_file:
            QtWidgets.QMessageBox.warning(self, "경고", "열린 파일이 없습니다.")
            return

        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QtWidgets.QMessageBox.information(self, "알림", "완료 처리할 행을 선택해주세요.")
            return

        if "작업유무" not in self.current_df.columns:
            self.current_df["작업유무"] = ""

        done_color = QtGui.QColor(255, 255, 100)
        white_color = QtGui.QColor(255, 255, 255)

        # [★핵심] 첫 번째 선택된 줄의 '진짜 번호'를 확인해서 토글 상태 결정
        first_visual_row = selected_rows[0].row()
        first_real_idx = self.table.item(first_visual_row, 0).data(Qt.UserRole)

        current_status = str(self.current_df.at[first_real_idx, "작업유무"]).strip()
        is_already_done = (current_status == "완료")

        target_color = white_color if is_already_done else done_color
        target_text = "" if is_already_done else "완료"
        status_msg = "취소" if is_already_done else "완료"

        try:
            for idx in selected_rows:
                visual_r = idx.row()
                # [★핵심] 화면상 줄번호(visual_r)로 아이템을 찾고, 그 안의 진짜 번호(real_r)를 꺼냄
                real_r = self.table.item(visual_r, 0).data(Qt.UserRole)

                # (1) 데이터프레임 저장 -> 진짜 번호(real_r) 사용
                self.current_df.at[real_r, "작업유무"] = target_text

                # (2) 화면 색칠 -> 보이는 번호(visual_r) 사용 (눈에 보이는 걸 바꿔야 하니까)
                for c in range(self.table.columnCount()):
                    it = self.table.item(visual_r, c)
                    if it:
                        it.setBackground(target_color)

                # 사진 버튼 색상 변경
                photo_col_idx = self._col_index.get("사진")
                if photo_col_idx:
                    widget = self.table.cellWidget(visual_r, photo_col_idx)
                    if widget:
                        bg_style = "background-color: #ffffaa;" if not is_already_done else ""
                        widget.setStyleSheet(bg_style)

            # 엑셀 저장
            self.current_df.to_excel(self.current_file, index=False, engine="openpyxl")
            self.log.appendPlainText(f"[저장] {len(selected_rows)}건 {status_msg} 처리 완료.")

        except PermissionError:
            QtWidgets.QMessageBox.critical(self, "저장 실패", "엑셀 파일이 열려있습니다.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "오류", f"저장 중 오류: {e}")

        self.table.clearSelection()
        self._update_dashboard_counts()

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
        df = pd.read_excel(file_path)

        # [추가] 빈 칸(NaN)을 숫자로 착각하지 않도록, 문자열(object)로 강제 변환
        # 이렇게 하면 '메모'나 '작업유무'에 글자를 넣어도 경고가 안 뜹니다.
        for col in ["작업유무", "메모"]:
            if col in df.columns:
                df[col] = df[col].astype(object)

        return df

    @staticmethod
    def _load_coupang_invoice(file_path: str) -> pd.DataFrame:
        df = pd.read_excel(file_path)

        # [추가] 쿠팡도 똑같이 처리
        for col in ["작업유무", "메모"]:
            if col in df.columns:
                df[col] = df[col].astype(object)

        return df

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
        # 1. 데이터 채우는 동안 정렬 기능 잠시 끄기 (안전장치)
        self.table.setSortingEnabled(False)
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

        done_bg = QtGui.QColor(255, 255, 100)  # 노란색

        for row_idx in range(row_count):
            # 작업유무 확인
            is_done = False
            if "작업유무" in df.columns:
                val = str(df.iloc[row_idx]["작업유무"]).strip()
                if val == "완료":
                    is_done = True

            for col_idx, col_name in enumerate(cols):
                val = df.iloc[row_idx, col_idx]
                text = "" if pd.isna(val) else str(val)
                item = QTableWidgetItem(text)

                # [★핵심] 0번 컬럼(맨 앞칸)에 '진짜 데이터 번호(row_idx)'를 숨겨둡니다!
                # 화면이 뒤섞여도 이 값은 변하지 않습니다.
                if col_idx == 0:
                    item.setData(Qt.UserRole, row_idx)

                # 완료된 행 색칠
                if is_done:
                    item.setBackground(done_bg)

                # 툴팁 등 기존 옵션
                if col_name in ("품목명", "상품명"):
                    item.setToolTip(text)

                self.table.setItem(row_idx, col_idx, item)

            # 사진 버튼
            photo_col_idx = self._col_index[photo_col_name]
            btn = QPushButton()
            row_id = row_idx + 1
            count = len(self._image_map.get(row_id, []))
            btn.setText(f"사진({count}장)…")
            # [중요] 사진 관리창도 '진짜 번호'인 row_idx를 가져가야 함 (이미 잘 되어 있음)
            btn.clicked.connect(self._make_photo_button_handler(row_idx))
            self.table.setCellWidget(row_idx, photo_col_idx, btn)

            if is_done:
                btn.setStyleSheet("background-color: #ffffaa;")

        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()

        # 2. 다 채웠으니 정렬 기능 켜기 (이제 헤더 클릭 가능!)
        self.table.setSortingEnabled(True)

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


    def on_click_memo(self):
        """메모 버튼 클릭 시 실행"""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QtWidgets.QMessageBox.information(self, "알림", "메모를 작성할 행을 선택해주세요.")
            return

        # 첫 번째 선택된 행에 대해서만 메모 창을 엽니다.
        row_idx = selected_rows[0].row()
        self._open_memo_dialog(row_idx)

    # [수정] 인자를 (visual_row, real_row) 두 개 받도록 변경
    def _open_memo_dialog(self, visual_row: int, real_row: int):
        if self.current_df is None or not self.current_file:
            return

        # 1. 현재 메모 내용 가져오기 (화면 기준)
        col_idx = self._col_index.get("메모")
        if col_idx is None:
            QtWidgets.QMessageBox.warning(self, "오류", "'메모' 컬럼이 없습니다.")
            return

        current_item = self.table.item(visual_row, col_idx)
        current_text = current_item.text() if current_item else ""

        # 2. 다이얼로그 띄우기
        dlg = MemoDialog(current_text, self)
        if dlg.exec_() == QDialog.Accepted:
            new_text = dlg.get_text()

            try:
                # (1) 화면 업데이트 (visual_row 사용)
                if current_item:
                    current_item.setText(new_text)
                else:
                    self.table.setItem(visual_row, col_idx, QTableWidgetItem(new_text))

                # (2) DataFrame 업데이트 (real_row 사용 - ★여기가 핵심)
                if "메모" not in self.current_df.columns:
                    self.current_df.insert(0, "메모", "")

                self.current_df.at[real_row, "메모"] = new_text

                # (3) 엑셀 파일 저장
                self.current_df.to_excel(self.current_file, index=False, engine="openpyxl")
                self.log.appendPlainText(f"[메모 저장] 행 {real_row + 1}: {new_text}")

            except PermissionError:
                QtWidgets.QMessageBox.critical(self, "저장 실패", "엑셀이 열려있습니다.")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "오류", f"저장 중 오류: {e}")

    

    def _update_dashboard_counts(self):
        """
        [가운데 창] 업데이트:
        1. 전체 작업 (완료/전체)
        2. 각인 갯수 (완료/전체)
        3. 시작 시간
        """
        if self.current_df is None or self._start_time is None:
            self.info_dash_counts.setPlainText("▶ 파일 로드 대기 중...")
            return

        total_rows = len(self.current_df)

        # 타겟 컬럼 찾기
        target_col = None
        for col in ["품목명", "상품명"]:
            if col in self.current_df.columns:
                target_col = col
                break

        # --- 카운팅 변수 ---
        completed_rows = 0  # 1번: 완료된 행 갯수

        total_gagin_count = 0  # 2번: 전체 각인 글자 수
        completed_gagin_count = 0  # 2번: 완료된 각인 글자 수

        if target_col:
            if "작업유무" not in self.current_df.columns:
                self.current_df["작업유무"] = ""

            for i in range(total_rows):
                item_name = str(self.current_df.iloc[i][target_col])
                status = str(self.current_df.iloc[i]["작업유무"]).strip()
                is_done = (status == "완료")

                # 1. 완료된 행 카운트
                if is_done:
                    completed_rows += 1

                # 2. 각인 글자 수 카운트 (행별 합산)
                row_gagin = 0
                for kw in self._target_fonts:
                    row_gagin += item_name.count(kw)

                total_gagin_count += row_gagin
                if is_done:
                    completed_gagin_count += row_gagin

        # 3. 시작 시간
        start_time_str = self._start_time.toString("yyyy-MM-dd HH:mm:ss")

        # --- 화면 출력 포맷 ---
        # (완료 / 전체) 형식으로 통일하여 직관적으로 표시
        text_counts = (
            f"📊 [작업 현황]\n"
            f"━━━━━━━━━━━━━━━━━━\n"
            f"1. 전체 작업 갯수 :  {completed_rows} / {total_rows} 건\n"
            f"   (완료 / 전체 행)\n\n"
            f"2. 각인 갯수 현황 :  {completed_gagin_count} / {total_gagin_count} 개\n"
            f"   (완료 / 전체 키워드)\n\n"
            f"3. 시작 시간 :\n"
            f"   {start_time_str}"
        )
        self.info_dash_counts.setPlainText(text_counts)

    def _update_dashboard_timer(self):
        """
        [오른쪽 창] 업데이트: 5번 항목 (실시간 소요 시간)
        - 이 함수만 1초마다 실행됩니다.
        """
        if self._start_time is None:
            self.info_dash_time.setPlainText("-")
            return

        now = QDateTime.currentDateTime()
        seconds_diff = self._start_time.secsTo(now)

        hours = seconds_diff // 3600
        minutes = (seconds_diff % 3600) // 60
        seconds = seconds_diff % 60

        time_elapsed_str = f"{minutes}분 {seconds}초"
        if hours > 0:
            time_elapsed_str = f"{hours}시간 " + time_elapsed_str

        text_time = (
            f"⏰ [실시간 소요]\n"
            f"━━━━━━━━━━━━━━━━━━\n"
            f"5. 현재 소요 시간 :\n\n"
            f"   {time_elapsed_str}"
        )
        self.info_dash_time.setPlainText(text_time)


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
