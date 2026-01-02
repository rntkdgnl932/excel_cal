# read_excel.py
# 네이버·쿠팡 송장 엑셀을 읽어와서 보여주고,
# 품목명 파싱, 복사용 문구 COPY, 사진 첨부/삭제/재사용,
# 문자 전송 UI 뼈대 및 검색/폰트조절/취소선 기능을 포함한 탭 위젯.

import os
import re
import json
import shutil
from pathlib import Path
from typing import Optional, List, Dict
import subprocess
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
    QSpinBox,
)
# ----------------------------------------------------------------------
# 캐시 무시하고 이미지 불러오는 헬퍼 함수 추가
# ----------------------------------------------------------------------
def load_pixmap_clean(path: str) -> QtGui.QPixmap:
    """
    파일 경로(path)에서 바이너리를 직접 읽어 QPixmap을 생성합니다.
    Qt의 내부 캐싱(같은 파일명일 때 이전 이미지 보여주는 문제)을 방지합니다.
    """
    p = Path(path)
    if not p.is_file():
        return QtGui.QPixmap()

    with open(p, "rb") as f:
        data = f.read()

    pix = QtGui.QPixmap()
    pix.loadFromData(data)
    return pix
# ----------------------------------------------------------------------
# 전체 문구 수정용 다이얼로그 클래스
# ----------------------------------------------------------------------
class FullTextEditDialog(QDialog):
    """전체 문구를 통째로 수정하는 다이얼로그"""

    def __init__(self, text: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("전체 문구 수정")
        self.resize(500, 600)

        layout = QVBoxLayout(self)

        # 안내 문구
        lbl_info = QLabel("문구 전체를 수정하세요. [완료]를 누르면 리스트가 갱신됩니다.")
        layout.addWidget(lbl_info)

        # 텍스트 에디트 (여러 줄 입력 가능)
        self.txt_edit = QPlainTextEdit()
        self.txt_edit.setPlainText(text)
        layout.addWidget(self.txt_edit)

        # 버튼
        btn_layout = QHBoxLayout()
        btn_save = QPushButton("완료")
        btn_cancel = QPushButton("취소")

        btn_save.setFixedHeight(40)
        btn_save.setStyleSheet("background-color: #4dabf7; color: white; font-weight: bold;")

        btn_save.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch(1)
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_cancel)

        layout.addLayout(btn_layout)

    def get_text(self):
        return self.txt_edit.toPlainText()

# ----------------------------------------------------------------------
# COPY 다이얼로그: 파싱된 문구들 + 각 줄별 COPY 버튼
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# COPY 다이얼로그: 파싱된 문구들 + 각 줄별 COPY 버튼
# ----------------------------------------------------------------------
class CopyLinesDialog(QDialog):
    def __init__(self, full_text: str, invoice_type: str, save_callback, parent=None):
        super().__init__(parent)
        self.setWindowTitle("문구 복사 & 수정")
        self.resize(650, 600)

        self.save_callback = save_callback  # 엑셀 저장 콜백
        self.invoice_type = invoice_type  # 송장 타입 (네이버/쿠팡)
        self.raw_lines = full_text.splitlines()  # 원본 줄 데이터

        # 메인 레이아웃
        self.main_layout = QVBoxLayout(self)

        # 1. 상단 영역 (안내문구 + 전체수정 버튼)
        top_layout = QHBoxLayout()

        info_layout = QVBoxLayout()
        info = QLabel("개별 [수정] 또는 우측 [전체 수정]을 이용하세요.")
        self.lbl_last = QLabel("마지막 복사: 없음")
        self.lbl_last.setStyleSheet("color: blue; font-weight: bold;")
        info_layout.addWidget(info)
        info_layout.addWidget(self.lbl_last)

        # [NEW] 전체 수정 버튼
        self.btn_edit_all = QPushButton("전체 문구 수정")
        self.btn_edit_all.setFixedSize(120, 40)
        self.btn_edit_all.setStyleSheet("background-color: #69db7c; color: white; font-weight: bold;")
        self.btn_edit_all.clicked.connect(self.on_click_edit_all)

        top_layout.addLayout(info_layout)
        top_layout.addStretch(1)
        top_layout.addWidget(self.btn_edit_all)

        self.main_layout.addLayout(top_layout)

        # 2. 스크롤 영역 (리스트가 들어갈 곳)
        self.scroll_area = QtWidgets.QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QtWidgets.QFrame.NoFrame)

        # 컨텐츠 위젯과 레이아웃 초기화
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.scroll_area.setWidget(self.content_widget)

        self.main_layout.addWidget(self.scroll_area, 1)

        # 3. 하단 닫기 버튼
        btn_close = QPushButton("닫기")
        btn_close.clicked.connect(self.accept)
        self.main_layout.addWidget(btn_close)

        # 4. 초기 리스트 그리기
        self._refresh_list_ui()

    def _refresh_list_ui(self):
        """현재 self.raw_lines를 기반으로 화면을 다시 그립니다. (폰트 콤보박스 추가됨)"""
        # 기존 아이템들 삭제 (레이아웃 청소)
        while self.content_layout.count():
            child = self.content_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
            elif child.layout():
                import sip
                sip.delete(child)

        self.parsed_items = []
        self._ui_rows = []

        # 폰트 옵션 정의 (오타 수정됨)
        font_options = ["확인불가", "영문필기체", "영문바탕체", "한문바탕체", "한글바탕체"]

        # 다시 파싱
        for idx, line in enumerate(self.raw_lines):
            parts = self._parse_structure(line, self.invoice_type)
            if parts:
                item_data = {
                    "line_idx": idx,
                    "prefix": parts[0],
                    "core": parts[1],
                    "suffix": parts[2]
                }
                self.parsed_items.append(item_data)

                # --- UI Row 생성 ---
                row_layout = QHBoxLayout()

                # [1] 배지 생성 (원본 줄 전체 넘김)
                badges = self._create_badges(line)
                for badge in badges:
                    row_layout.addWidget(badge)

                # [2] 텍스트 에디트 (문구만 표시)
                edit = QLineEdit()
                edit.setText(item_data["core"])
                edit.setReadOnly(True)
                edit.setStyleSheet("background-color: #f0f0f0; color: #333;")

                # [3] 폰트 선택 콤보박스 (NEW)
                combo_font = QComboBox()
                combo_font.addItems(font_options)
                combo_font.setFixedWidth(100)
                combo_font.setEnabled(False)  # 기본은 비활성 (수정 버튼 눌러야 활성)

                # 현재 줄에 어떤 폰트가 있는지 확인해서 디폴트 값 설정
                current_font = "확인불가"
                for f_opt in font_options:
                    if f_opt != "확인불가" and f_opt in line:
                        current_font = f_opt
                        break
                combo_font.setCurrentText(current_font)

                # 콤보박스 스타일 (비활성일 때 회색조)
                combo_font.setStyleSheet("QComboBox:!enabled { background-color: #f0f0f0; color: #888; }")

                # [4] 버튼들
                btn_edit = QPushButton("수정")
                btn_edit.setFixedWidth(50)

                btn_copy = QPushButton("COPY")
                btn_copy.setFixedWidth(60)

                # 핸들러 연결 (콤보박스도 같이 넘김)
                self._connect_handlers(edit, combo_font, btn_edit, btn_copy, item_data)

                row_layout.addWidget(edit, 1)
                row_layout.addWidget(combo_font)  # 에디트 옆에 추가
                row_layout.addWidget(btn_edit)
                row_layout.addWidget(btn_copy)
                self.content_layout.addLayout(row_layout)

                self._ui_rows.append({
                    "edit": edit,
                    "combo": combo_font,
                    "btn_edit": btn_edit,
                    "btn_copy": btn_copy
                })

        self.content_layout.addStretch(1)

    def _create_badges(self, full_text: str) -> List[QLabel]:
        """
        원본 줄(full_text)에서 키워드(M1, S, 폰트명)를 찾아 배지(라벨) 리스트로 반환.
        폰트가 없으면 '확인불가' 배지를 추가함.
        """
        badges = []

        # 1. 유형 코드 (M1, S, W2 등)
        type_keywords = {

            "M1)": ("#ffcccc", "#d6336c"),  # 핑크

            "S)": ("#cce5ff", "#004085"),  # 블루

            "W2)": ("#e2e3e5", "#383d41"),  # 그레이

        }

        # 2. 폰트 코드 (오타 포함해서 다양하게 감지)
        # 키(검색어): (배경색, 글자색, 화면표시텍스트)
        font_keywords = {
            "영문필기체": ("#fff3cd", "#856404", "영문필기체"),
            "영문 필기체": ("#fff3cd", "#856404", "영문필기체"),
            "영문바탕채": ("#d4edda", "#155724", "영문바탕체"),  # 오타 보정
            "영문바탕체": ("#d4edda", "#155724", "영문바탕체"),
            "영문 바탕체": ("#d4edda", "#155724", "영문바탕체"),
            "한문바탕체": ("#d1ecf1", "#0c5460", "한문바탕체"),
            "한문 바탕체": ("#d1ecf1", "#0c5460", "한문바탕체"),
            "한글바탕체": ("#f8d7da", "#721c24", "한글바탕체"),
            "한글 바탕체": ("#f8d7da", "#721c24", "한글바탕체"),
        }

        # (1) 유형 코드 배지 생성
        for key, colors in type_keywords.items():
            # M1, S 등은 괄호나 공백과 붙어있을 수 있으므로 단순 포함 여부 확인
            if key in full_text:
                lbl = QLabel(key)
                lbl.setStyleSheet(
                    f"background-color: {colors[0]}; color: {colors[1]}; "
                    f"border-radius: 4px; padding: 2px 6px; font-weight: bold; font-family: 맑은 고딕;"
                )
                badges.append(lbl)

        # (2) 폰트 코드 배지 생성
        font_found = False
        for key, val in font_keywords.items():
            bg_color, txt_color, display_text = val
            if key in full_text:
                font_found = True
                lbl = QLabel(display_text)
                lbl.setStyleSheet(
                    f"background-color: {bg_color}; color: {txt_color}; "
                    f"border-radius: 4px; padding: 2px 6px; font-weight: bold; font-family: 맑은 고딕;"
                )
                badges.append(lbl)
                break  # 폰트는 한 줄에 하나만 나올 테니 찾으면 중단 (중복 방지)

        # (3) 폰트를 못 찾았을 경우 -> '확인불가' 배지
        if not font_found:
            lbl = QLabel("확인불가")
            lbl.setStyleSheet(
                "background-color: #343a40; color: #ffffff; "  # 진한 검회색 배경, 흰 글씨
                "border-radius: 4px; padding: 2px 6px; font-weight: bold; font-family: 맑은 고딕;"
            )
            badges.append(lbl)

        return badges

    def on_click_edit_all(self):
        """[전체 문구 수정] 버튼 클릭 핸들러"""
        # 현재 줄들을 합쳐서 전체 텍스트로 만듦
        full_text = "\n".join(self.raw_lines)

        # 새 다이얼로그 띄우기
        dlg = FullTextEditDialog(full_text, self)
        if dlg.exec_() == QDialog.Accepted:
            new_text = dlg.get_text()

            # 1. 엑셀 및 데이터 저장
            self.save_callback(new_text)

            # 2. 내부 데이터 업데이트
            self.raw_lines = new_text.splitlines()

            # 3. 리스트 UI 새로고침 (재파싱)
            self._refresh_list_ui()

            # 로그/알림
            self.lbl_last.setText("마지막 작업: 전체 수정 완료")

    def _parse_structure(self, line: str, invoice_type: str):
        """(기존 로직 유지) 한 줄을 (접두어, 핵심문구, 접미어)로 분리"""
        s = line.strip()
        if not re.match(r"^\d+\.", s):
            return None

        temp_core = ""
        if invoice_type == "네이버 송장":
            body = s.split(".", 1)[1].lstrip()
            idx = body.find("/ 각인체")
            if idx != -1:
                temp_core = body[:idx].strip()
            elif "=>" in body:
                temp_core = body.split("=>", 1)[0].strip()
            else:
                temp_core = body.strip()
        else:
            body = s
            if ":" in body:
                body = body.split(":", 1)[1]
            if "=>" in body:
                temp_core = body.split("=>", 1)[0].strip()
            else:
                temp_core = body.strip()

        if not temp_core:
            return None

        start_idx = line.find(temp_core)
        if start_idx == -1:
            return None

        prefix = line[:start_idx]
        suffix = line[start_idx + len(temp_core):]
        return (prefix, temp_core, suffix)

    def _connect_handlers(self, edit, combo_font, btn_edit, btn_copy, item_data):
        def on_edit_click():
            if btn_edit.text() == "수정":
                # 수정 모드 진입
                edit.setReadOnly(False)
                edit.setFocus()
                edit.setStyleSheet("background-color: #ffffff; color: #000; border: 2px solid #4dabf7;")

                # 폰트 선택 활성화
                combo_font.setEnabled(True)
                combo_font.setStyleSheet("background-color: #ffffff; color: #000; border: 1px solid #4dabf7;")

                btn_edit.setText("저장")
                btn_edit.setStyleSheet("color: blue; font-weight: bold;")
                btn_copy.setEnabled(False)
            else:
                # [저장 로직]

                # 1. 텍스트 박스의 변경된 문구 가져오기
                new_core_text = edit.text()

                # 2. 콤보박스에서 선택된 폰트 가져오기
                new_font = combo_font.currentText()

                # 3. 원본 줄(raw_line) 가져오기
                original_line = self.raw_lines[item_data["line_idx"]]

                # 4. 폰트 교체 작업
                # (1) 기존에 있던 폰트 텍스트 제거 (확인불가 제외)
                font_options = ["영문필기체", "영문 필기체", "영문바탕체", "영문 바탕체", "영문바탕채",
                                "한문바탕체", "한문 바탕체", "한글바탕체", "한글 바탕체"]

                temp_line = original_line
                for f in font_options:
                    temp_line = temp_line.replace(f, "")  # 기존 폰트 삭제

                # (2) 문구(Core) 부분 교체
                # 기존 Core를 찾아서 새 Core로 바꿈 (단순 replace는 위험할 수 있으니 구조 기반 재조립 권장)
                # 여기서는 안전하게 prefix + new_core + suffix 구조를 사용하려 했으나,
                # 폰트 위치가 prefix/suffix 어디에 있을지 모르므로 '=>' 기준으로 삽입합니다.

                # 일단 Core 텍스트 업데이트 (item_data 구조상)
                # 하지만 temp_line은 이미 폰트가 빠진 상태의 전체 줄임.
                # 여기서 Core 부분을 교체해줘야 함.
                if item_data["core"] in temp_line:
                    temp_line = temp_line.replace(item_data["core"], new_core_text)

                # (3) 새 폰트 삽입 ("확인불가"가 아니면 삽입)
                if new_font != "확인불가":
                    # '=>' 가 있으면 그 앞에 폰트 삽입
                    if "=>" in temp_line:
                        parts = temp_line.split("=>", 1)  # 첫번째 => 기준 분리
                        # 앞부분 + 공백 + 새 폰트 + => + 뒷부분
                        temp_line = f"{parts[0].strip()} {new_font}=>{parts[1]}"
                    else:
                        # 혹시 '=>'가 없으면(드물지만) 맨 뒤에 붙임
                        temp_line = f"{temp_line} {new_font}"

                # 5. 최종 데이터 업데이트
                item_data["core"] = new_core_text  # 내부 데이터 갱신
                self.raw_lines[item_data["line_idx"]] = temp_line  # 전체 줄 갱신

                # 전체 텍스트 합쳐서 콜백으로 저장 요청 (엑셀 저장)
                new_full_text = "\n".join(self.raw_lines)
                self.save_callback(new_full_text)

                # 6. UI 복귀 (읽기 전용)
                edit.setReadOnly(True)
                edit.setStyleSheet("background-color: #f0f0f0; color: #333;")

                combo_font.setEnabled(False)
                combo_font.setStyleSheet("QComboBox:!enabled { background-color: #f0f0f0; color: #888; }")

                btn_edit.setText("수정")
                btn_edit.setStyleSheet("")
                btn_copy.setEnabled(True)

                # 리스트 새로고침 (배지 업데이트 등을 위해)
                self._refresh_list_ui()

        def on_copy_click():
            text = edit.text()
            QApplication.clipboard().setText(text)
            self._mark_copied(text)

        btn_edit.clicked.connect(on_edit_click)
        btn_copy.clicked.connect(on_copy_click)

    def _mark_copied(self, text: str):
        for row in self._ui_rows:
            row["btn_copy"].setText("COPY")
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
# [업데이트] 이미지 관리 다이얼로그 (ADB로 폰 사진 가져오기 추가)
# ----------------------------------------------------------------------
class ImageManageDialog(QDialog):
    def __init__(self, parent, row_id: int, image_dir: Path, current_files: List[str]):
        super().__init__(parent)
        self.setWindowTitle(f"사진 관리 - 행 {row_id}")
        self.resize(600, 450)

        self.row_id = row_id
        self.image_dir = image_dir
        self._images: List[str] = list(current_files)

        main_layout = QVBoxLayout(self)

        # 상단 리스트 + 미리보기 영역
        top_layout = QHBoxLayout()
        main_layout.addLayout(top_layout)

        self.list_widget = QListWidget()
        self.list_widget.currentRowChanged.connect(self._on_list_selection_changed)
        top_layout.addWidget(self.list_widget, 2)

        right_layout = QVBoxLayout()
        top_layout.addLayout(right_layout, 3)

        self.lbl_preview = QLabel("미리보기 없음")
        self.lbl_preview.setFrameShape(QLabel.Box)
        self.lbl_preview.setAlignment(Qt.AlignCenter)
        self.lbl_preview.setFixedSize(260, 260)
        right_layout.addWidget(self.lbl_preview)

        self.lbl_filename = QLabel("")
        self.lbl_filename.setWordWrap(True)
        right_layout.addWidget(self.lbl_filename)

        # ---------------------------------------------------------
        # [NEW] 스마트폰 연동 버튼 구역
        # ---------------------------------------------------------
        adb_box = QGroupBox("스마트폰(ADB) 연동")
        adb_layout = QHBoxLayout(adb_box)

        self.btn_adb_latest = QPushButton("📷 방금 찍은 최신 사진 가져오기")
        self.btn_adb_latest.setStyleSheet("background-color: #fff0f6; color: #d6336c; font-weight: bold;")
        self.btn_adb_latest.clicked.connect(self._on_adb_import_latest)

        adb_layout.addWidget(self.btn_adb_latest)
        main_layout.addWidget(adb_box)

        # ---------------------------------------------------------
        # 기본 버튼 구역
        # ---------------------------------------------------------
        btn_layout = QHBoxLayout()
        self.btn_add = QPushButton("+ PC에서 파일찾기")
        self.btn_primary = QPushButton("대표로 지정")
        self.btn_del = QPushButton("삭제")
        self.btn_close = QPushButton("닫기")

        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_primary)
        btn_layout.addWidget(self.btn_del)
        btn_layout.addStretch(1)
        btn_layout.addWidget(self.btn_close)

        main_layout.addLayout(btn_layout)

        self.btn_add.clicked.connect(self._on_add)
        self.btn_primary.clicked.connect(self._on_set_primary)
        self.btn_del.clicked.connect(self._on_del)
        self.btn_close.clicked.connect(self.accept)

        self._reload_list()

    def images(self) -> List[str]:
        return list(self._images)

    def _reload_list(self):
        self.list_widget.clear()
        for fname in self._images:
            item = QListWidgetItem(fname)
            fpath = self.image_dir / fname

            # [수정] 캐시 무시하고 로드
            pix = load_pixmap_clean(str(fpath))

            if not pix.isNull():
                icon = QtGui.QIcon(pix.scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                item.setIcon(icon)
            self.list_widget.addItem(item)

        if self._images:
            self.list_widget.setCurrentRow(0)
        else:
            self.lbl_preview.setText("미리보기 없음")
            self.lbl_preview.setPixmap(QtGui.QPixmap())
            self.lbl_filename.setText("")

    def _on_list_selection_changed(self, row: int):
        if row < 0 or row >= len(self._images):
            self.lbl_preview.setText("미리보기 없음")
            self.lbl_preview.setPixmap(QtGui.QPixmap())
            self.lbl_filename.setText("")
            return

        fname = self._images[row]
        fpath = self.image_dir / fname

        # [수정] 캐시 무시하고 로드
        pix = load_pixmap_clean(str(fpath))

        if pix.isNull():
            self.lbl_preview.setText("파일 읽기 실패")
            self.lbl_filename.setText(fname)
            return

        scaled = pix.scaled(self.lbl_preview.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.lbl_preview.setPixmap(scaled)
        self.lbl_filename.setText(fname)

    def _on_add(self):
        src_path, _ = QFileDialog.getOpenFileName(self, "이미지 선택", "", "Images (*.jpg *.png *.jpeg *.bmp)")
        if src_path:
            self._save_local_image(Path(src_path))

    def _save_local_image(self, src_path: Path):
        try:
            ext = src_path.suffix.lower()
            if not ext: ext = ".jpg"

            # 파일명 생성 (row_0001_1.jpg)
            next_idx = len(self._images) + 1
            new_name = f"row_{self.row_id:04d}_{next_idx}{ext}"

            self.image_dir.mkdir(parents=True, exist_ok=True)
            target = self.image_dir / new_name
            target.write_bytes(src_path.read_bytes())

            self._images.append(new_name)
            self._reload_list()
            # 방금 추가한 것을 선택
            self.list_widget.setCurrentRow(len(self._images) - 1)

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "오류", str(e))

    # [핵심] ADB로 최신 사진 가져오기
    def _on_adb_import_latest(self):
        import subprocess

        try:
            # 1. 폰의 카메라 폴더(DCIM/Camera)에서 날짜순 정렬(-t)하여 가장 최신 파일 1개만 조회
            # 안드로이드 표준 경로: /sdcard/DCIM/Camera/
            cmd_ls = ["adb", "shell", "ls", "-t", "/sdcard/DCIM/Camera/"]

            # process 실행
            proc = subprocess.Popen(cmd_ls, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            out, err = proc.communicate()

            if proc.returncode != 0:
                QtWidgets.QMessageBox.warning(self, "연동 실패", "폰이 연결되지 않았거나 권한이 없습니다.\nUSB 디버깅을 확인해주세요.")
                return

            # 결과 파싱 (파일명들 중 jpg/png만 필터링)
            files = out.decode("utf-8").splitlines()
            target_file = None
            for f in files:
                f = f.strip()
                if f.lower().endswith(('.jpg', '.jpeg', '.png')):
                    target_file = f
                    break

            if not target_file:
                QtWidgets.QMessageBox.information(self, "알림", "최신 사진을 찾을 수 없습니다.")
                return

            # 2. 파일 가져오기 (adb pull)
            phone_path = f"/sdcard/DCIM/Camera/{target_file}"

            # 임시 경로에 다운로드
            temp_path = self.image_dir / f"temp_{target_file}"
            self.image_dir.mkdir(parents=True, exist_ok=True)

            cmd_pull = ["adb", "pull", phone_path, str(temp_path)]
            subprocess.run(cmd_pull, check=True)

            # 3. 정식 등록 (이름 변경 및 리스트 추가)
            self._save_local_image(temp_path)

            # 임시 파일 삭제
            if temp_path.exists():
                temp_path.unlink()

            QtWidgets.QMessageBox.information(self, "성공", "휴대폰의 최신 사진을 가져왔습니다!")

        except FileNotFoundError:
            QtWidgets.QMessageBox.critical(self, "오류", "adb.exe가 없습니다.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "오류", f"사진 가져오기 실패: {e}")

    def _on_set_primary(self):
        row = self.list_widget.currentRow()
        if row <= 0: return
        fname = self._images.pop(row)
        self._images.insert(0, fname)
        self._reload_list()
        self.list_widget.setCurrentRow(0)

    def _on_del(self):
        row = self.list_widget.currentRow()
        if row < 0: return
        fname = self._images.pop(row)
        fpath = self.image_dir / fname
        try:
            if fpath.is_file(): fpath.unlink()
        except:
            pass
        self._reload_list()


# ----------------------------------------------------------------------
# [추가] 색상 혼합을 위한 전용 페인트공 (Delegate) + [개선] 취소선 기능 추가
# ----------------------------------------------------------------------
class BlendDelegate(QtWidgets.QStyledItemDelegate):
    def paint(self, painter, option, index):
        # 1. 먼저 "선택되지 않은 척"하고 원래 배경(노랑/흰색)과 글자를 그립니다.
        opt = QtWidgets.QStyleOptionViewItem(option)
        opt.state &= ~QtWidgets.QStyle.State_Selected

        # [NEW] 취소선 확인 (UserRole + 2 에 True가 있으면 취소선)
        is_strike = index.data(Qt.UserRole + 2)
        if is_strike:
            font = opt.font
            font.setStrikeOut(True)
            opt.font = font

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

        # [요청 1] 원하는 컬럼 순서 지정
        self.target_columns = [
            "주문일시",
            "특기사항",
            "받으시는 분",
            "구매자명",
            "구매자연락처",
            "품목명",
            "메모",
            "작업유무",
            "문자발송처리"
        ]

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

        # [필터] 엑셀 불러오기 버튼 앞
        self.combo_filter = QComboBox()
        self.combo_filter.addItems(["전체", "작업중", "작업완료", "문자발송완료"])
        # [핵심] 필터 변경 시 즉시 목록 갱신
        self.combo_filter.currentIndexChanged.connect(self._apply_filter)

        self.btn_open = QPushButton("엑셀 불러오기")

        top_layout.addWidget(lbl_type)
        top_layout.addWidget(self.combo_type)
        top_layout.addSpacing(14)
        top_layout.addWidget(self.lbl_file, 1)
        top_layout.addSpacing(14)

        top_layout.addWidget(self.combo_filter)
        top_layout.addSpacing(10)
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

        # -------------------------
        # 하단 기능 패널
        # -------------------------
        control_layout = QHBoxLayout()
        control_layout.setSpacing(10)

        # 1. 글자 크기
        lbl_font = QLabel("글자크기:")
        self.spin_font = QSpinBox()
        self.spin_font.setRange(8, 30)
        self.spin_font.setValue(10)
        self.spin_font.valueChanged.connect(self.on_font_size_changed)

        # 2. 검색 기능
        self.le_search = QLineEdit()
        self.le_search.setPlaceholderText("검색어 입력...")
        self.le_search.setFixedWidth(150)
        self.le_search.returnPressed.connect(self.on_click_search)

        self.btn_search = QPushButton("🔍 검색")
        self.btn_search.clicked.connect(self.on_click_search)

        # 3. 작업 완료 버튼
        self.btn_complete = QPushButton("▼ 선택된 주문 '작업 완료' 체크 (노란색)")
        self.btn_complete.setFixedHeight(30)
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

        # 4. 문자 발송 처리 버튼
        self.btn_sms_done = QPushButton("📩 문자 발송 처리 (취소선)")
        self.btn_sms_done.setFixedHeight(30)
        self.btn_sms_done.setStyleSheet("""
                    QPushButton {
                        font-weight: bold;
                        background-color: #ADD8E6; 
                        border: 1px solid #87CEEB;
                        border-radius: 4px;
                    }
                    QPushButton:hover {
                        background-color: #87CEFA;
                    }
                """)
        self.btn_sms_done.clicked.connect(self.on_click_sms_done)

        control_layout.addWidget(lbl_font)
        control_layout.addWidget(self.spin_font)
        control_layout.addStretch(1)
        control_layout.addWidget(self.le_search)
        control_layout.addWidget(self.btn_search)
        control_layout.addSpacing(10)
        control_layout.addWidget(self.btn_complete)
        control_layout.addWidget(self.btn_sms_done)

        main_layout.addLayout(control_layout)

        # 토글 버튼
        self.btn_toggle_sms = QPushButton("💬 문자 전송 패널 열기 (클릭)")
        self.btn_toggle_sms.setFixedHeight(30)
        self.btn_toggle_sms.clicked.connect(self.on_toggle_sms)
        main_layout.addWidget(self.btn_toggle_sms)

        # 문자 전송 패널
        self._build_sms_panel(main_layout)

        # 하단 로그/대시보드
        bottom_log_layout = QHBoxLayout()
        bottom_log_layout.setSpacing(10)

        self.log = QPlainTextEdit()
        self.log.setObjectName("ship_log")
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("▶ [1] 작업 로그 기록")
        self.log.setFixedHeight(150)

        self.info_dash_counts = QPlainTextEdit()
        self.info_dash_counts.setReadOnly(True)
        self.info_dash_counts.setPlaceholderText("▶ [2] 작업/각인 갯수 현황")
        self.info_dash_counts.setFixedHeight(150)

        self.info_dash_time = QPlainTextEdit()
        self.info_dash_time.setReadOnly(True)
        self.info_dash_time.setPlaceholderText("▶ [3] 시작/소요 시간")
        self.info_dash_time.setFixedHeight(150)

        bottom_log_layout.addWidget(self.log, 1)
        bottom_log_layout.addWidget(self.info_dash_counts, 1)
        bottom_log_layout.addWidget(self.info_dash_time, 1)
        main_layout.addLayout(bottom_log_layout)

        # 시그널
        self.btn_open.clicked.connect(self.on_click_open)
        self.table.itemDoubleClicked.connect(self.on_item_double_clicked)
        self.table.itemSelectionChanged.connect(self.on_table_selection_changed)

        # 대시보드 타이머
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

        self.lbl_img_preview.setCursor(Qt.PointingHandCursor)  # 마우스 올리면 손가락 모양
        self.lbl_img_preview.mousePressEvent = self._on_preview_clicked  # 클릭하면 관리창 열기

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

    # [3] 글자 크기 변경 핸들러
    def on_font_size_changed(self, val):
        self.table.setStyleSheet(f"font-size: {val}pt;")
        self.table.resizeRowsToContents()
        self.table.resizeColumnsToContents()

    # [2] 검색 핸들러
    def on_click_search(self):
        """검색 버튼 클릭 시 필터 적용"""
        self._apply_filter()

    def _apply_filter(self):
        """
        [핵심 로직] 콤보박스(필터)와 검색어(Search)를 모두 반영하여
        조건에 맞지 않는 행을 숨깁니다(Hidden).
        """
        if self.current_df is None: return

        filter_mode = self.combo_filter.currentText()
        keyword = self.le_search.text().strip().lower()

        row_count = self.table.rowCount()
        col_count = self.table.columnCount()

        for r in range(row_count):
            # 0번 컬럼에 숨겨진 real_row 인덱스 가져오기 (정렬 무관)
            item_0 = self.table.item(r, 0)
            if not item_0: continue

            real_row = item_0.data(Qt.UserRole)

            # 데이터프레임에서 상태값 조회
            work_val = str(self.current_df.at[real_row, "작업유무"]).strip()
            sms_val = str(self.current_df.at[real_row, "문자발송처리"]).strip()

            is_work_done = (work_val == "완료")
            is_sms_done = (sms_val == "완료")

            # 1. 콤보박스 필터 조건 확인
            show_by_filter = False

            if filter_mode == "전체":
                show_by_filter = True

            elif filter_mode == "작업중":
                # 작업X AND 문자X
                if not is_work_done and not is_sms_done:
                    show_by_filter = True

            elif filter_mode == "작업완료":
                # 작업O AND 문자X
                if is_work_done and not is_sms_done:
                    show_by_filter = True

            elif filter_mode == "문자발송완료":
                # 문자O (작업은 당연히 되어 있어야 함)
                if is_sms_done:
                    show_by_filter = True

            # 2. 검색어 조건 확인 (필터를 통과한 경우에만 검사)
            final_show = False
            if show_by_filter:
                if not keyword:
                    final_show = True
                else:
                    # 해당 줄의 모든 컬럼 텍스트에서 검색어 찾기
                    for c in range(col_count):
                        item = self.table.item(r, c)
                        if item and keyword in item.text().lower():
                            final_show = True
                            break

            # 3. 최종 반영
            self.table.setRowHidden(r, not final_show)

    # [4,5] 문자 발송 처리 핸들러 (취소선)
    def on_click_sms_done(self):
        """[문자 발송 처리] 버튼: 취소선 표시 + 완료 텍스트"""
        if self.current_df is None or not self.current_file: return

        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QtWidgets.QMessageBox.information(self, "알림", "처리할 행을 선택해주세요.")
            return

        if "문자발송처리" not in self.current_df.columns:
            self.current_df["문자발송처리"] = ""
        if "작업유무" not in self.current_df.columns:
            self.current_df["작업유무"] = ""

        first_vis = selected_rows[0].row()
        first_real = self.table.item(first_vis, 0).data(Qt.UserRole)
        current_val = str(self.current_df.at[first_real, "문자발송처리"]).strip()
        is_already_done = (current_val == "완료")

        target_val = "" if is_already_done else "완료"
        target_strike = False if is_already_done else True

        # [안전장치 2] 문자발송처리를 '설정(완료)'하려는 경우, 작업이 완료되었는지 확인
        if not is_already_done:  # 설정 시도
            for idx in selected_rows:
                v_r = idx.row()
                r_r = self.table.item(v_r, 0).data(Qt.UserRole)
                work_status = str(self.current_df.at[r_r, "작업유무"]).strip()
                if work_status != "완료":
                    QtWidgets.QMessageBox.warning(
                        self, "처리 불가",
                        "작업 완료하지 않은 항목이 포함되어 있습니다.\n먼저 [작업 완료] 처리를 해주세요."
                    )
                    return  # 중단

        try:
            for idx in selected_rows:
                visual_r = idx.row()
                real_r = self.table.item(visual_r, 0).data(Qt.UserRole)

                self.current_df.at[real_r, "문자발송처리"] = target_val

                # 취소선 그리기
                for c in range(self.table.columnCount()):
                    item = self.table.item(visual_r, c)
                    if item:
                        item.setData(Qt.UserRole + 2, target_strike)

            self.current_df.to_excel(self.current_file, index=False, engine="openpyxl")

            action = "취소(해제)" if is_already_done else "완료(취소선)"
            self.log.appendPlainText(f"[문자발송] {len(selected_rows)}건 {action} 처리됨.")

            self.table.repaint()

            # [핵심] 변경 즉시 필터 재적용
            self._apply_filter()

        except PermissionError:
            QtWidgets.QMessageBox.critical(self, "저장 실패", "엑셀 파일이 열려있습니다.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "오류", f"저장 중 오류: {e}")

        self.table.clearSelection()

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

        # [핵심] 새 파일을 열 때는 항상 필터를 '전체'로 초기화
        self.combo_filter.setCurrentIndex(0)

        invoice_type = self.combo_type.currentText()
        try:
            if invoice_type == "네이버 송장":
                df = self._load_naver_invoice(file_path)
            else:
                df = self._load_coupang_invoice(file_path)

            df = self._add_item_count_column(df, invoice_type)
            df = self._standardize_columns(df, invoice_type)

            self.current_df = df

            self._setup_image_store()
            self._show_df_in_table(df)
            self._log_columns(df, invoice_type, file_path)

            self._start_time = QDateTime.currentDateTime()
            self._dashboard_timer.start()

            self._update_dashboard_counts()
            self._update_dashboard_timer()

            # 필터 적용 (초기화 상태)
            self._apply_filter()

        except (OSError, IOError, ValueError) as e:
            QtWidgets.QMessageBox.critical(self, "엑셀 읽기 오류", str(e))
            self.log.appendPlainText(f"[오류] 엑셀을 읽는 중 문제가 발생했습니다: {e}")

    def _standardize_columns(self, df: pd.DataFrame, invoice_type: str) -> pd.DataFrame:
        """
        네이버/쿠팡의 다양한 컬럼명을 표준 컬럼명으로 변경하고,
        요청하신 'target_columns' 순서대로 정렬 및 없는 컬럼 생성
        """
        # 1. 컬럼 매핑 (원본 -> 타겟)
        rename_map = {}
        if invoice_type == "네이버 송장":
            rename_map = {
                "상품명": "품목명",
                "수취인명": "받으시는 분",
                "수취인연락처1": "받으시는 분 전화",
                "배송메세지": "특기사항",
                "옵션정보": "메모",  # 기존 로직상
            }
        else:  # 쿠팡
            rename_map = {
                "등록옵션명": "품목명",
                "상품옵션명": "품목명",  # 둘 중 하나
                "수취인이름": "받으시는 분",
                "수취인전화번호": "받으시는 분 전화",
                "배송메세지": "특기사항",
                "주문자 추가메시지": "메모",
            }

        # 실제로 존재하는 컬럼만 rename
        actual_rename = {k: v for k, v in rename_map.items() if k in df.columns}
        df.rename(columns=actual_rename, inplace=True)

        # 2. 필요한 컬럼이 없으면 생성
        for col in self.target_columns:
            if col not in df.columns:
                df[col] = ""

        # 3. 값 문자열 변환 (NaN 방지)
        for col in df.columns:
            df[col] = df[col].fillna("").astype(str)

        return df

    def on_item_double_clicked(self, item: QTableWidgetItem):
        if self.current_df is None: return

        visual_row = item.row()
        real_row = self.table.item(visual_row, 0).data(Qt.UserRole)
        col = item.column()
        header = self.table.horizontalHeaderItem(col)
        col_name = header.text() if header else ""

        # 1. 메모 수정
        if col_name == "메모":
            self._open_memo_dialog(visual_row, real_row)
            return

        # 2. 품목명 수정
        if col_name in ("품목명", "상품명"):
            cell_text = item.text()
            invoice_type = self.combo_type.currentText()

            def save_modified_text(new_full_text: str):
                try:
                    # [1] 화면 글자 즉시 변경
                    item.setText(new_full_text)
                    item.setToolTip(new_full_text)  # 툴팁도 갱신

                    # [2] 데이터 업데이트
                    self.current_df.at[real_row, col_name] = new_full_text

                    # [3] 파일 저장
                    if self.current_file:
                        self.current_df.to_excel(self.current_file, index=False, engine="openpyxl")
                        self.log.appendPlainText(f"[수정] 행 {real_row + 1} 문구 업데이트 완료.")

                    # [★핵심 해결책] 내용이 길어졌으니 행 높이를 늘려라! (이게 없어서 ...으로 뜸)
                    self.table.resizeRowsToContents()
                    self.table.viewport().update()

                except PermissionError:
                    QtWidgets.QMessageBox.critical(self, "저장 실패", "엑셀 파일이 열려있습니다.")
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

        # visual -> real mapping
        visual_row = selected[0].row()
        real_row = self.table.item(visual_row, 0).data(Qt.UserRole)

        self._current_row_idx = real_row
        self._update_sms_panel_for_row(real_row)

    def on_click_complete(self):
        """[작업 완료] 버튼: 노란색 표시 + 완료 텍스트"""
        if self.current_df is None or not self.current_file:
            QtWidgets.QMessageBox.warning(self, "경고", "열린 파일이 없습니다.")
            return

        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QtWidgets.QMessageBox.information(self, "알림", "완료 처리할 행을 선택해주세요.")
            return

        if "작업유무" not in self.current_df.columns:
            self.current_df["작업유무"] = ""
        if "문자발송처리" not in self.current_df.columns:
            self.current_df["문자발송처리"] = ""

        # 첫 번째 선택 행 기준으로 '완료'할지 '취소'할지 결정
        first_vis = selected_rows[0].row()
        first_real = self.table.item(first_vis, 0).data(Qt.UserRole)
        current_val = str(self.current_df.at[first_real, "작업유무"]).strip()

        is_already_done = (current_val == "완료")
        target_val = "" if is_already_done else "완료"  # 토글

        # [안전장치 1] 작업완료를 '해제(취소)'하려는 경우, 문자발송이 되어있는지 확인
        if is_already_done:  # 해제 시도
            for idx in selected_rows:
                v_r = idx.row()
                r_r = self.table.item(v_r, 0).data(Qt.UserRole)
                sms_status = str(self.current_df.at[r_r, "문자발송처리"]).strip()
                if sms_status == "완료":
                    QtWidgets.QMessageBox.warning(
                        self, "해제 불가",
                        "문자 발송이 완료된 주문입니다.\n먼저 [문자 발송 처리]를 취소한 뒤, 작업 완료를 해제해주세요."
                    )
                    return  # 중단

        # UI 색상 준비
        done_color = QtGui.QColor(255, 255, 100)
        white_color = QtGui.QColor(255, 255, 255)
        target_color = white_color if is_already_done else done_color

        try:
            for idx in selected_rows:
                visual_r = idx.row()
                real_r = self.table.item(visual_r, 0).data(Qt.UserRole)

                # 데이터 변경
                self.current_df.at[real_r, "작업유무"] = target_val

                # 화면 색상 변경
                for c in range(self.table.columnCount()):
                    it = self.table.item(visual_r, c)
                    if it:
                        it.setBackground(target_color)

                # 사진 버튼 색상 변경
                if "사진" in self._col_index:
                    photo_col_idx = self._col_index.get("사진")
                    widget = self.table.cellWidget(visual_r, photo_col_idx)
                    if widget:
                        bg_style = "background-color: #ffffaa;" if not is_already_done else ""
                        widget.setStyleSheet(bg_style)

            self.current_df.to_excel(self.current_file, index=False, engine="openpyxl")

            action_name = "취소" if is_already_done else "완료"
            self.log.appendPlainText(f"[작업] {len(selected_rows)}건 {action_name} 처리됨.")

            # [핵심] 변경 즉시 필터 재적용 (화면에서 사라지게 하기 위함)
            self._apply_filter()
            self._update_dashboard_counts()

        except PermissionError:
            QtWidgets.QMessageBox.critical(self, "저장 실패", "엑셀 파일이 열려있습니다.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "오류", f"저장 중 오류: {e}")

        self.table.clearSelection()

    # ------------------------------------------------------------------
    # [수정] 문자 전송: '받으시는 분' -> '구매자'로 변경
    # ------------------------------------------------------------------
    def on_send_selected(self):
        """선택된 한 명에게 문자 보내기 (ADB)"""
        if self._current_row_idx is None:
            QtWidgets.QMessageBox.information(self, "알림", "선택된 고객이 없습니다.")
            return

        # [수정] 받으시는 분 -> 구매자명 / 전화 -> 구매자연락처
        name = self._get_cell_text(self._current_row_idx, "구매자명")
        phone = self._get_cell_text(self._current_row_idx, "구매자연락처")
        item_name = self._get_cell_text(self._current_row_idx, "품목명")

        clean_phone = phone.replace("-", "").replace(" ", "").strip()

        if not clean_phone:
            QtWidgets.QMessageBox.warning(self, "오류", "구매자 전화번호가 없습니다.")
            return

        # 메시지 내용
        message = f"[하비브라운] {name}님, 주문하신 상품이 완성되어 안내 드립니다."

        self._send_via_adb(clean_phone, message)

    def on_send_all(self):
        """표시된(숨겨지지 않은) 전체 고객에게 일괄 보내기"""
        row_count = self.table.rowCount()

        # [수정] 실제로 눈에 보이는 행의 개수만 세기
        visible_rows = []
        for r in range(row_count):
            if not self.table.isRowHidden(r):
                visible_rows.append(r)

        count = len(visible_rows)
        if count == 0:
            QtWidgets.QMessageBox.information(self, "알림", "현재 화면에 표시된 고객이 없습니다.")
            return

        reply = QtWidgets.QMessageBox.question(
            self, "전체 발송",
            f"현재 화면에 표시된 {count}명에게 문자를 보낼까요?\n"
            "(필터로 숨겨진 고객은 제외됩니다)",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        if reply == QtWidgets.QMessageBox.No:
            return

        for visual_row in visible_rows:
            # 숨겨진 행은 건너뛰고 진행
            real_row = self.table.item(visual_row, 0).data(Qt.UserRole)

            name = self._get_cell_text(real_row, "구매자명")
            phone = self._get_cell_text(real_row, "구매자연락처")

            clean_phone = phone.replace("-", "").replace(" ", "").strip()
            if not clean_phone: continue

            message = f"[하비브라운] {name}님, 주문하신 상품이 완성되어 안내 드립니다."

            self._send_via_adb(clean_phone, message)

            import PyQt5.QtTest as QTest
            QTest.QTest.qWait(1500)

        self.log.appendPlainText(f"[전체발송] {count}건 명령 전달 완료.")

    # ------------------------------------------------------------------
    # [추가] 실제 ADB 명령 수행 함수
    # ------------------------------------------------------------------
    def _send_via_adb(self, phone_no: str, msg: str):
        """
        ADB로 문자 전송 명령을 내리고, 결과를 로그에 출력합니다.
        (수정: 문자 내용에 공백이 있어도 끊기지 않도록 따옴표 처리 추가)
        """
        import subprocess
        import time

        self.log.appendPlainText(f"[ADB] {phone_no}에게 전송 시도 중...")

        try:
            # [핵심 수정] 메시지 앞뒤에 따옴표(")를 붙여서 한 덩어리로 인식시킴
            # 윈도우 cmd와 안드로이드 shell 사이에서 공백이 잘리지 않게 함
            safe_msg = f'"{msg}"'

            # 1. 문자 앱 실행 및 내용 입력
            cmd = [
                "adb", "shell", "am", "start",
                "-a", "android.intent.action.SENDTO",
                "-d", f"sms:{phone_no}",
                "--es", "sms_body", safe_msg
            ]

            # 명령 실행 (CREATE_NO_WINDOW 옵션은 필요시 추가)
            proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            out, err = proc.communicate()

            # 결과 확인
            if proc.returncode == 0:
                # 에러 메시지 디코딩 (한글 윈도우는 cp949)
                err_msg = err.decode('cp949', errors='ignore').strip()

                # 'Error'나 'Exception' 문구가 있으면 실패로 간주
                if "error" in err_msg.lower() or "exception" in err_msg.lower():
                    self.log.appendPlainText(f"❌ 실패 (오류 발생): {err_msg}")
                    return

                # 성공 시 엔터키 입력 (전송 버튼 누름 효과)
                time.sleep(0.5)
                subprocess.Popen("adb shell input keyevent 66", shell=True)
                time.sleep(0.3)
                subprocess.Popen("adb shell input keyevent 66", shell=True)

                self.log.appendPlainText(f"✅ 명령 전달 성공!")

            else:
                err_msg = err.decode('cp949', errors='ignore')
                self.log.appendPlainText(f"❌ 전송 실패: {err_msg}")

        except Exception as e:
            self.log.appendPlainText(f"❌ 오류 발생: {e}")

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
        for col in ["작업유무", "메모", "문자발송처리"]:
            if col in df.columns:
                df[col] = df[col].astype(object)

        return df

    @staticmethod
    def _load_coupang_invoice(file_path: str) -> pd.DataFrame:
        df = pd.read_excel(file_path)

        # [추가] 쿠팡도 똑같이 처리
        for col in ["작업유무", "메모", "문자발송처리"]:
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

        # [수정] target_columns 기준으로 컬럼 필터링
        display_cols = [c for c in self.target_columns if c in df.columns]
        # 사진 컬럼은 UI 전용이므로 맨 뒤에 추가
        photo_col_name = "사진"
        full_cols = display_cols + [photo_col_name]

        self.table.setColumnCount(len(full_cols))
        self.table.setRowCount(len(df))
        self.table.setHorizontalHeaderLabels(full_cols)

        self._col_index = {name: idx for idx, name in enumerate(full_cols)}

        done_bg = QtGui.QColor(255, 255, 100)  # 노란색

        for row_idx in range(len(df)):
            # 작업유무 확인
            work_status = str(df.iloc[row_idx].get("작업유무", "")).strip()
            sms_status = str(df.iloc[row_idx].get("문자발송처리", "")).strip()

            is_done = (work_status == "완료")
            is_sms_done = (sms_status == "완료")

            for col_idx, col_name in enumerate(display_cols):
                val = df.iloc[row_idx][col_name]
                item = QTableWidgetItem(str(val))

                # [★핵심] 0번 컬럼(맨 앞칸)에 '진짜 데이터 번호(row_idx)'를 숨겨둡니다!
                # 화면이 뒤섞여도 이 값은 변하지 않습니다.
                if col_idx == 0:
                    item.setData(Qt.UserRole, row_idx)

                # 완료된 행 색칠
                if is_done:
                    item.setBackground(done_bg)

                # 문자발송 완료시 취소선 (Delegate에서 처리)
                if is_sms_done:
                    item.setData(Qt.UserRole + 2, True)

                # 툴팁 등 기존 옵션
                if col_name in ("품목명", "상품명"):
                    item.setToolTip(str(val))

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
            # 정렬 상태일 땐 row_idx가 뒤죽박죽이므로 real index를 찾아야 함
            item_0 = self.table.item(row_idx, 0)
            if not item_0: continue

            real_row = item_0.data(Qt.UserRole)
            row_id = real_row + 1

            widget = self.table.cellWidget(row_idx, photo_col_idx)
            if isinstance(widget, QPushButton):
                count = len(self._image_map.get(row_id, []))
                widget.setText(f"사진({count}장)…")

    # ------------------------------------------------------------------
    # 로그 출력 (복구됨: 상세 컬럼 리스트)
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
    def _add_item_count_column(self, df: pd.DataFrame, invoice_type: str) -> pd.DataFrame:
        col_key = None
        for cand in ("품목명", "상품명"):
            if cand in df.columns:
                col_key = cand
                break
        if col_key is None:
            return df

        # [설정] 갯수를 셀 '유형 코드' 목록 (여기에 없는 줄은 '=> 숫자 ea'가 있어도 무시함)
        target_codes = ["M1)", "S)", "W2)"]

        counts: List[int] = []
        for idx in range(len(df)):
            val = df.iloc[idx][col_key]
            text = "" if pd.isna(val) else str(val)

            row_total = 0
            # 줄 단위로 쪼개서 확인
            for line in text.splitlines():
                # 1. 해당 줄에 우리가 원하는 코드(M1, S, W2)가 있는지 확인
                if any(code in line for code in target_codes):
                    # 2. 있다면 '=> 숫자 ea' 패턴을 찾아 숫자만 추출
                    match = re.search(r"=>\s*(\d+)\s*ea", line)
                    if match:
                        row_total += int(match.group(1))

            counts.append(row_total)

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
        # 주의: row_idx는 visual row가 아니라 real row여야 함 (DataFrame 접근용)
        # 하지만 여기서 row_idx는 호출하는 쪽에서 무엇을 넘기냐에 따라 다름.
        # 이 함수는 DataFrame 접근용이므로 real_row_idx를 받아야 합니다.
        if self.current_df is None: return ""
        val = self.current_df.iloc[row_idx][col_name]
        return "" if pd.isna(val) else str(val)

    def _update_sms_panel_for_row(self, row_idx: int):
        # row_idx는 real index
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
        # row_idx = real index
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

        # [수정] 캐시 무시하고 로드 (여기서도 중요!)
        pix = load_pixmap_clean(str(fpath))

        if pix.isNull():
            self._clear_preview()
            return

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
        # visual row
        visual_row = selected_rows[0].row()
        real_row = self.table.item(visual_row, 0).data(Qt.UserRole)
        self._open_memo_dialog(visual_row, real_row)

    # [수정] 인자를 (visual_row, real_row) 두 개 받도록 변경
    def _open_memo_dialog(self, visual_row: int, real_row: int):
        if self.current_df is None or not self.current_file: return
        col_idx = self._col_index.get("메모")
        if col_idx is None: return

        current_item = self.table.item(visual_row, col_idx)
        current_text = current_item.text() if current_item else ""

        dlg = MemoDialog(current_text, self)
        if dlg.exec_() == QDialog.Accepted:
            new_text = dlg.get_text()
            try:
                # [핵심] 화면 텍스트 즉시 변경
                if current_item:
                    current_item.setText(new_text)
                else:
                    self.table.setItem(visual_row, col_idx, QTableWidgetItem(new_text))

                # 데이터 업데이트
                if "메모" not in self.current_df.columns:
                    self.current_df.insert(0, "메모", "")
                self.current_df.at[real_row, "메모"] = new_text

                # 파일 저장
                self.current_df.to_excel(self.current_file, index=False, engine="openpyxl")
                self.log.appendPlainText(f"[메모 저장] 행 {real_row + 1}: {new_text}")

                # 화면 강제 갱신
                self.table.repaint()

            except PermissionError:
                QtWidgets.QMessageBox.critical(self, "저장 실패", "엑셀 파일이 열려있습니다.")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "오류", f"저장 중 오류: {e}")

    # ------------------------------------------------------------------
    # 대시보드 업데이트 (복구됨: 상세 로직)
    # ------------------------------------------------------------------
    def _update_dashboard_counts(self):
        """
        [가운데 창] 업데이트:
        1. 전체 작업 (완료/전체)
        2. 각인 갯수 (완료/전체) -> M1, S, W2 코드가 있는 줄의 ea만 합산
        3. 시작 시간
        """
        if self.current_df is None or self._start_time is None:
            self.info_dash_counts.setPlainText("▶ 파일 로드 대기 중...")
            return

        total_rows = len(self.current_df)
        target_col = None
        for col in ["품목명", "상품명"]:
            if col in self.current_df.columns:
                target_col = col
                break

        # [설정] 카운팅 대상 코드
        target_codes = ["M1)", "S)", "W2)"]

        completed_rows = 0
        total_gagin_count = 0
        completed_gagin_count = 0

        if target_col:
            if "작업유무" not in self.current_df.columns:
                self.current_df["작업유무"] = ""

            for i in range(total_rows):
                item_name = str(self.current_df.iloc[i][target_col])
                status = str(self.current_df.iloc[i]["작업유무"]).strip()
                is_done = (status == "완료")

                if is_done:
                    completed_rows += 1

                # 행 내부의 줄(Line) 별로 검사
                row_gagin = 0
                for line in item_name.splitlines():
                    # M1, S, W2 중 하나라도 포함되어 있어야 카운트
                    if any(code in line for code in target_codes):
                        match = re.search(r"=>\s*(\d+)\s*ea", line)
                        if match:
                            row_gagin += int(match.group(1))

                total_gagin_count += row_gagin
                if is_done:
                    completed_gagin_count += row_gagin

        start_time_str = self._start_time.toString("yyyy-MM-dd HH:mm:ss")
        text_counts = (
            f"📊 [작업 현황]\n"
            f"━━━━━━━━━━━━━━━━━━\n"
            f"1. 전체 작업 갯수 :  {completed_rows} / {total_rows} 건\n"
            f"   (완료 / 전체 행)\n\n"
            f"2. 각인 갯수 현황 :  {completed_gagin_count} / {total_gagin_count} 개\n"
            f"   (완료 / 전체 수량(ea))\n"
            f"   ※ M1, S, W2 포함 항목만 집계\n\n"
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

    # ------------------------------------------------------------------
    # [추가] 이미지 클릭 시 사진 관리창 열기
    # ------------------------------------------------------------------
    def _on_preview_clicked(self, event):
        """이미지 라벨 클릭 시 실행: 현재 선택된 줄의 사진 관리창을 엽니다."""

        # 1. 현재 테이블에서 선택된 줄이 있는지 확인
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QtWidgets.QMessageBox.information(self, "알림", "먼저 목록에서 고객을 선택해주세요.")
            return

        # 2. 눈에 보이는 줄 번호 (Visual Row)
        visual_row = selected_rows[0].row()

        # 3. [핵심] 섞여도 변하지 않는 '진짜 엑셀 번호' (Real Row) 꺼내기
        real_row = self.table.item(visual_row, 0).data(Qt.UserRole)

        # 4. 관리창 열기
        self._open_image_manage_dialog(real_row)

    # [수정] 이미지 클릭 시 열리는 관리창 로직을 대폭 수정하여 동기화 문제 해결
    def _open_image_manage_dialog(self, real_row: int):
        """사진 관리창 열기 및 저장 처리"""
        if not self._image_dir:
            QtWidgets.QMessageBox.information(self, "알림", "이미지 저장 경로가 설정되지 않았습니다.")
            return

        # [핵심 수정] current_df가 아닌 _image_map(최신 상태)을 사용해야 함
        row_id = real_row + 1
        current_files = self._image_map.get(row_id, [])

        # 2. 다이얼로그 열기
        dlg = ImageManageDialog(self, row_id, self._image_dir, current_files)

        if dlg.exec_() == QDialog.Accepted:
            # 3. 변경된 내용 저장
            new_files = dlg.images()

            # (A) 메모리 데이터(_image_map) 갱신
            if new_files:
                self._image_map[row_id] = new_files
            else:
                self._image_map.pop(row_id, None)

            # (B) meta.json 저장 (사진 관리의 핵심)
            self._save_image_meta()

            # (C) 엑셀 데이터프레임에도 문자열로 기록 (선택 사항이나 싱크 맞추기용)
            if "이미지" not in self.current_df.columns:
                self.current_df["이미지"] = ""
            self.current_df.at[real_row, "이미지"] = ",".join(new_files)

            # (D) 엑셀 파일 저장
            if self.current_file:
                try:
                    self.current_df.to_excel(self.current_file, index=False, engine="openpyxl")
                    self.log.appendPlainText(f"[사진] 행 {row_id} 변경사항 저장 완료.")
                except Exception as e:
                    self.log.appendPlainText(f"[오류] 엑셀 저장 실패: {e}")

            # (E) [핵심] 화면 갱신 (미리보기 사진 + 테이블 버튼 글자)
            self._update_preview_for_row(real_row)
            self._refresh_photo_buttons()
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