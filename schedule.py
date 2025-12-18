# schedule.py
# -*- coding: utf-8 -*-
# 달력 셀 내부에 스케쥴 제목 표시 + 날짜 숫자 왼쪽 위 정렬 + CRUD + 완료 취소선
# 저장: schedule_data.json (schedule.py와 같은 폴더)

from __future__ import annotations

import json
import uuid
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Callable

from PyQt5 import QtWidgets, QtCore, QtGui


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _elide_by_px(text: str, fm: QtGui.QFontMetrics, max_px: int) -> str:
    return fm.elidedText(text, QtCore.Qt.ElideRight, max_px)


@dataclass
class ScheduleItem:
    id: str
    date: str  # "yyyy-MM-dd"
    title: str
    content: str
    completed: bool = False
    created_at: str = ""
    updated_at: str = ""

    @staticmethod
    def from_dict(d: dict) -> "ScheduleItem":
        return ScheduleItem(
            id=str(d.get("id", "")),
            date=str(d.get("date", "")),
            title=str(d.get("title", "")),
            content=str(d.get("content", "")),
            completed=bool(d.get("completed", False)),
            created_at=str(d.get("created_at", "")),
            updated_at=str(d.get("updated_at", "")),
        )


class ScheduleStore:
    """로컬 JSON 저장소"""

    def __init__(self, path: Path):
        self.path = path
        self.items: Dict[str, ScheduleItem] = {}

    def load(self) -> None:
        if not self.path.exists():
            self.items = {}
            return
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            raw_items = data.get("items", [])
            self.items = {}
            for r in raw_items:
                it = ScheduleItem.from_dict(r)
                if it.id:
                    self.items[it.id] = it
        except Exception:
            self.items = {}

    def save(self) -> None:
        payload = {
            "version": 1,
            "items": [asdict(x) for x in self.items.values()],
            "saved_at": _now_iso(),
        }
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        tmp.replace(self.path)

    def list_all_sorted(self) -> List[ScheduleItem]:
        def key(it: ScheduleItem):
            return (it.date, it.updated_at or it.created_at)
        return sorted(self.items.values(), key=key, reverse=True)

    def list_by_date(self, date_str: str) -> List[ScheduleItem]:
        items = [x for x in self.items.values() if x.date == date_str]
        return sorted(items, key=lambda x: (x.updated_at or x.created_at), reverse=True)

    def add(self, date_str: str, title: str, content: str, completed: bool) -> ScheduleItem:
        _id = uuid.uuid4().hex
        now = _now_iso()
        it = ScheduleItem(
            id=_id,
            date=date_str,
            title=title,
            content=content,
            completed=completed,
            created_at=now,
            updated_at=now,
        )
        self.items[_id] = it
        self.save()
        return it

    def update(self, item_id: str, *, date: str, title: str, content: str, completed: bool) -> None:
        it = self.items.get(item_id)
        if not it:
            return
        it.date = date
        it.title = title
        it.content = content
        it.completed = completed
        it.updated_at = _now_iso()
        self.save()

    def delete(self, item_id: str) -> None:
        if item_id in self.items:
            del self.items[item_id]
            self.save()


class ScheduleEditDialog(QtWidgets.QDialog):
    """추가/수정 공용 다이얼로그 (키워드/포지셔널 모두 안전하게 받음)"""

    def __init__(self, parent=None, *args, **kwargs):
        super().__init__(parent)

        mode = kwargs.pop("mode", None)
        initial_date = kwargs.pop("initial_date", None)
        item = kwargs.pop("item", None)

        # 포지셔널 보정: (mode, initial_date, item)
        if mode is None and len(args) >= 1:
            mode = args[0]
        if initial_date is None and len(args) >= 2:
            initial_date = args[1]
        if item is None and len(args) >= 3:
            item = args[2]

        mode = mode or "add"
        initial_date = initial_date or ""

        self.setWindowTitle("스케쥴 추가" if mode == "add" else "스케쥴 수정")
        self.setModal(True)
        self.setMinimumWidth(520)

        self._mode = mode
        self._item = item

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        root.addLayout(form)

        # 날짜
        self.date_edit = QtWidgets.QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy-MM-dd")

        if item and getattr(item, "date", ""):
            self.date_edit.setDate(QtCore.QDate.fromString(item.date, "yyyy-MM-dd"))
        else:
            qd = QtCore.QDate.fromString(str(initial_date), "yyyy-MM-dd")
            self.date_edit.setDate(qd if qd.isValid() else QtCore.QDate.currentDate())

        form.addRow("날짜", self.date_edit)

        # 제목
        self.le_title = QtWidgets.QLineEdit()
        self.le_title.setPlaceholderText("예) 네이버 송장 업로드 점검")
        if item:
            self.le_title.setText(getattr(item, "title", ""))
        form.addRow("제목", self.le_title)

        # 내용
        self.te_content = QtWidgets.QTextEdit()
        self.te_content.setPlaceholderText("내용을 입력하세요.")
        self.te_content.setMinimumHeight(180)
        if item:
            self.te_content.setPlainText(getattr(item, "content", ""))
        form.addRow("내용", self.te_content)

        # 완료
        self.cb_completed = QtWidgets.QCheckBox("완료 처리")
        if item:
            self.cb_completed.setChecked(bool(getattr(item, "completed", False)))
        form.addRow("", self.cb_completed)

        # 버튼
        btns = QtWidgets.QHBoxLayout()
        root.addLayout(btns)
        btns.addStretch(1)

        self.btn_save = QtWidgets.QPushButton("저장")
        self.btn_cancel = QtWidgets.QPushButton("취소")
        btns.addWidget(self.btn_save)
        btns.addWidget(self.btn_cancel)

        self.btn_save.clicked.connect(self._on_save)
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_save.setDefault(True)

    def _on_save(self):
        title = self.le_title.text().strip()
        if not title:
            QtWidgets.QMessageBox.warning(self, "입력 오류", "제목을 입력해 주세요.")
            return
        self.accept()

    def get_values(self):
        date_str = self.date_edit.date().toString("yyyy-MM-dd")
        title = self.le_title.text().strip()
        content = self.te_content.toPlainText().strip()
        completed = self.cb_completed.isChecked()
        return date_str, title, content, completed


class CalendarScheduleDelegate(QtWidgets.QStyledItemDelegate):
    """
    달력 셀 커스텀 렌더링:
      - 날짜 숫자: 왼쪽 위 정렬
      - 셀 내부에 스케쥴 제목 1~2줄 표시(작은 글씨, 길면 …)
      - 완료 스케쥴: 취소선
      - 환경별 row/col 오프셋 문제를 "자동 보정(delta_days)"으로 해결
    """

    def __init__(
            self,
            calendar: QtWidgets.QCalendarWidget,
            parent=None,
            get_items_for_date=None,
    ):
        super().__init__(parent)
        self.calendar = calendar
        self.get_items_for_date = get_items_for_date

        # 셀 안에 제목 5개까지
        self.max_lines = 5
        self.pad = 4

        # 환경별 보정
        self.delta_days = 0

    def _cell_date_raw(self, row: int, col: int) -> QtCore.QDate:
        """
        (보정 전) row/col -> date
        """
        year = self.calendar.yearShown()
        month = self.calendar.monthShown()
        first_of_month = QtCore.QDate(year, month, 1)

        first_dow = int(self.calendar.firstDayOfWeek())  # 1=Mon ... 7=Sun
        month_dow = first_of_month.dayOfWeek()          # 1=Mon ... 7=Sun

        offset = (month_dow - first_dow) % 7
        start_date = first_of_month.addDays(-offset)

        return start_date.addDays(row * 7 + col)

    def _cell_date(self, row: int, col: int) -> QtCore.QDate:
        """
        (보정 후) row/col -> date
        """
        return self._cell_date_raw(row, col).addDays(self.delta_days)

    def paint(self, painter: QtGui.QPainter, option: QtWidgets.QStyleOptionViewItem, index: QtCore.QModelIndex):
        painter.save()

        # 기본 셀(선택/호버 포함)
        style = option.widget.style() if option.widget else QtWidgets.QApplication.style()
        style.drawPrimitive(QtWidgets.QStyle.PE_PanelItemViewItem, option, painter, option.widget)

        row = index.row()
        col = index.column()

        qdate = self._cell_date(row, col)
        date_str = qdate.toString("yyyy-MM-dd")

        outer = option.rect
        rect = outer.adjusted(self.pad, self.pad, -self.pad, -self.pad)

        is_current_month = (qdate.month() == self.calendar.monthShown() and qdate.year() == self.calendar.yearShown())

        items = []
        if self.get_items_for_date:
            items = self.get_items_for_date(date_str) or []

        # 일정 있으면 배경 강조
        if items:
            has_active = any(not x.completed for x in items)
            bg = QtGui.QColor(255, 245, 200) if has_active else QtGui.QColor(235, 235, 235)
            bg.setAlpha(140)
            painter.fillRect(outer, bg)

        # 1) 날짜 숫자: 왼쪽 위
        day_font = QtGui.QFont(option.font)
        day_font.setBold(True)
        painter.setFont(day_font)

        day_color = option.palette.color(QtGui.QPalette.Text)
        if not is_current_month:
            day_color = QtGui.QColor(140, 140, 140)
        painter.setPen(day_color)

        fm_day = QtGui.QFontMetrics(day_font)
        day_h = fm_day.height()

        day_rect = QtCore.QRect(rect.left(), rect.top(), rect.width(), day_h)
        painter.drawText(day_rect, QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop, str(qdate.day()))

        # 2) 셀 내부: 제목 리스트(작은 글씨)
        if items:
            text_font = QtGui.QFont(option.font)
            text_font.setPointSize(max(8, text_font.pointSize() - 2))
            fm = QtGui.QFontMetrics(text_font)
            line_h = fm.height()

            text_top = rect.top() + day_h + 2
            text_rect = QtCore.QRect(rect.left(), text_top, rect.width(), rect.bottom() - text_top)

            y = text_rect.top()

            for it in items[: self.max_lines]:
                f2 = QtGui.QFont(text_font)
                f2.setStrikeOut(bool(getattr(it, "completed", False)))
                painter.setFont(f2)

                line = fm.elidedText(str(getattr(it, "title", "")), QtCore.Qt.ElideRight, text_rect.width())
                line_rect = QtCore.QRect(text_rect.left(), y, text_rect.width(), line_h)

                color = option.palette.color(QtGui.QPalette.Text)
                if not is_current_month:
                    color = QtGui.QColor(140, 140, 140)
                painter.setPen(color)

                painter.drawText(line_rect, QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter, line)
                y += line_h

            remain = len(items) - self.max_lines
            if remain > 0 and y + line_h <= rect.bottom():
                painter.setFont(text_font)
                painter.setPen(QtGui.QColor(90, 90, 90))
                painter.drawText(
                    QtCore.QRect(text_rect.left(), y, text_rect.width(), line_h),
                    QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter,
                    f"+{remain}",
                )

        painter.restore()




class ScheduleWidget(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self._data_path = Path(__file__).resolve().parent / "schedule_data.json"
        self.store = ScheduleStore(self._data_path)
        self.store.load()

        root = QtWidgets.QHBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(12)

        # =========================================
        # LEFT (달력 꽉 차게)
        # =========================================
        left = QtWidgets.QVBoxLayout()
        left.setSpacing(10)
        root.addLayout(left, 6)

        title = QtWidgets.QLabel("스케쥴 (달력)")
        title.setStyleSheet("font-weight: bold; font-size: 14px;")
        left.addWidget(title)

        self.calendar = QtWidgets.QCalendarWidget()
        self.calendar.setGridVisible(True)
        self.calendar.setVerticalHeaderFormat(QtWidgets.QCalendarWidget.NoVerticalHeader)
        self.calendar.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)

        # 달력을 “왼쪽 영역에 꽉” 채우기: 최소 크기만 잡고, 아래 요소는 최소화
        self.calendar.setMinimumSize(860, 720)

        # 내부 테이블 뷰 확보 후 delegate 장착
        self._cal_view = self.calendar.findChild(QtWidgets.QTableView)
        if self._cal_view is None:
            raise RuntimeError("QCalendarWidget 내부 QTableView를 찾지 못했습니다.")

        self._cal_delegate = CalendarScheduleDelegate(
            self.calendar,
            self._cal_view,
            get_items_for_date=self._get_items_for_date_for_calendar,
        )
        # 셀에 제목 5개까지 표시
        self._cal_delegate.max_lines = 5
        self._cal_view.setItemDelegate(self._cal_delegate)

        left.addWidget(self.calendar, 1)

        # 하단: 선택 날짜 라벨 + 버튼 (달력 밑에 얇게)
        self.lbl_date = QtWidgets.QLabel("선택 날짜: -")
        self.lbl_date.setStyleSheet("font-weight: bold;")
        left.addWidget(self.lbl_date)

        btns = QtWidgets.QHBoxLayout()
        left.addLayout(btns)

        self.btn_add = QtWidgets.QPushButton("스케쥴 추가")
        self.btn_edit = QtWidgets.QPushButton("수정")
        self.btn_delete = QtWidgets.QPushButton("삭제")
        btns.addWidget(self.btn_add)
        btns.addWidget(self.btn_edit)
        btns.addWidget(self.btn_delete)
        btns.addStretch(1)

        # =========================================
        # RIGHT (각 영역이 각각 스크롤 + 4:3:3)
        # =========================================
        right = QtWidgets.QVBoxLayout()
        right.setSpacing(10)
        root.addLayout(right, 4)

        # 1) 전체 제목 리스트 (비율 4)
        all_box = QtWidgets.QGroupBox("전체 스케쥴 제목 리스트")
        all_layout = QtWidgets.QVBoxLayout(all_box)
        all_layout.setContentsMargins(8, 10, 8, 8)
        all_layout.setSpacing(8)

        self.all_list = QtWidgets.QListWidget()
        self.all_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.all_list.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.all_list.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)  # ✅ 이 박스만 스크롤
        all_layout.addWidget(self.all_list, 1)

        right.addWidget(all_box, 4)

        # 2) 선택 날짜 스케쥴 (비율 3)
        day_box = QtWidgets.QGroupBox("선택 날짜 스케쥴")
        day_layout = QtWidgets.QVBoxLayout(day_box)
        day_layout.setContentsMargins(8, 10, 8, 8)
        day_layout.setSpacing(6)

        self.day_list = QtWidgets.QListWidget()
        self.day_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.day_list.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.day_list.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)  # ✅ 이 박스만 스크롤
        day_layout.addWidget(self.day_list, 1)

        right.addWidget(day_box, 3)

        # 3) 상세 내용 (비율 3)
        detail_box = QtWidgets.QGroupBox("상세 내용")
        detail_layout = QtWidgets.QVBoxLayout(detail_box)
        detail_layout.setContentsMargins(8, 10, 8, 8)
        detail_layout.setSpacing(6)

        self.detail = QtWidgets.QPlainTextEdit()
        self.detail.setReadOnly(True)
        self.detail.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)  # ✅ 이 박스만 스크롤
        detail_layout.addWidget(self.detail, 1)

        right.addWidget(detail_box, 3)

        # =========================================
        # 시그널
        # =========================================
        self.calendar.selectionChanged.connect(self._on_date_changed)
        self.calendar.currentPageChanged.connect(lambda *_: self._calibrate_calendar_delta())

        self.btn_add.clicked.connect(self._on_add_clicked)
        self.btn_edit.clicked.connect(self._on_edit_clicked)
        self.btn_delete.clicked.connect(self._on_delete_clicked)

        self.all_list.currentItemChanged.connect(self._on_all_selection_changed)

        # ✅ day_list 클릭해도 상세내용 갱신되도록
        self.day_list.currentItemChanged.connect(self._on_day_selection_changed)

        # =========================================
        # 초기 보정 + 초기 갱신
        # =========================================
        self._calibrate_calendar_delta()
        self._on_date_changed()
        self._refresh_all_list()
        self._select_first_item_if_any()
        self._refresh_calendar_view()

    # ---------- Calendar data ----------
    def _get_items_for_date_for_calendar(self, date_str: str) -> List[ScheduleItem]:
        # 달력 셀 내부 표시용: 정렬(미완료 우선 -> 최신)
        items = self.store.list_by_date(date_str)
        # 미완료 먼저
        items = sorted(items, key=lambda x: (x.completed, -(datetime.fromisoformat(x.updated_at or x.created_at).timestamp() if (x.updated_at or x.created_at) else 0)))
        return items

    def _refresh_calendar_view(self) -> None:
        # delegate로 그리는 달력 셀 다시 그리기
        if hasattr(self, "_cal_view") and self._cal_view is not None:
            self._cal_view.viewport().update()
        else:
            try:
                self.calendar.updateCells()
            except Exception:
                self.calendar.viewport().update()

    # ---------- List rendering ----------
    def _apply_completed_style(self, item: QtWidgets.QListWidgetItem, completed: bool) -> None:
        f = item.font()
        f.setStrikeOut(bool(completed))
        item.setFont(f)

    def _refresh_day_list(self) -> None:
        self.day_list.clear()
        date_str = self._selected_date_str()
        items = self.store.list_by_date(date_str)

        if not items:
            x = QtWidgets.QListWidgetItem("스케쥴 없음")
            x.setFlags(QtCore.Qt.NoItemFlags)
            self.day_list.addItem(x)
            return

        for it in items:
            li = QtWidgets.QListWidgetItem(it.title)
            li.setData(QtCore.Qt.UserRole, it.id)
            self._apply_completed_style(li, it.completed)
            self.day_list.addItem(li)

    def _refresh_all_list(self) -> None:
        self.all_list.clear()
        items = self.store.list_all_sorted()
        if not items:
            x = QtWidgets.QListWidgetItem("전체 스케쥴 없음")
            x.setFlags(QtCore.Qt.NoItemFlags)
            self.all_list.addItem(x)
            self.detail.setPlainText("")
            return

        for it in items:
            prefix = "(완료) " if it.completed else ""
            li = QtWidgets.QListWidgetItem(f"{it.date} | {prefix}{it.title}")
            li.setData(QtCore.Qt.UserRole, it.id)
            self._apply_completed_style(li, it.completed)
            self.all_list.addItem(li)

    def _select_first_item_if_any(self) -> None:
        if self.all_list.count() <= 0:
            return
        first = self.all_list.item(0)
        if first and (first.flags() & QtCore.Qt.ItemIsEnabled):
            self.all_list.setCurrentRow(0)

    def _selected_date_str(self) -> str:
        return self.calendar.selectedDate().toString("yyyy-MM-dd")

    def _render_detail(self, it: Optional[ScheduleItem]) -> None:
        if not it:
            self.detail.setPlainText("")
            return
        status = "완료" if it.completed else "진행중"
        text = (
            f"날짜: {it.date}\n"
            f"상태: {status}\n"
            f"제목: {it.title}\n"
            f"생성: {it.created_at}\n"
            f"수정: {it.updated_at}\n"
            f"\n"
            f"{it.content}"
        )
        self.detail.setPlainText(text)

    # ---------- Events ----------
    def _on_date_changed(self) -> None:
        # 클릭할 때마다 보정 먼저(셀-날짜 싱크 맞춤)
        self._calibrate_calendar_delta()

        qdate = self.calendar.selectedDate()
        date_str = qdate.toString("yyyy-MM-dd (ddd)")
        self.lbl_date.setText(f"선택 날짜: {date_str}")

        self._refresh_day_list()
        self._refresh_calendar_view()

    def _get_selected_all_item_id(self) -> Optional[str]:
        cur = self.all_list.currentItem()
        if not cur:
            return None
        item_id = cur.data(QtCore.Qt.UserRole)
        return item_id if isinstance(item_id, str) and item_id else None

    def _on_all_selection_changed(self, cur: QtWidgets.QListWidgetItem, prev: QtWidgets.QListWidgetItem) -> None:
        item_id = self._get_selected_all_item_id()
        if not item_id:
            self._render_detail(None)
            return
        it = self.store.items.get(item_id)
        self._render_detail(it)

    def _on_add_clicked(self) -> None:
        date_str = self._selected_date_str()
        dlg = ScheduleEditDialog(self, mode="add", initial_date=date_str, item=None)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return

        d, title, content, completed = dlg.get_values()
        self.store.add(d, title, content, completed)

        qd = QtCore.QDate.fromString(d, "yyyy-MM-dd")
        if qd.isValid():
            self.calendar.setSelectedDate(qd)

        self._refresh_day_list()
        self._refresh_all_list()
        self._refresh_calendar_view()

    def _on_edit_clicked(self) -> None:
        item_id = self._get_selected_all_item_id()
        if not item_id:
            QtWidgets.QMessageBox.information(self, "안내", "오른쪽 제목 리스트에서 수정할 항목을 선택해 주세요.")
            return

        it = self.store.items.get(item_id)
        if not it:
            return

        dlg = ScheduleEditDialog(self, mode="edit", initial_date=it.date, item=it)
        if dlg.exec_() != QtWidgets.QDialog.Accepted:
            return

        d, title, content, completed = dlg.get_values()
        self.store.update(item_id, date=d, title=title, content=content, completed=completed)

        qd = QtCore.QDate.fromString(d, "yyyy-MM-dd")
        if qd.isValid():
            self.calendar.setSelectedDate(qd)

        self._refresh_day_list()
        self._refresh_all_list()
        self._refresh_calendar_view()

    def _on_delete_clicked(self) -> None:
        item_id = self._get_selected_all_item_id()
        if not item_id:
            QtWidgets.QMessageBox.information(self, "안내", "오른쪽 제목 리스트에서 삭제할 항목을 선택해 주세요.")
            return

        it = self.store.items.get(item_id)
        if not it:
            return

        msg = f"다음 스케쥴을 삭제할까요?\n\n{it.date} | {it.title}"
        ret = QtWidgets.QMessageBox.question(
            self,
            "삭제 확인",
            msg,
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            QtWidgets.QMessageBox.No,
        )
        if ret != QtWidgets.QMessageBox.Yes:
            return

        self.store.delete(item_id)

        self._refresh_day_list()
        self._refresh_all_list()
        self._refresh_calendar_view()
        self._select_first_item_if_any()

    def _calibrate_calendar_delta(self) -> None:
        """
        현재 선택된 셀(row/col)에 대해:
          - delegate가 계산한 날짜(보정 전) vs calendar.selectedDate()를 비교
          - 차이를 delta_days로 저장 -> 이후 그리기가 선택과 일치하게 됨
        """
        if not hasattr(self, "_cal_view") or self._cal_view is None:
            return
        if not hasattr(self, "_cal_delegate") or self._cal_delegate is None:
            return

        idx = self._cal_view.currentIndex()
        if not idx.isValid():
            return

        row = idx.row()
        col = idx.column()

        raw = self._cal_delegate._cell_date_raw(row, col)
        selected = self.calendar.selectedDate()

        if raw.isValid() and selected.isValid():
            # raw에서 selected까지 몇 일 차이인지
            delta = raw.daysTo(selected)
            self._cal_delegate.delta_days = delta

    def _get_selected_day_item_id(self) -> Optional[str]:
        cur = self.day_list.currentItem()
        if not cur:
            return None
        item_id = cur.data(QtCore.Qt.UserRole)
        if not item_id or not isinstance(item_id, str):
            return None
        return item_id

    def _on_day_selection_changed(self, cur: QtWidgets.QListWidgetItem, prev: QtWidgets.QListWidgetItem) -> None:
        """
        day_list(선택 날짜 스케쥴) 클릭 시:
          - 해당 id를 찾아 all_list도 같은 항목을 선택시키고
          - 상세내용은 기존 all_list 핸들러(_on_all_selection_changed) 로직을 그대로 타게 한다.
        """
        item_id = self._get_selected_day_item_id()
        if not item_id:
            return

        # ✅ all_list에서 같은 id를 찾아 선택(이때 _on_all_selection_changed가 상세 갱신)
        for i in range(self.all_list.count()):
            it = self.all_list.item(i)
            if it and it.data(QtCore.Qt.UserRole) == item_id:
                # 불필요한 재귀/튐 방지: day_list 쪽은 그대로 두고 all_list만 동기화
                self.all_list.blockSignals(True)
                self.all_list.setCurrentRow(i)
                self.all_list.scrollToItem(it, QtWidgets.QAbstractItemView.PositionAtCenter)
                self.all_list.blockSignals(False)

                # ✅ 상세내용 직접 갱신 (blockSignals로 막았으니 여기서 렌더)
                obj = self.store.items.get(item_id)
                self._render_detail(obj)
                return

