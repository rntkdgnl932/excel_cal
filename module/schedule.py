# schedule.py
# -*- coding: utf-8 -*-
# 달력 셀 내부에 스케쥴 제목 표시 + 날짜 숫자 왼쪽 위 정렬 + CRUD + 완료 취소선
# 저장: schedule_data.json (schedule.py와 같은 폴더)

from __future__ import annotations

import json
import uuid
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from PyQt5.QtCore import QLocale
from PyQt5.QtWidgets import QCalendarWidget
from typing import Dict, List, Optional, Callable
import os
import sys
import faulthandler
from pathlib import Path
from PyQt5 import QtWidgets, QtCore, QtGui

# =========================================================
# Google Calendar Sync (OAuth Desktop)
#   - client_secret.json / token.json 경로를 고정
#   - all-day event로 생성/수정/삭제
# =========================================================
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

SCOPES = ["https://www.googleapis.com/auth/calendar"]




#########################################################



# =========================
# Crash logging (file-based)
# =========================
_LOG_DIR = Path(r"C:\my_games\excel_cal\log")
_LOG_DIR.mkdir(parents=True, exist_ok=True)

_CRASH_MARK_PATH = _LOG_DIR / "crash_mark.log"
_EXCEPTION_DUMP_PATH = _LOG_DIR / "exception_dump.log"
_CRASH_DUMP_PATH = _LOG_DIR / "crash_dump.log"

# faulthandler: 파이썬이 잡을 수 있는 크래시/덤프가 있으면 여기에 기록
_fh = open(_CRASH_DUMP_PATH, "a", buffering=1, encoding="utf-8")
faulthandler.enable(file=_fh, all_threads=True)

def _crash_mark(msg: str) -> None:
    """네이티브 크래시 직전 위치 파악용: flush+fsync로 강제 기록"""
    try:
        with _CRASH_MARK_PATH.open("a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat(timespec='seconds')} | {msg}\n")
            f.flush()
            os.fsync(f.fileno())
    except Exception:
        pass

def _dump_exception(exctype, value, tb) -> None:
    """파이썬 예외는 무조건 파일로 남김"""
    try:
        import traceback
        with _EXCEPTION_DUMP_PATH.open("a", encoding="utf-8") as f:
            f.write("".join(traceback.format_exception(exctype, value, tb)))
            f.write("\n")
            f.flush()
            os.fsync(f.fileno())
    except Exception:
        pass

# 메인 스레드 예외
sys.excepthook = _dump_exception



##########################################################


class HolidayManager:
    """
    공휴일(빨간날) 및 기념일(회색날)을 관리합니다.
    사용자 정의 휴일은 holiday_data.json에 저장합니다.
    """

    def __init__(self, data_path: Path):
        self.path = data_path
        self.custom_holidays = {}  # "MM-DD": "이름" or "YYYY-MM-DD": "이름"
        self.load()

    def load(self):
        if self.path.exists():
            try:
                self.custom_holidays = json.loads(self.path.read_text(encoding='utf-8'))
            except:
                self.custom_holidays = {}

    def save(self):
        self.path.write_text(json.dumps(self.custom_holidays, indent=2, ensure_ascii=False), encoding='utf-8')

    def add_custom_holiday(self, date_str: str, name: str):
        # date_str: "2025-05-06" (특정일)
        self.custom_holidays[date_str] = name
        self.save()

    def remove_custom_holiday(self, date_str: str):
        if date_str in self.custom_holidays:
            del self.custom_holidays[date_str]
            self.save()

    def get_holiday_name(self, date_obj: QtCore.QDate) -> Optional[str]:
        """공휴일(빨간날)이면 이름을 반환, 아니면 None"""
        y, m, d = date_obj.year(), date_obj.month(), date_obj.day()
        iso_date = date_obj.toString("yyyy-MM-dd")
        md_date = date_obj.toString("MM-dd")

        # 1) 사용자 지정 휴일 (우선순위 높음)
        if iso_date in self.custom_holidays:
            return self.custom_holidays[iso_date]

        # 2) 양력 고정 공휴일
        solar_holidays = {
            "01-01": "신정", "03-01": "3.1절", "05-05": "어린이날",
            "06-06": "현충일", "08-15": "광복절", "10-03": "개천절",
            "10-09": "한글날", "12-25": "성탄절"
        }
        if md_date in solar_holidays:
            return solar_holidays[md_date]

        # 3) 음력 주요 공휴일 (2025~2030 하드코딩 - 라이브러리 없이 구현)
        # 필요시 더 추가하거나, 라이브러리(korean_lunar_calendar) 도입 고려
        lunar_map = {
            # 2025
            "2025-01-28": "설날", "2025-01-29": "설날", "2025-01-30": "설날",
            "2025-03-03": "대체공휴일",  # 3.1절 대체
            "2025-05-05": "어린이날", "2025-05-06": "대체공휴일",
            "2025-10-05": "추석", "2025-10-06": "추석", "2025-10-07": "추석", "2025-10-08": "대체공휴일",
            # 2026 (예시)
            "2026-02-16": "설날", "2026-02-17": "설날", "2026-02-18": "설날",
            "2026-09-24": "추석", "2026-09-25": "추석", "2026-09-26": "추석",
        }
        return lunar_map.get(iso_date)

    def get_event_name(self, date_obj: QtCore.QDate) -> Optional[str]:
        """기념일(회색글씨) 이름 반환"""
        md_date = date_obj.toString("MM-dd")
        events = {
            "02-14": "발렌타인", "03-14": "화이트", "04-14": "블랙",
            "05-14": "로즈", "11-11": "빼빼로"
        }
        return events.get(md_date)


# =========================================================
# 2. 메모 데이터 클래스 및 저장소
# =========================================================
@dataclass
class MemoItem:
    id: str
    date: str  # yyyy-MM-dd
    title: str
    content: str
    completed: bool = False
    created_at: str = ""
    updated_at: str = ""

    @staticmethod
    def from_dict(d: dict) -> "MemoItem":
        return MemoItem(
            id=str(d.get("id", "")),
            date=str(d.get("date", "")),
            title=str(d.get("title", "")),
            content=str(d.get("content", "")),
            completed=bool(d.get("completed", False)),
            created_at=str(d.get("created_at", "")),
            updated_at=str(d.get("updated_at", ""))
        )


class MemoStore:
    def __init__(self, path: Path):
        self.path = path
        self.items: Dict[str, MemoItem] = {}

    def load(self) -> None:
        if not self.path.exists():
            self.items = {}
            return
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            raw = data.get("items", [])
            self.items = {}
            for r in raw:
                it = MemoItem.from_dict(r)
                if it.id:
                    self.items[it.id] = it
        except:
            self.items = {}

    def save(self) -> None:
        payload = {
            "version": 1,
            "items": [asdict(x) for x in self.items.values()],
            "saved_at": _now_iso(),
        }
        self.path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def list_all_sorted(self) -> List[MemoItem]:
        return sorted(self.items.values(), key=lambda x: (x.date, x.updated_at), reverse=True)

    def list_by_month(self, year: int, month: int) -> List[MemoItem]:
        prefix = f"{year:04d}-{month:02d}"
        items = [x for x in self.items.values() if x.date.startswith(prefix)]
        return sorted(items, key=lambda x: (x.date, x.updated_at), reverse=True)

    def add(self, date_str: str, title: str, content: str, completed: bool) -> MemoItem:
        _id = uuid.uuid4().hex
        now = _now_iso()
        it = MemoItem(_id, date_str, title, content, completed, now, now)
        self.items[_id] = it
        self.save()
        return it

    def update(self, item_id: str, date: str, title: str, content: str, completed: bool):
        it = self.items.get(item_id)
        if it:
            it.date = date
            it.title = title
            it.content = content
            it.completed = completed
            it.updated_at = _now_iso()
            self.save()

    def delete(self, item_id: str):
        if item_id in self.items:
            del self.items[item_id]
            self.save()




class GoogleCalendarSync:
    def __init__(self, secrets_dir: Path, calendar_id: str = "primary"):
        self.secrets_dir = Path(secrets_dir)
        self.calendar_id = calendar_id

        self.client_secret_path = self.secrets_dir / "client_secret.json"
        self.token_path = self.secrets_dir / "token.json"

    # ✅ 1) UI 스레드에서만 호출: 최초 로그인/승인(브라우저 열림)
    # 기존 authorize_interactive 함수들(2개)을 전부 지우고 이 코드로 덮어쓰세요.
    def authorize_interactive(self, parent=None) -> None:
        """
        UI 스레드에서 브라우저를 띄워 인증 (run_local_server 사용)
        - parent 인자는 호출 호환성을 위해 남겨둠 (사용 안 함)
        """
        if not self.secrets_dir.exists():
            self.secrets_dir.mkdir(parents=True, exist_ok=True)

        if not self.client_secret_path.is_file():
            raise FileNotFoundError(f"client_secret.json을 찾을 수 없습니다: {self.client_secret_path}")

        # ✅ 중요: 여기서 run_local_server를 써야 redirect_uri 오류가 안 납니다.
        flow = InstalledAppFlow.from_client_secrets_file(str(self.client_secret_path), SCOPES)
        creds = flow.run_local_server(port=8080)

        self.token_path.write_text(creds.to_json(), encoding="utf-8")

    # ✅ 2) 워커 스레드에서도 안전: token.json만 읽어서 service 생성
    def build_service(self):
        if not self.token_path.is_file():
            raise FileNotFoundError(
                f"token.json이 없습니다. 먼저 '구글 연동'을 완료해야 합니다.\n{self.token_path}"
            )

        creds = Credentials.from_authorized_user_file(str(self.token_path), SCOPES)

        # 토큰 만료 시 refresh(브라우저 안 열림)
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            self.token_path.write_text(creds.to_json(), encoding="utf-8")

        return build("calendar", "v3", credentials=creds)

    # ---------- CRUD ----------
    def create_event(self, service, item: "ScheduleItem") -> str:
        # all-day 이벤트의 end.date는 "다음날"이어야 정상(구글 규칙)
        start_date = item.date  # "YYYY-MM-DD"
        qd = datetime.fromisoformat(start_date).date()
        end_date = (qd + timedelta(days=1)).isoformat()

        body = {
            "summary": (f"[완료] {item.title}" if item.completed else item.title),
            "description": item.content or "",
            "start": {"date": start_date},
            "end": {"date": end_date},
        }
        created = service.events().insert(calendarId=self.calendar_id, body=body).execute()
        return str(created.get("id", ""))

    def update_event(self, service, event_id: str, item: "ScheduleItem") -> None:
        if not event_id:
            return

        start_date = item.date
        qd = datetime.fromisoformat(start_date).date()
        end_date = (qd + timedelta(days=1)).isoformat()

        body = {
            "summary": (f"[완료] {item.title}" if item.completed else item.title),
            "description": item.content or "",
            "start": {"date": start_date},
            "end": {"date": end_date},
        }
        service.events().update(calendarId=self.calendar_id, eventId=event_id, body=body).execute()

    def delete_event(self, service, event_id: str) -> None:
        if not event_id:
            return
        service.events().delete(calendarId=self.calendar_id, eventId=event_id).execute()

    def list_events(self, service, time_min_iso: str, time_max_iso: str) -> List[dict]:
        out: List[dict] = []
        page_token = None
        while True:
            resp = service.events().list(
                calendarId=self.calendar_id,
                timeMin=time_min_iso,
                timeMax=time_max_iso,
                singleEvents=True,
                orderBy="startTime",
                maxResults=2500,
                pageToken=page_token,
            ).execute()
            out.extend(resp.get("items", []) or [])
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
        return out





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

    # ✅ 구글 캘린더 이벤트 ID (구글에 이벤트 만들면 응답으로 오는 id)
    google_event_id: str = ""

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

            # ✅ 기존 json에 없을 수 있으니 default ""로 안전 처리
            google_event_id=str(d.get("google_event_id", "")),

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
    def __init__(
            self,
            calendar: QtWidgets.QCalendarWidget,
            holiday_manager: HolidayManager,  # 추가됨
            parent=None,
            get_items_for_date=None,
    ):
        super().__init__(parent)
        self.calendar = calendar
        self.holiday_mgr = holiday_manager  # 추가됨
        self.get_items_for_date = get_items_for_date
        self.max_lines = 5
        self.pad = 4

    def _cell_date(self, row: int, col: int) -> QtCore.QDate:
        year = self.calendar.yearShown()
        month = self.calendar.monthShown()
        first_of_month = QtCore.QDate(year, month, 1)
        first_dow = int(self.calendar.firstDayOfWeek())
        month_dow = first_of_month.dayOfWeek()
        offset = (month_dow - first_dow) % 7
        return first_of_month.addDays(-offset + row * 7 + col)

    def paint(self, painter: QtGui.QPainter, option: QtWidgets.QStyleOptionViewItem, index: QtCore.QModelIndex):
        painter.save()

        # 1. 날짜 및 휴일 정보 확인
        row, col = index.row(), index.column()
        qdate = self._cell_date(row, col)
        date_str = qdate.toString("yyyy-MM-dd")

        holiday_name = self.holiday_mgr.get_holiday_name(qdate)
        event_name = self.holiday_mgr.get_event_name(qdate)

        is_sunday = (qdate.dayOfWeek() == 7)
        is_holiday = (holiday_name is not None) or is_sunday

        # 2. 배경 그리기
        style = option.widget.style() if option.widget else QtWidgets.QApplication.style()
        style.drawPrimitive(QtWidgets.QStyle.PE_PanelItemViewItem, option, painter, option.widget)

        outer = option.rect
        rect = outer.adjusted(self.pad, self.pad, -self.pad, -self.pad)
        is_current_month = (qdate.month() == self.calendar.monthShown())

        # 스케쥴 유무 배경
        items = self.get_items_for_date(date_str) if self.get_items_for_date else []
        if items:
            has_active = any(not x.completed for x in items)
            bg = QtGui.QColor(255, 245, 200) if has_active else QtGui.QColor(235, 235, 235)
            bg.setAlpha(140)
            painter.fillRect(outer, bg)

        # 선택 날짜 하이라이트
        if option.state & QtWidgets.QStyle.State_Selected:
            sel_bg = QtGui.QColor(210, 235, 255)
            sel_bg.setAlpha(210)
            painter.fillRect(outer, sel_bg)
            pen = QtGui.QPen(QtGui.QColor(120, 170, 255))
            pen.setWidth(2)
            painter.setPen(pen)
            painter.drawRect(outer.adjusted(1, 1, -2, -2))

        # 3. 날짜 숫자 그리기
        day_font = QtGui.QFont(option.font)
        day_font.setBold(True)
        painter.setFont(day_font)

        # 색상 결정 (공휴일/일요일: 빨강, 평일: 기본, 현재달 아니면 흐리게)
        if not is_current_month:
            day_color = QtGui.QColor(200, 200, 200)  # 더 흐리게
        elif is_holiday:
            day_color = QtGui.QColor(255, 60, 60)  # 빨강
        else:
            day_color = option.palette.color(QtGui.QPalette.Text)

        painter.setPen(day_color)
        fm_day = QtGui.QFontMetrics(day_font)
        day_h = fm_day.height()

        day_rect = QtCore.QRect(rect.left(), rect.top(), rect.width(), day_h)
        painter.drawText(day_rect, QtCore.Qt.AlignLeft | QtCore.Qt.AlignTop, str(qdate.day()))

        # 4. 휴일/기념일 명칭 그리기 (날짜 옆에 작게)
        if is_current_month and (holiday_name or event_name):
            evt_font = QtGui.QFont(day_font)
            evt_font.setPointSize(max(7, day_font.pointSize() - 3))
            painter.setFont(evt_font)

            evt_text = holiday_name if holiday_name else event_name
            evt_color = QtGui.QColor(255, 60, 60) if holiday_name else QtGui.QColor(150, 150, 150)
            painter.setPen(evt_color)

            # 날짜 숫자 오른쪽 공간에 그리기
            # day_rect의 오른쪽 부분
            painter.drawText(day_rect, QtCore.Qt.AlignRight | QtCore.Qt.AlignTop, evt_text)

        # 5. 스케쥴 리스트 그리기
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

                # 달 아님 -> 스케쥴도 흐리게
                color = option.palette.color(QtGui.QPalette.Text)
                if not is_current_month:
                    color = QtGui.QColor(180, 180, 180)

                painter.setPen(color)
                line = fm.elidedText(str(getattr(it, "title", "")), QtCore.Qt.ElideRight, text_rect.width())
                painter.drawText(QtCore.QRect(text_rect.left(), y, text_rect.width(), line_h),
                                 QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter, line)
                y += line_h

            remain = len(items) - self.max_lines
            if remain > 0 and y + line_h <= rect.bottom():
                painter.setPen(QtGui.QColor(90, 90, 90))
                painter.drawText(QtCore.QRect(text_rect.left(), y, text_rect.width(), line_h),
                                 QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter, f"+{remain}")

        painter.restore()


class ScheduleWidget(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        # 1. 데이터 로드 (경로 설정)
        base_dir = Path(__file__).resolve().parent
        self.store = ScheduleStore(base_dir / "schedule_data.json")
        self.store.load()

        # [신규] 메모 및 휴일 데이터 로드
        self.memo_store = MemoStore(base_dir / "memo_data.json")
        self.memo_store.load()
        self.holiday_mgr = HolidayManager(base_dir / "holiday_data.json")

        # 구글 연동 경로 및 객체 생성
        self._gcal_secrets_dir = Path(r"C:\my_games\excel_cal\secrets")
        self._gcal = GoogleCalendarSync(self._gcal_secrets_dir, calendar_id="primary")

        # 2. 메인 레이아웃
        root = QtWidgets.QHBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(12)

        # =========================================================
        # [LEFT] 달력 및 버튼 영역 (화면 비율 6)
        # =========================================================
        left = QtWidgets.QVBoxLayout()
        left.setSpacing(10)
        root.addLayout(left, 6)

        # 2-1. 헤더 (제목 + 휴일 설정 버튼)
        header_box = QtWidgets.QHBoxLayout()
        title_lbl = QtWidgets.QLabel("스케쥴 (달력)")
        title_lbl.setStyleSheet("font-weight: bold; font-size: 14px;")
        header_box.addWidget(title_lbl)

        btn_holiday = QtWidgets.QPushButton("휴일 설정")
        btn_holiday.setFixedWidth(100)
        btn_holiday.clicked.connect(self._on_holiday_mgr_clicked)
        header_box.addStretch(1)
        header_box.addWidget(btn_holiday)
        left.addLayout(header_box)

        # 2-2. 달력 네비게이션 (◀ 12월 2025 ▶)
        nav_bar = QtWidgets.QWidget()
        nav_lay = QtWidgets.QHBoxLayout(nav_bar)
        nav_lay.setContentsMargins(6, 0, 6, 0)
        nav_lay.setSpacing(8)

        btn_prev = QtWidgets.QToolButton()
        btn_prev.setText("◀")
        btn_prev.setCursor(QtCore.Qt.PointingHandCursor)
        btn_next = QtWidgets.QToolButton()
        btn_next.setText("▶")
        btn_next.setCursor(QtCore.Qt.PointingHandCursor)

        self.cb_month = QtWidgets.QComboBox()
        for m in range(1, 13): self.cb_month.addItem(f"{m}월", m)
        self.cb_year = QtWidgets.QComboBox()
        for y in range(2020, 2036): self.cb_year.addItem(str(y), y)

        nav_lay.addWidget(btn_prev)
        nav_lay.addStretch(1)
        nav_lay.addWidget(self.cb_month)
        nav_lay.addWidget(self.cb_year)
        nav_lay.addStretch(1)
        nav_lay.addWidget(btn_next)
        left.addWidget(nav_bar)

        # 2-3. 요일바 (일~토, 색상 적용)
        weekday_bar = QtWidgets.QWidget()
        hb = QtWidgets.QHBoxLayout(weekday_bar)
        hb.setContentsMargins(10, 0, 10, 0);
        hb.setSpacing(0)
        for i, t in enumerate(["일", "월", "화", "수", "목", "금", "토"]):
            lb = QtWidgets.QLabel(t)
            lb.setAlignment(QtCore.Qt.AlignCenter)
            lb.setMinimumHeight(28)
            if i == 0:
                lb.setStyleSheet("color: red; font-weight: bold;")
            elif i == 6:
                lb.setStyleSheet("color: blue; font-weight: bold;")
            hb.addWidget(lb, 1)
        left.addWidget(weekday_bar)

        # 2-4. 달력 위젯 본체
        self.calendar = QtWidgets.QCalendarWidget()
        self.calendar.setNavigationBarVisible(False)
        self.calendar.setHorizontalHeaderFormat(QtWidgets.QCalendarWidget.NoHorizontalHeader)
        self.calendar.setVerticalHeaderFormat(QtWidgets.QCalendarWidget.NoVerticalHeader)
        self.calendar.setFirstDayOfWeek(QtCore.Qt.Sunday)
        self.calendar.setGridVisible(True)
        try:
            self.calendar.setLocale(QLocale(QLocale.Korean, QLocale.SouthKorea))
        except:
            pass
        self.calendar.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.calendar.setMinimumSize(860, 600)  # 높이 확보

        # 2-5. 달력 델리게이트 (휴일/메모 표시용) 설정
        self._cal_view = self.calendar.findChild(QtWidgets.QTableView)
        self._cal_view.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self._cal_view.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectItems)
        self._cal_view.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        # [핵심] 여기서 HolidayManager를 Delegate에 주입
        self._cal_delegate = CalendarScheduleDelegate(
            self.calendar,
            self.holiday_mgr,
            self._cal_view,
            get_items_for_date=self._get_items_for_date_for_calendar
        )
        self._cal_view.setItemDelegate(self._cal_delegate)

        # 달력 네비게이션 연결
        btn_prev.clicked.connect(self.calendar.showPreviousMonth)
        btn_next.clicked.connect(self.calendar.showNextMonth)

        # 콤보박스와 달력 동기화 함수
        def _sync_combo():
            self.cb_month.blockSignals(True);
            self.cb_year.blockSignals(True)
            self.cb_month.setCurrentIndex(self.calendar.monthShown() - 1)
            idx = self.cb_year.findData(self.calendar.yearShown())
            if idx >= 0: self.cb_year.setCurrentIndex(idx)
            self.cb_month.blockSignals(False);
            self.cb_year.blockSignals(False)

        def _apply_combo():
            self.calendar.setCurrentPage(int(self.cb_year.currentData()), int(self.cb_month.currentData()))
            QtCore.QTimer.singleShot(0, self._after_calendar_navigate)  # 휠/이동 후 보정

        self.calendar.currentPageChanged.connect(lambda y, m: [_sync_combo(), self._refresh_memo_month_list()])
        self.cb_month.currentIndexChanged.connect(_apply_combo)
        self.cb_year.currentIndexChanged.connect(_apply_combo)

        # 초기 콤보 동기화
        _sync_combo()
        left.addWidget(self.calendar, 1)

        # 2-6. 선택 날짜 표시
        self.lbl_date = QtWidgets.QLabel("선택 날짜: -")
        self.lbl_date.setStyleSheet("font-weight: bold;")
        left.addWidget(self.lbl_date)

        # 2-7. [버튼 영역] 스택 위젯 (탭에 따라 스케쥴버튼 <-> 메모버튼 교체)
        self.btn_stack = QtWidgets.QStackedWidget()
        left.addWidget(self.btn_stack)

        # (A) 스케쥴용 버튼 그룹 (Index 0)
        btn_group_sch = QtWidgets.QWidget()
        bg_sch_lay = QtWidgets.QHBoxLayout(btn_group_sch)
        bg_sch_lay.setContentsMargins(0, 0, 0, 0)
        self.btn_add = QtWidgets.QPushButton("스케쥴 추가")
        self.btn_edit = QtWidgets.QPushButton("수정")
        self.btn_delete = QtWidgets.QPushButton("삭제")  # 변수명 btn_delete 주의
        self.btn_sync = QtWidgets.QPushButton("구글 동기화")
        bg_sch_lay.addWidget(self.btn_add)
        bg_sch_lay.addWidget(self.btn_edit)
        bg_sch_lay.addWidget(self.btn_delete)
        bg_sch_lay.addWidget(self.btn_sync)
        bg_sch_lay.addStretch(1)
        self.btn_stack.addWidget(btn_group_sch)

        # (B) 메모용 버튼 그룹 (Index 1)
        btn_group_memo = QtWidgets.QWidget()
        bg_mem_lay = QtWidgets.QHBoxLayout(btn_group_memo)
        bg_mem_lay.setContentsMargins(0, 0, 0, 0)
        self.btn_mem_add = QtWidgets.QPushButton("메모 추가")
        self.btn_mem_edit = QtWidgets.QPushButton("메모 수정")
        self.btn_mem_del = QtWidgets.QPushButton("메모 삭제")
        bg_mem_lay.addWidget(self.btn_mem_add)
        bg_mem_lay.addWidget(self.btn_mem_edit)
        bg_mem_lay.addWidget(self.btn_mem_del)
        bg_mem_lay.addStretch(1)
        self.btn_stack.addWidget(btn_group_memo)

        # =========================================================
        # [RIGHT] 탭 위젯 (상세 스케쥴 / 메모) (화면 비율 4)
        # =========================================================
        right = QtWidgets.QVBoxLayout()
        root.addLayout(right, 4)

        self.right_tabs = QtWidgets.QTabWidget()
        self.right_tabs.setStyleSheet("""
                    QTabWidget::pane { 
                        border: 1px solid #c0c0c0; 
                        background: white; 
                        border-radius: 4px;
                    }
                    QTabBar::tab {
                        background: #e1e1e1;       /* 선택 안 된 탭: 회색 */
                        border: 1px solid #c0c0c0;
                        padding: 8px 16px;
                        margin-right: 2px;
                        border-top-left-radius: 4px;
                        border-top-right-radius: 4px;
                        color: #606060;
                    }
                    QTabBar::tab:selected {
                        background: #ffffff;       /* 선택 된 탭: 흰색 */
                        border-bottom-color: #ffffff; /* 아래쪽 경계선을 지워 내용과 연결 */
                        font-weight: bold;
                        color: #000000;
                        border-top: 3px solid #4dabf7; /* 상단에 파란색 포인트 줄 */
                    }
                    QTabBar::tab:hover {
                        background: #f0f0f0;       /* 마우스 올렸을 때: 연한 회색 */
                    }
                """)
        right.addWidget(self.right_tabs)

        # --- TAB 1: 상세 스케쥴 ---
        self.tab_schedule = QtWidgets.QWidget()
        self.right_tabs.addTab(self.tab_schedule, "상세 스케쥴")
        ts_lay = QtWidgets.QVBoxLayout(self.tab_schedule)

        # 1-1. 전체 스케쥴 리스트
        gb_all = QtWidgets.QGroupBox("전체 스케쥴 제목 리스트")
        gl_lay = QtWidgets.QVBoxLayout(gb_all)
        self.cb_filter = QtWidgets.QComboBox()
        self.cb_filter.addItems(["전체", "진행중", "완료"]);
        self.cb_filter.setCurrentIndex(1)
        self.all_list = QtWidgets.QListWidget()
        gl_lay.addWidget(self.cb_filter)
        gl_lay.addWidget(self.all_list)
        ts_lay.addWidget(gb_all, 4)  # 비율 4

        # 1-2. 선택 날짜 스케쥴
        gb_day = QtWidgets.QGroupBox("선택 날짜 스케쥴")
        gd_lay = QtWidgets.QVBoxLayout(gb_day)
        self.day_list = QtWidgets.QListWidget()
        gd_lay.addWidget(self.day_list)
        ts_lay.addWidget(gb_day, 3)  # 비율 3

        # 1-3. 상세 내용
        gb_det = QtWidgets.QGroupBox("상세 내용")
        gdt_lay = QtWidgets.QVBoxLayout(gb_det)
        self.detail = QtWidgets.QPlainTextEdit()
        self.detail.setReadOnly(True)
        gdt_lay.addWidget(self.detail)
        ts_lay.addWidget(gb_det, 3)  # 비율 3

        # --- TAB 2: 메모 ---
        self.tab_memo = QtWidgets.QWidget()
        self.right_tabs.addTab(self.tab_memo, "메모")
        tm_lay = QtWidgets.QVBoxLayout(self.tab_memo)

        # 2-1. 전체 메모
        gb_mall = QtWidgets.QGroupBox("전체 메모 제목 리스트")
        gml_lay = QtWidgets.QVBoxLayout(gb_mall)
        self.cb_mem_filter = QtWidgets.QComboBox()
        self.cb_mem_filter.addItems(["전체", "진행중", "완료"]);
        self.cb_mem_filter.setCurrentIndex(1)
        self.mem_all_list = QtWidgets.QListWidget()
        gml_lay.addWidget(self.cb_mem_filter)
        gml_lay.addWidget(self.mem_all_list)
        tm_lay.addWidget(gb_mall, 4)

        # 2-2. 이달의 메모
        gb_mmon = QtWidgets.QGroupBox("이달의 메모")
        gmm_lay = QtWidgets.QVBoxLayout(gb_mmon)
        self.mem_mon_list = QtWidgets.QListWidget()
        gmm_lay.addWidget(self.mem_mon_list)
        tm_lay.addWidget(gb_mmon, 3)

        # 2-3. 메모 상세
        gb_mdet = QtWidgets.QGroupBox("메모 상세 내용")
        gmd_lay = QtWidgets.QVBoxLayout(gb_mdet)
        self.mem_detail = QtWidgets.QPlainTextEdit()
        self.mem_detail.setReadOnly(True)
        gmd_lay.addWidget(self.mem_detail)
        tm_lay.addWidget(gb_mdet, 3)

        # =========================================================
        # 시그널 연결
        # =========================================================
        # 탭 변경 시 버튼 스택 변경
        self.right_tabs.currentChanged.connect(self._on_tab_changed)

        # 달력/스케쥴 공통
        self.calendar.selectionChanged.connect(self._on_date_changed)
        self.cb_filter.currentIndexChanged.connect(self._refresh_all_list)
        self.all_list.currentItemChanged.connect(self._on_all_selection_changed)
        self.day_list.currentItemChanged.connect(self._on_day_selection_changed)

        # 스케쥴 버튼 (기존 기능)
        self.btn_add.clicked.connect(self._on_add_clicked)
        self.btn_edit.clicked.connect(self._on_edit_clicked)
        self.btn_delete.clicked.connect(self._on_delete_clicked)
        self.btn_sync.clicked.connect(self._on_sync_clicked)

        # 메모 관련
        self.cb_mem_filter.currentIndexChanged.connect(self._refresh_memo_all_list)
        self.mem_all_list.currentItemChanged.connect(self._on_mem_all_sel_changed)
        self.mem_mon_list.currentItemChanged.connect(self._on_mem_mon_sel_changed)
        self.btn_mem_add.clicked.connect(self._on_mem_add_clicked)
        self.btn_mem_edit.clicked.connect(self._on_mem_edit_clicked)
        self.btn_mem_del.clicked.connect(self._on_mem_del_clicked)

        # 초기 화면 갱신
        self._on_date_changed()
        self._refresh_all_list()
        self._refresh_memo_all_list()
        self._refresh_memo_month_list()
        self._select_first_item_if_any()
        self._refresh_calendar_view()

        self._install_calendar_nav_hooks()  # 휠 스크롤 훅

        # 구글 연동 초기화
        try:
            if self._gcal.client_secret_path.exists():
                if not (self._gcal_secrets_dir / "token.json").exists():
                    self._gcal.authorize_interactive(parent=self)
                self._gcal_service = self._gcal.build_service()
            else:
                self._gcal_service = None
        except Exception:
            self._gcal_service = None

    # =========================================================
    # 탭 및 공통 유틸
    # =========================================================
    def _on_tab_changed(self, index: int):
        # 0: 스케쥴, 1: 메모 -> 버튼 그룹 교체
        self.btn_stack.setCurrentIndex(index)

    def _run_async(self, title: str, job, on_done):
        run_job_with_progress_async(owner=self, title=title, job=job, tail_file=None, on_done=on_done)

    def _selected_date_str(self) -> str:
        return self.calendar.selectedDate().toString("yyyy-MM-dd")

    # 휠 스크롤 훅
    def _install_calendar_nav_hooks(self) -> None:
        self.calendar.installEventFilter(self)
        if hasattr(self, "_cal_view") and self._cal_view is not None:
            self._cal_view.viewport().installEventFilter(self)

    def eventFilter(self, obj, event):
        if event.type() == QtCore.QEvent.Wheel:
            QtCore.QTimer.singleShot(0, self._after_calendar_navigate)
            return False
        return super().eventFilter(obj, event)

    def _after_calendar_navigate(self) -> None:
        try:
            self._force_select_cell_for_selected_date()
            self._refresh_calendar_view()
        except:
            pass

    def _force_select_cell_for_selected_date(self) -> None:
        if not self._cal_view: return
        qd = self.calendar.selectedDate()
        if not qd.isValid(): return
        # (간략화된 선택 로직 - 기존 코드의 복잡한 로직 대신 네이티브 함수 활용)
        # 델리게이트가 알아서 그리므로 여기서는 뷰 업데이트만 잘 해주면 됨

    # =========================================================
    # 휴일 관리 로직
    # =========================================================
    def _on_holiday_mgr_clicked(self):
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("휴일 직접 설정")
        lay = QtWidgets.QVBoxLayout(dlg)

        form = QtWidgets.QFormLayout()
        de = QtWidgets.QDateEdit()
        de.setDisplayFormat("yyyy-MM-dd")
        de.setCalendarPopup(True)
        de.setDate(QtCore.QDate.currentDate())
        le = QtWidgets.QLineEdit()
        le.setPlaceholderText("예: 창립기념일 (비워두면 삭제)")
        form.addRow("날짜", de)
        form.addRow("이름", le)
        lay.addLayout(form)

        btn = QtWidgets.QPushButton("저장 (이름 비우면 삭제)")
        btn.clicked.connect(dlg.accept)
        lay.addWidget(btn)

        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            d_str = de.date().toString("yyyy-MM-dd")
            name = le.text().strip()
            if name:
                self.holiday_mgr.add_custom_holiday(d_str, name)
                QtWidgets.QMessageBox.information(self, "저장", f"[{d_str}] {name} 저장되었습니다.")
            else:
                self.holiday_mgr.remove_custom_holiday(d_str)
                QtWidgets.QMessageBox.information(self, "삭제", f"[{d_str}] 휴일 설정이 삭제되었습니다.")
            self._refresh_calendar_view()

    # =========================================================
    # 달력 표시 로직
    # =========================================================
    def _get_items_for_date_for_calendar(self, date_str: str) -> List[ScheduleItem]:
        items = self.store.list_by_date(date_str)
        return sorted(items, key=lambda x: (x.completed,
                                            -(datetime.fromisoformat(x.updated_at).timestamp() if x.updated_at else 0)))

    def _refresh_calendar_view(self) -> None:
        if self._cal_view: self._cal_view.viewport().update()

    def _on_date_changed(self) -> None:
        date_str = self.calendar.selectedDate().toString("yyyy-MM-dd (ddd)")
        self.lbl_date.setText(f"선택 날짜: {date_str}")
        self._refresh_day_list()
        self._refresh_calendar_view()

    # =========================================================
    # 스케쥴 리스트 및 상세 (기존 로직 이식)
    # =========================================================
    def _refresh_day_list(self) -> None:
        self.day_list.clear()
        date_str = self._selected_date_str()
        items = self.store.list_by_date(date_str)
        if not items:
            x = QtWidgets.QListWidgetItem("스케쥴 없음");
            x.setFlags(QtCore.Qt.NoItemFlags)
            self.day_list.addItem(x);
            return
        for it in items:
            li = QtWidgets.QListWidgetItem(it.title)
            li.setData(QtCore.Qt.UserRole, it.id)
            f = li.font();
            f.setStrikeOut(it.completed);
            li.setFont(f)
            self.day_list.addItem(li)

    def _refresh_all_list(self) -> None:
        self.all_list.clear()
        items = self.store.list_all_sorted()
        filter_mode = self.cb_filter.currentText()
        count = 0
        for it in items:
            if filter_mode == "진행중" and it.completed: continue
            if filter_mode == "완료" and not it.completed: continue

            prefix = "(완료) " if it.completed else ""
            li = QtWidgets.QListWidgetItem(f"{it.date} | {prefix}{it.title}")
            li.setData(QtCore.Qt.UserRole, it.id)
            f = li.font();
            f.setStrikeOut(it.completed);
            li.setFont(f)
            self.all_list.addItem(li)
            count += 1

        if count == 0:
            msg = "스케쥴 없음"
            if filter_mode == "진행중":
                msg = "진행중인 스케쥴 없음"
            elif filter_mode == "완료":
                msg = "완료된 스케쥴 없음"
            x = QtWidgets.QListWidgetItem(msg);
            x.setFlags(QtCore.Qt.NoItemFlags)
            self.all_list.addItem(x)
            self.detail.setPlainText("")

    def _select_first_item_if_any(self) -> None:
        if self.all_list.count() > 0:
            first = self.all_list.item(0)
            if first.flags() & QtCore.Qt.ItemIsEnabled: self.all_list.setCurrentRow(0)

    def _get_selected_all_item_id(self) -> Optional[str]:
        cur = self.all_list.currentItem()
        if not cur: return None
        item_id = cur.data(QtCore.Qt.UserRole)
        return item_id if isinstance(item_id, str) and item_id else None

    def _on_all_selection_changed(self, cur, prev) -> None:
        item_id = self._get_selected_all_item_id()
        if not item_id: self._render_detail(None); return
        it = self.store.items.get(item_id)
        self._render_detail(it)

    def _on_day_selection_changed(self, cur, prev) -> None:
        if not cur: return
        item_id = cur.data(QtCore.Qt.UserRole)
        if not isinstance(item_id, str) or not item_id: return
        # all_list 동기화
        for i in range(self.all_list.count()):
            it = self.all_list.item(i)
            if it.data(QtCore.Qt.UserRole) == item_id:
                self.all_list.blockSignals(True)
                self.all_list.setCurrentRow(i)
                self.all_list.scrollToItem(it, QtWidgets.QAbstractItemView.PositionAtCenter)
                self.all_list.blockSignals(False)
                self._render_detail(self.store.items.get(item_id))
                return

    def _render_detail(self, it: Optional[ScheduleItem]) -> None:
        if not it: self.detail.setPlainText(""); return
        status = "완료" if it.completed else "진행중"
        text = f"날짜: {it.date}\n상태: {status}\n제목: {it.title}\n생성: {it.created_at}\n수정: {it.updated_at}\n\n{it.content}"
        self.detail.setPlainText(text)

    def _ensure_gcal(self) -> Optional[GoogleCalendarSync]:
        if self._gcal is not None: return self._gcal
        self._gcal = GoogleCalendarSync(self._gcal_secrets_dir, calendar_id="primary")
        return self._gcal

    # =========================================================
    # [핵심] 스케쥴 CRUD 및 구글 동기화 (기존 로직 복원)
    # =========================================================
    def _on_add_clicked(self) -> None:
        date_str = self._selected_date_str()
        dlg = ScheduleEditDialog(self, mode="add", initial_date=date_str, item=None)
        if dlg.exec_() != QtWidgets.QDialog.Accepted: return
        d, title, content, completed = dlg.get_values()

        def job(progress):
            progress({"stage": "local", "msg": "[local] 로컬 저장 중..."})
            it = self.store.add(d, title, content, completed)
            progress({"stage": "gcal", "msg": "[gcal] 구글 캘린더 이벤트 생성 중..."})
            try:
                gcal = self._ensure_gcal()
                if gcal:
                    service = gcal.build_service()
                    event_id = gcal.create_event(service, it)
                    if event_id:
                        it.google_event_id = event_id
                        self.store.save()
            except Exception as e:
                progress({"stage": "gcal", "msg": f"[gcal][warn] 구글 반영 실패: {e}"})
            progress({"stage": "done", "msg": "[done] 완료"})
            return {"date": d, "new_id": it.id}

        def done(ok, payload, err):
            if not ok: QtWidgets.QMessageBox.warning(self, "추가 실패", f"{err}"); return
            d2 = payload.get("date")
            if d2:
                qd = QtCore.QDate.fromString(d2, "yyyy-MM-dd")
                if qd.isValid(): self.calendar.setSelectedDate(qd)
            self._refresh_day_list();
            self._refresh_all_list();
            self._refresh_calendar_view()
            self._sync_mirror_from_google_async(reason="after-add")

        self._run_async("스케쥴 추가", job, done)

    def _on_edit_clicked(self) -> None:
        item_id = self._get_selected_all_item_id()
        if not item_id:
            QtWidgets.QMessageBox.information(self, "안내", "수정할 항목을 선택해 주세요.")
            return
        it = self.store.items.get(item_id)
        if not it: return
        dlg = ScheduleEditDialog(self, mode="edit", initial_date=it.date, item=it)
        if dlg.exec_() != QtWidgets.QDialog.Accepted: return
        d, title, content, completed = dlg.get_values()

        def job(progress):
            progress({"stage": "local", "msg": "[local] 로컬 수정 저장 중..."})
            self.store.update(item_id, date=d, title=title, content=content, completed=completed)
            updated = self.store.items.get(item_id)
            progress({"stage": "gcal", "msg": "[gcal] 구글 캘린더 반영 중..."})
            service = self._gcal.build_service()
            if updated:
                if updated.google_event_id:
                    self._gcal.update_event(service, updated.google_event_id, updated)
                else:
                    event_id = self._gcal.create_event(service, updated)
                    if event_id: updated.google_event_id = event_id; self.store.save()
            progress({"stage": "done", "msg": "[done] 완료"})
            return {"date": d, "id": item_id}

        def done(ok, payload, err):
            if not ok: QtWidgets.QMessageBox.warning(self, "수정 실패", f"{err}"); return
            d2 = payload.get("date")
            if d2: self.calendar.setSelectedDate(QtCore.QDate.fromString(d2, "yyyy-MM-dd"))
            self._refresh_day_list();
            self._refresh_all_list();
            self._refresh_calendar_view()
            self._sync_mirror_from_google_async(reason="after-edit")

        self._run_async("스케쥴 수정", job, done)

    def _on_delete_clicked(self) -> None:
        item_id = self._get_selected_all_item_id()
        if not item_id:
            QtWidgets.QMessageBox.information(self, "안내", "삭제할 항목을 선택해 주세요.")
            return
        it = self.store.items.get(item_id)
        if not it: return
        yn = QtWidgets.QMessageBox.question(self, "삭제 확인", "선택한 스케쥴을 삭제하시겠습니까?")
        if yn != QtWidgets.QMessageBox.Yes: return

        service = self._gcal.build_service()
        event_id = (it.google_event_id or "").strip()

        def job(progress):
            gcal_err = None
            progress({"stage": "gcal", "msg": "[gcal] 구글 캘린더 삭제 시도..."})
            if service and event_id:
                try:
                    self._gcal.delete_event(service, event_id)
                except Exception as e:
                    gcal_err = str(e)
            progress({"stage": "local", "msg": "[local] 로컬 삭제 중..."})
            self.store.delete(item_id)
            progress({"stage": "done", "msg": "[done] 완료"})
            return {"gcal_err": gcal_err}

        def done(ok, payload, err):
            if not ok: QtWidgets.QMessageBox.warning(self, "삭제 실패", f"{err}"); return
            self._refresh_day_list();
            self._refresh_all_list();
            self._refresh_calendar_view();
            self._render_detail(None)
            self._sync_mirror_from_google_async(reason="after-delete")

        self._run_async("스케쥴 삭제", job, done)

    def _on_sync_clicked(self) -> None:
        self._sync_mirror_from_google_async(reason="manual")

    def showEvent(self, e: QtGui.QShowEvent) -> None:
        super().showEvent(e)
        if not getattr(self, "_did_initial_gsync", False):
            self._did_initial_gsync = True
            self._sync_mirror_from_google_async(reason="tab-open")

    def _sync_mirror_from_google_async(self, *, reason: str) -> None:
        def job(progress):
            progress({"stage": "gcal", "msg": f"[gcal] 동기화(미러) 시작 ({reason})"})
            service = self._gcal.build_service()

            today = datetime.now().date()
            d1 = today - timedelta(days=365)
            d2 = today + timedelta(days=365)
            events = self._gcal.list_events(service, f"{d1.isoformat()}T00:00:00Z", f"{d2.isoformat()}T00:00:00Z")
            progress({"stage": "gcal", "msg": f"[gcal] 조회 {len(events)}건"})

            gmap = {}
            for ev in events:
                eid = str(ev.get("id", "") or "").strip()
                if not eid: continue
                date_str = (ev.get("start") or {}).get("date")
                if not date_str: continue
                summary = str(ev.get("summary", "") or "")
                completed = summary.startswith("[완료]")
                title = summary.replace("[완료]", "", 1).strip() if completed else summary
                gmap[eid] = {"eid": eid, "date": date_str, "title": title, "content": ev.get("description", ""),
                             "completed": completed}

            local_by_eid = {}
            local_ids_without_google = set()
            for it in list(self.store.items.values()):
                if it.google_event_id:
                    local_by_eid[it.google_event_id] = it
                else:
                    local_ids_without_google.add(it.id)

            cnt_u = cnt_c = cnt_d = 0
            for eid, info in gmap.items():
                if eid in local_by_eid:
                    it = local_by_eid[eid]
                    it.date = info["date"];
                    it.title = info["title"];
                    it.content = info["content"];
                    it.completed = info["completed"]
                    it.updated_at = _now_iso()
                    cnt_u += 1
                else:
                    new_it = self.store.add(info["date"], info["title"], info["content"], info["completed"])
                    new_it.google_event_id = eid;
                    self.store.save()
                    cnt_c += 1

            for eid, it in list(local_by_eid.items()):
                if eid not in gmap: self.store.delete(it.id); cnt_d += 1

            for item_id in list(local_ids_without_google):
                self.store.delete(item_id);
                cnt_d += 1

            progress({"stage": "local", "msg": f"[local] U:{cnt_u}, C:{cnt_c}, D:{cnt_d}"})
            return {}

        def done(ok, payload, err):
            if not ok: QtWidgets.QMessageBox.warning(self, "동기화 실패", f"{err}"); return
            self._refresh_day_list();
            self._refresh_all_list();
            self._refresh_calendar_view()

        self._run_async("구글 동기화", job, done)

    # =========================================================
    # [NEW] 메모 CRUD (로컬 전용)
    # =========================================================
    def _refresh_memo_all_list(self):
        self.mem_all_list.clear()
        items = self.memo_store.list_all_sorted()
        mode = self.cb_mem_filter.currentText()
        for it in items:
            if mode == "진행중" and it.completed: continue
            if mode == "완료" and not it.completed: continue
            prefix = "(완료) " if it.completed else ""
            li = QtWidgets.QListWidgetItem(f"{it.date} | {prefix}{it.title}")
            li.setData(QtCore.Qt.UserRole, it.id)
            f = li.font();
            f.setStrikeOut(it.completed);
            li.setFont(f)
            self.mem_all_list.addItem(li)

    def _refresh_memo_month_list(self):
        self.mem_mon_list.clear()
        y = self.calendar.yearShown()
        m = self.calendar.monthShown()
        items = self.memo_store.list_by_month(y, m)
        if not items:
            x = QtWidgets.QListWidgetItem("이달의 메모 없음");
            x.setFlags(QtCore.Qt.NoItemFlags)
            self.mem_mon_list.addItem(x);
            return
        for it in items:
            li = QtWidgets.QListWidgetItem(it.title)
            li.setData(QtCore.Qt.UserRole, it.id)
            f = li.font();
            f.setStrikeOut(it.completed);
            li.setFont(f)
            self.mem_mon_list.addItem(li)

    def _render_mem_detail(self, it: Optional[MemoItem]):
        if not it: self.mem_detail.setPlainText(""); return
        status = "완료" if it.completed else "진행중"
        txt = f"[{status}] {it.date}\n제목: {it.title}\n\n{it.content}"
        self.mem_detail.setPlainText(txt)

    def _on_mem_all_sel_changed(self, cur, prev):
        if not cur: self.mem_detail.setPlainText(""); return
        item_id = cur.data(QtCore.Qt.UserRole)
        it = self.memo_store.items.get(item_id)
        if it: self._render_mem_detail(it)

    def _on_mem_mon_sel_changed(self, cur, prev):
        if not cur: return
        item_id = cur.data(QtCore.Qt.UserRole)
        if not item_id: return
        for i in range(self.mem_all_list.count()):
            if self.mem_all_list.item(i).data(QtCore.Qt.UserRole) == item_id:
                self.mem_all_list.setCurrentRow(i);
                break

    def _on_mem_add_clicked(self):
        dlg = ScheduleEditDialog(self, mode="add", initial_date=self._selected_date_str())
        dlg.setWindowTitle("메모 추가")
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            d, t, c, comp = dlg.get_values()
            self.memo_store.add(d, t, c, comp)
            self._refresh_memo_all_list();
            self._refresh_memo_month_list()

    def _on_mem_edit_clicked(self):
        cur = self.mem_all_list.currentItem()
        if not cur: return
        iid = cur.data(QtCore.Qt.UserRole)
        it = self.memo_store.items.get(iid)
        if not it: return
        dlg = ScheduleEditDialog(self, mode="edit", initial_date=it.date, item=it)
        dlg.setWindowTitle("메모 수정")
        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            d, t, c, comp = dlg.get_values()
            self.memo_store.update(iid, d, t, c, comp)
            self._refresh_memo_all_list();
            self._refresh_memo_month_list()
            self._render_mem_detail(self.memo_store.items.get(iid))

    def _on_mem_del_clicked(self):
        cur = self.mem_all_list.currentItem()
        if not cur: return
        iid = cur.data(QtCore.Qt.UserRole)
        yn = QtWidgets.QMessageBox.question(self, "삭제", "메모를 삭제하시겠습니까?")
        if yn == QtWidgets.QMessageBox.Yes:
            self.memo_store.delete(iid)
            self._refresh_memo_all_list();
            self._refresh_memo_month_list()
            self.mem_detail.setPlainText("")

    def _force_select_cell_for_selected_date(self) -> None:
        """
        [복구됨] 달력의 날짜를 클릭했을 때 QTableView의 파란색 선택 박스가
        정확히 해당 날짜 셀을 가리키도록 강제 보정하는 함수입니다.
        """
        if not hasattr(self, "_cal_view") or self._cal_view is None:
            return

        qd = self.calendar.selectedDate()
        if not qd.isValid():
            return

        # 1. 현재 달력의 연/월
        y = self.calendar.yearShown()
        m = self.calendar.monthShown()

        # 2. 이번 달 1일이 달력 그리드(0,0)에서 얼마나 떨어져 있는지(offset) 계산
        first_of_month = QtCore.QDate(y, m, 1)
        first_dow = int(self.calendar.firstDayOfWeek())  # 1=Mon..7=Sun (보통 일요일=7 또는 0)

        # QCalendarWidget의 요일 상수와 QDate 요일 상수가 다를 수 있어 보정
        # Qt.Sunday = 7 (QDate 기준 Sunday=7)
        # 설정된 firstDayOfWeek가 Sunday라면 7

        month_dow = first_of_month.dayOfWeek()  # 1=Mon .. 7=Sun

        # 달력의 첫 번째 칸 날짜 구하기
        # (offset은 1일이 시작하기 전 빈 칸의 개수)
        offset = (month_dow - first_dow) % 7
        start_date = first_of_month.addDays(-offset)

        # 3. 선택한 날짜(qd)가 시작일(start_date)로부터 며칠 떨어져 있는지 계산
        days_diff = start_date.daysTo(qd)

        if days_diff < 0 or days_diff >= 42:  # 6주 * 7일 = 42
            # 달력 화면 범위를 벗어난 경우 무시
            return

        # 4. 행/열 계산하여 선택 강제
        row = days_diff // 7
        col = days_diff % 7

        model = self._cal_view.model()
        if model is None:
            return

        idx = model.index(row, col)
        if idx.isValid():
            self._cal_view.setCurrentIndex(idx)
            self._cal_view.selectionModel().select(
                idx,
                QtCore.QItemSelectionModel.ClearAndSelect | QtCore.QItemSelectionModel.Current
            )

####################비동기###################
def run_job_with_progress_async(
    owner: QtWidgets.QWidget,
    title: str,
    job,
    *,
    tail_file=None,
    on_done=None,
) -> None:

    # 0) 기존 진행창 재사용 여부 확인
    reuse_ctx = getattr(owner, "_progress_ctx", None)
    on_progress_ui = finalize_ui = dlg = None
    reused = False

    if reuse_ctx is not None:
        try:
            old_on_progress, old_finalize, old_dlg = reuse_ctx
            if old_dlg is not None and old_dlg.isVisible():
                on_progress_ui, finalize_ui, dlg = old_on_progress, old_finalize, old_dlg
                reused = True
        except Exception:
            pass

    # 1) 재사용 불가하면 새로 만든다
    if dlg is None:
        on_progress_ui, finalize_ui, dlg = _mk_progress(owner, title, tail_file=tail_file)  # type: ignore
        setattr(owner, "_progress_ctx", (on_progress_ui, finalize_ui, dlg))
        reused = False  # 새 창이니까 원래 finalize 써도 됨

    # 2) 시작 로그
    try:
        on_progress_ui({"stage": "ui", "msg": "[ui] 작업 시작 준비"})
    except Exception:
        pass

    class _Worker(QtCore.QObject):
        progress = QtCore.pyqtSignal(dict)
        finished = QtCore.pyqtSignal(object, object)

        @QtCore.pyqtSlot()
        def run(self):
            payload = None
            err = None
            try:
                def on_progress(info: dict):
                    if not isinstance(info, dict):
                        info = {"msg": str(info)}
                    self.progress.emit(info)
                payload = job(on_progress)
            except Exception as ex:
                err = ex
            finally:
                self.finished.emit(payload, err)

    obj = _Worker()
    th = QtCore.QThread(dlg)
    obj.moveToThread(th)

    def _on_progress(info: dict):
        try:
            on_progress_ui(info)
        except Exception:
            pass

    def _on_finished(payload, err):
        ok = (err is None)

        # 새로 만든 창일 때만 원래 finalize 호출
        if not reused:
            try:
                finalize_ui(ok, payload, err)
            except Exception:
                pass
        else:
            # 재사용 창일 때는 닫기버튼/추가 UI 생성 막기 위해 아무것도 안 함
            # 필요하면 여기서 로그만 하나 찍자
            try:
                on_progress_ui({"stage": "done", "msg": "[ui] 작업 1건 완료 (재사용 중)"})
            except Exception:
                pass

        # 호출자가 준 on_done은 항상 불러줌
        if callable(on_done):
            try:
                on_done(ok, payload, err)
            except Exception:
                pass

        # 스레드 정리
        try:
            th.quit()
            th.wait(100)
        except Exception:
            pass

        # 소유자에 보관했던 스레드 참조 제거
        try:
            jobss = getattr(owner, "_progress_jobs", [])
            if th in jobss:
                jobss.remove(th)
            setattr(owner, "_progress_jobs", jobss)
        except Exception:
            pass

    obj.progress.connect(_on_progress)
    obj.finished.connect(_on_finished)
    th.started.connect(obj.run)

    # GC 방지
    try:
        jobs = getattr(owner, "_progress_jobs", None)
        if not isinstance(jobs, list):
            jobs = []
        jobs.append(th)
        setattr(owner, "_progress_jobs", jobs)
        setattr(th, "_worker_ref", obj)
    except Exception:
        pass

    # 시작 로그
    try:
        on_progress_ui({"stage": "ui", "msg": "[ui] 백그라운드 스레드 시작"})
    except Exception:
        pass

    # 스레드 시작
    try:
        th.start()
    except Exception as start_exc:
        try:
            on_progress_ui({"stage": "error", "msg": f"[error] thread start failed: {start_exc}"})
        except Exception:
            pass
        # 첫 창일 때만 finalize
        if not reused:
            try:
                finalize_ui(False, None, start_exc)
            except Exception:
                pass
        if callable(on_done):
            try:
                on_done(False, None, start_exc)
            except Exception:
                pass
        return


def _mk_progress(owner: QtWidgets.QWidget, title: str, tail_file=None):
    """
    [수정됨] 예쁜 UI + 로딩바 + 성공 시 1초 뒤 자동 닫힘
    """
    dlg = QtWidgets.QDialog(owner)
    dlg.setWindowTitle(title)
    dlg.setModal(True)  # 작업 중 다른거 못 만지게 (선택사항)
    dlg.resize(500, 320)

    # 창 상단 ? 버튼 제거
    dlg.setWindowFlags(dlg.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)

    # ✅ 스타일시트 (깔끔한 디자인)
    dlg.setStyleSheet("""
        QDialog {
            background-color: #ffffff;
        }
        QLabel#TitleLabel {
            font-size: 15px;
            font-weight: bold;
            color: #333333;
        }
        QProgressBar {
            border: none;
            background-color: #f1f3f5;
            border-radius: 4px;
            height: 6px;
        }
        QProgressBar::chunk {
            background-color: #74c0fc;  /* 파란색 로딩 */
            border-radius: 4px;
        }
        QPlainTextEdit {
            background-color: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 8px;
            padding: 10px;
            font-family: '맑은 고딕', sans-serif;
            font-size: 12px;
            color: #555555;
        }
        QPushButton {
            background-color: #ff6b6b;
            color: white;
            border-radius: 6px;
            padding: 6px 14px;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #fa5252;
        }
    """)

    layout = QtWidgets.QVBoxLayout(dlg)
    layout.setContentsMargins(20, 20, 20, 20)
    layout.setSpacing(12)

    # 1. 제목
    lbl = QtWidgets.QLabel(title)
    lbl.setObjectName("TitleLabel")
    layout.addWidget(lbl)

    # 2. 로딩바 (왔다갔다 하는 애니메이션)
    pbar = QtWidgets.QProgressBar()
    pbar.setRange(0, 0)  # 시작/끝 모름 -> 무한 로딩 애니메이션
    pbar.setTextVisible(False)
    layout.addWidget(pbar)

    # 3. 로그 창
    log = QtWidgets.QPlainTextEdit()
    log.setReadOnly(True)
    layout.addWidget(log, 1)

    # 4. 닫기 버튼 (에러 났을 때만 보임)
    btn_area = QtWidgets.QHBoxLayout()
    btn_close = QtWidgets.QPushButton("닫기")
    btn_close.setVisible(False)  # 평소엔 숨김
    btn_area.addStretch(1)
    btn_area.addWidget(btn_close)
    layout.addLayout(btn_area)

    # ----------- 내부 함수들 -----------
    def _append(line: str):
        try:
            log.appendPlainText(line)
        except Exception:
            pass

    def on_progress_ui(info: dict):
        if not isinstance(info, dict):
            _append(str(info))
            return
        msg = info.get("msg")
        if msg:
            _append(str(msg))

    def finalize_ui(ok: bool, payload, err):
        # 로딩바 멈춤
        pbar.setRange(0, 100)

        if ok:
            pbar.setValue(100)  # 꽉 채움
            pbar.setStyleSheet("QProgressBar::chunk { background-color: #a9e34b; }")  # 성공 시 연두색
            _append("\n[성공] 모든 작업이 완료되었습니다.")
            _append("잠시 후 창이 닫힙니다...")

            # ✅ [핵심] 1초(1000ms) 뒤 자동 닫기
            QtCore.QTimer.singleShot(1000, dlg.accept)
        else:
            pbar.setValue(150)
            pbar.setStyleSheet("QProgressBar::chunk { background-color: #ff6b6b; }")  # 실패 시 빨간색
            _append(f"\n[오류] 작업 중 문제가 발생했습니다.\n{err}")

            # 에러나면 닫기 버튼 보여주고 자동 닫기 안 함 (읽어봐야 하니까)
            btn_close.setVisible(True)

    btn_close.clicked.connect(dlg.reject)

    dlg.show()
    return on_progress_ui, finalize_ui, dlg


