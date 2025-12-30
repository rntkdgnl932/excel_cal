# excel_ui.py
# 탭 통합 메인 UI: 엑셀 계산기 + 송장 + 스케줄 + [비동기] 타임클락 + 업데이트

import sys
import subprocess
from pathlib import Path
from datetime import datetime
import traceback
from PyQt5 import QtWidgets, QtCore, QtGui

# -------------------------------------------------------------------------
# [매우 중요] 경로 설정 (import보다 무조건 위에 있어야 함)
# -------------------------------------------------------------------------
# 타임클락 프로젝트가 있는 실제 폴더
TIMECLOCK_ROOT = Path(r"C:\my_games\timeclock")
HAS_TIMECLOCK_PATH = False

if TIMECLOCK_ROOT.exists():
    # 1. 'import timeclock.db' 가 가능하도록 상위 폴더(C:\my_games) 추가
    parent_dir = str(TIMECLOCK_ROOT.parent)
    if parent_dir not in sys.path:
        sys.path.append(parent_dir)

    # 2. 혹시 모를 내부 직접 import를 위해 루트도 추가
    if str(TIMECLOCK_ROOT) not in sys.path:
        sys.path.append(str(TIMECLOCK_ROOT))

    HAS_TIMECLOCK_PATH = True
else:
    print(f"⚠️ 경고: Timeclock 경로를 찾을 수 없습니다: {TIMECLOCK_ROOT}")

# -------------------------------------------------------------------------
# 모듈 임포트 (경로 설정이 끝난 후에 해야 오류가 안 남)
# -------------------------------------------------------------------------
try:
    from excel_cal_ui import ExcelCalWindow  # 기존 부가세/3종 엑셀 UI
    from read_excel import ReadInvoiceWidget  # 송장 읽기 탭
    from schedule import ScheduleWidget  # 스케쥴 탭
except ImportError as e:
    print(f"기본 모듈 임포트 실패: {e}")

# Timeclock 모듈 안전 임포트
HAS_TIMECLOCK = False
try:
    if HAS_TIMECLOCK_PATH:
        # 경로가 sys.path에 있으므로 이제 import가 가능합니다.
        from timeclock.db import DB
        from timeclock.settings import DB_PATH, DEFAULT_OWNER_USER, DEFAULT_OWNER_PASS
        from timeclock.ui.owner_page import OwnerPage
        from timeclock.ui.login_page import Session
        from timeclock import sync_manager

        HAS_TIMECLOCK = True
except ImportError as e:
    print(f"⚠️ Timeclock 모듈 임포트 실패: {e}")
    # 혹시 모듈 내부 경로 문제일 수 있으니 상세 정보 출력
    traceback.print_exc()


# =========================================================================
# [New] 비동기 작업용 워커 클래스 (로딩 & 배지 체크)
# =========================================================================
class TimeclockSyncWorker(QtCore.QThread):
    """DB 다운로드 및 동기화를 백그라운드에서 수행"""
    finished_sig = QtCore.pyqtSignal(bool, str)  # 성공여부, 메시지

    def run(self):
        try:
            if not HAS_TIMECLOCK:
                self.finished_sig.emit(False, "모듈 없음")
                return

            # 구글 드라이브에서 최신 DB 다운로드
            ok, msg = sync_manager.download_latest_db()
            self.finished_sig.emit(ok, msg)
        except Exception as e:
            self.finished_sig.emit(False, str(e))


class BadgeCheckWorker(QtCore.QThread):
    """
    백그라운드에서 주기적으로 DB를 확인하여 대기 건수를 체크
    (주의: 메인 DB 파일을 건드리지 않고, 임시로 체크하거나 DB 연결을 가볍게 사용)
    """
    badge_updated = QtCore.pyqtSignal(int)  # 대기 건수 합계

    def run(self):
        if not HAS_TIMECLOCK or not DB_PATH.exists():
            return

        try:
            # 단순히 현재 로컬 DB 기준으로만 체크 (서버 다운로드는 너무 무거움)
            # 1. DB 연결 (읽기 전용)
            temp_db = DB(DB_PATH)
            counts = temp_db.get_pending_counts()  # {work:0, dispute:0, signup:0}
            total = sum(counts.values())
            temp_db.close()

            self.badge_updated.emit(total)

        except Exception:
            pass


# =========================================================================
# [New] 로딩 오버레이 위젯
# =========================================================================
class LoadingOverlay(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(QtCore.Qt.WA_TransparentForMouseEvents, False)  # 마우스 막기
        self.setStyleSheet("background-color: rgba(255, 255, 255, 180);")  # 반투명 흰색

        layout = QtWidgets.QVBoxLayout(self)
        layout.setAlignment(QtCore.Qt.AlignCenter)

        # 로딩 아이콘 (텍스트나 GIF)
        self.lbl_icon = QtWidgets.QLabel("⏳")
        self.lbl_icon.setStyleSheet("font-size: 40px; background: transparent;")
        self.lbl_icon.setAlignment(QtCore.Qt.AlignCenter)

        self.lbl_text = QtWidgets.QLabel("데이터 동기화 중...")
        self.lbl_text.setStyleSheet("font-size: 16px; font-weight: bold; color: #333; background: transparent;")
        self.lbl_text.setAlignment(QtCore.Qt.AlignCenter)

        layout.addWidget(self.lbl_icon)
        layout.addWidget(self.lbl_text)

        self.hide()

    def show_loading(self):
        if self.parent():
            self.resize(self.parent().size())
        self.show()
        self.raise_()

    def hide_loading(self):
        self.hide()


# =========================================================================
# Git 업데이트 위젯 (기존 코드 복원)
# =========================================================================
class UpdateWidget(QtWidgets.QWidget):
    """
    Git 저장소 기준으로 현재 githash 확인 + git pull 실행하는 탭.
    """

    def __init__(self, repo_root: Path, parent=None):
        super().__init__(parent)
        self.repo_root = repo_root

        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # 상단 타이틀
        lbl_title = QtWidgets.QLabel("엑셀 도구 Git 업데이트")
        lbl_title.setObjectName("updateTitle")
        main_layout.addWidget(lbl_title)

        self.lbl_repo = QtWidgets.QLabel(f"저장소 경로: {str(self.repo_root)}")
        self.lbl_repo.setObjectName("updateRepo")
        main_layout.addWidget(self.lbl_repo)

        self.lbl_hash = QtWidgets.QLabel("현재 githash: (알 수 없음)")
        self.lbl_hash.setObjectName("updateHash")
        main_layout.addWidget(self.lbl_hash)

        main_layout.addSpacing(4)

        # 버튼 영역
        btn_layout = QtWidgets.QHBoxLayout()
        btn_layout.setSpacing(10)

        self.btn_refresh = QtWidgets.QPushButton("상태 새로고침 (해시/상태)")
        self.btn_refresh.setObjectName("btnUpdateRefresh")

        self.btn_pull = QtWidgets.QPushButton("업데이트 실행 (git pull)")
        self.btn_pull.setObjectName("btnUpdatePull")
        self.btn_pull.setProperty("accent", True)

        btn_layout.addWidget(self.btn_refresh, 0)
        btn_layout.addWidget(self.btn_pull, 0)
        btn_layout.addStretch(1)
        main_layout.addLayout(btn_layout)

        # 로그 영역
        self.log = QtWidgets.QPlainTextEdit()
        self.log.setObjectName("updateLog")
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(260)
        main_layout.addWidget(self.log, 1)

        # 시그널 연결
        self.btn_refresh.clicked.connect(self.on_refresh_clicked)
        self.btn_pull.clicked.connect(self.on_pull_clicked)

        # 초기 로드
        self.on_refresh_clicked()

    def _append_log(self, text: str) -> None:
        self.log.appendPlainText(text)

    def _find_git_root(self) -> Path | None:
        cur = self.repo_root
        for _ in range(5):
            if (cur / ".git").is_dir():
                return cur
            if cur.parent == cur:
                break
            cur = cur.parent
        return None

    def _run_git(self, args: list[str]) -> subprocess.CompletedProcess | None:
        git_root = self._find_git_root()
        if git_root is None:
            self._append_log("[오류] 현재 경로 기준으로 .git 폴더를 찾지 못했습니다.")
            return None

        cmd = ["git"] + args
        self._append_log(f"$ {' '.join(cmd)} (cwd={git_root})")

        try:
            proc = subprocess.run(
                cmd,
                cwd=str(git_root),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
        except FileNotFoundError:
            self._append_log("[오류] 'git' 명령을 찾을 수 없습니다. PC에 Git이 설치되어 있는지 확인하세요.")
            return None
        except Exception as e:
            self._append_log(f"[예외] git 실행 중 오류: {e}")
            return None

        if proc.stdout:
            self._append_log(proc.stdout.strip() or "(출력 없음)")
        return proc

    def on_refresh_clicked(self):
        self._append_log("=" * 60)
        self._append_log("[정보] Git 상태 새로고침 시작")
        proc_hash = self._run_git(["rev-parse", "--short", "HEAD"])
        if proc_hash and proc_hash.returncode == 0 and proc_hash.stdout:
            hash_str = proc_hash.stdout.strip().splitlines()[0]
            self.lbl_hash.setText(f"현재 githash: {hash_str}")
        else:
            self.lbl_hash.setText("현재 githash: (읽기 실패)")
        self._run_git(["status", "-sb"])
        self._append_log("[정보] 상태 새로고침 완료\n")

    def on_pull_clicked(self):
        import git
        import os
        try:
            my_repo = git.Repo()
            my_repo.remotes.origin.pull()
            self._append_log("Git Pull 성공! 재시작합니다...")
            os.execl(sys.executable, sys.executable, *sys.argv)
        except Exception as e:
            self._append_log(f"[업데이트 오류] {e}")


# =========================================================================
# 메인 윈도우
# =========================================================================
class MainTabbedWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        # 버전 정보 읽기
        version = ""
        try:
            base_dir = Path(__file__).resolve().parent
            ver_file = base_dir / "ver.txt"
            if ver_file.is_file():
                version = ver_file.read_text(encoding="utf-8").strip()
        except Exception:
            version = ""

        title = "하비 브라운 엑셀 도구 모음"
        if version:
            title = f"{title}  ({version})"

        self.setWindowTitle(title)
        self.resize(1200, 800)

        # 중앙 탭 위젯
        self.tabs = QtWidgets.QTabWidget(self)
        self.setCentralWidget(self.tabs)

        force_tab_size_only(self.tabs, tab_w=230, tab_h=45)

        tabbar = self.tabs.tabBar()
        tabbar.setElideMode(QtCore.Qt.ElideNone)
        tabbar.setUsesScrollButtons(True)
        tabbar.setExpanding(False)

        # 폰트 강제 조정
        app = QtWidgets.QApplication.instance()
        if app is not None:
            f = app.font()
            if f.pointSize() <= 0 or f.pointSize() > 11:
                f.setPointSize(10)
                app.setFont(f)

        # --------------------------
        # 탭 추가
        # --------------------------
        # 1. 엑셀 계산기
        self.cal_window = ExcelCalWindow(self)
        self.cal_window.setObjectName("tab_pink")
        self.tabs.addTab(self.cal_window, "부가세 계산 / 3종 엑셀")

        # 2. 송장 읽기
        self.read_invoice_widget = ReadInvoiceWidget(self)
        self.read_invoice_widget.setObjectName("tab_dark")
        self.tabs.addTab(self.read_invoice_widget, "네이버·쿠팡 송장 엑셀 읽기")

        # 3. 스케줄
        self.schedule_widget = ScheduleWidget(self)
        self.schedule_widget.setObjectName("tab_blue")
        self.tabs.addTab(self.schedule_widget, "스케쥴")

        # 4. [연동] 근태 관리 (Timeclock) 탭
        self.timeclock_widget = None
        self.timeclock_db = None
        self.timeclock_tab_index = -1

        if HAS_TIMECLOCK:
            self.timeclock_tab_index = self.tabs.addTab(QtWidgets.QWidget(), "⏰ 근태 관리(사업주)")
        else:
            lbl = QtWidgets.QLabel("Timeclock 모듈을 찾을 수 없습니다.\nC:\\my_games\\timeclock 경로 및 라이브러리를 확인해주세요.")
            lbl.setAlignment(QtCore.Qt.AlignCenter)
            self.tabs.addTab(lbl, "⏰ 근태 관리 (오류)")

        # 5. Git 업데이트
        base_dir = Path(__file__).resolve().parent
        self.update_widget = UpdateWidget(base_dir, self)
        self.update_widget.setObjectName("tab_orange")
        self.tabs.addTab(self.update_widget, "업데이트 (git pull)")

        # --------------------------
        # 비동기 로더 & 배지 타이머 설정
        # --------------------------
        if HAS_TIMECLOCK:
            # (1) 동기화 워커
            self.sync_worker = TimeclockSyncWorker(self)
            self.sync_worker.finished_sig.connect(self.on_sync_finished)

            # (2) 배지 체크 워커 & 타이머
            self.badge_worker = BadgeCheckWorker(self)
            self.badge_worker.badge_updated.connect(self.update_badge_ui)

            # 5분(300초)마다 배지 확인
            self.badge_timer = QtCore.QTimer(self)
            self.badge_timer.interval = 300 * 1000
            self.badge_timer.timeout.connect(self.run_badge_check)
            self.badge_timer.start()

        # 로딩 오버레이 (탭 위에 덮어씌움)
        self.loading_overlay = LoadingOverlay(self)

        # 탭 변경 이벤트
        self.tabs.currentChanged.connect(self.on_tab_changed)

        # 테마 설정
        self.theme_by_index = {}
        idx = 0
        self.theme_by_index[idx] = "pink";
        idx += 1
        self.theme_by_index[idx] = "dark";
        idx += 1
        self.theme_by_index[idx] = "blue";
        idx += 1
        if HAS_TIMECLOCK:
            self.theme_by_index[idx] = "green";
            idx += 1
        self.theme_by_index[idx] = "orange";
        idx += 1

        # 초기 테마 적용
        self.apply_theme_for_index(self.tabs.currentIndex())

    # ----------------------------------------------------
    # 로직: 탭 변경 시
    # ----------------------------------------------------
    def apply_theme_for_index(self, idx: int) -> None:
        theme = self.theme_by_index.get(idx, "pink")
        self.tabs.tabBar().setProperty("theme", theme)
        self.tabs.tabBar().style().unpolish(self.tabs.tabBar())
        self.tabs.tabBar().style().polish(self.tabs.tabBar())
        self.tabs.update()

    def on_tab_changed(self, index):
        self.apply_theme_for_index(index)

        # 타임클락 탭을 눌렀을 때
        if HAS_TIMECLOCK and index == self.timeclock_tab_index:
            self.start_async_sync()

    def start_async_sync(self):
        """비동기 동기화 시작 (로딩화면 표시)"""
        if self.sync_worker.isRunning():
            return

        # 1. 로딩 화면 띄우기
        current_widget = self.tabs.widget(self.timeclock_tab_index)
        self.loading_overlay.setParent(current_widget)
        self.loading_overlay.show_loading()

        # 2. 스레드 시작
        self.sync_worker.start()

    def on_sync_finished(self, ok, msg):
        """동기화 스레드 종료 시 호출"""
        # 로딩 끄기
        self.loading_overlay.hide_loading()

        if not ok:
            print(f"[Timeclock] Sync Failed: {msg}")
            # 실패해도 기존 데이터로 로드는 시도

        # UI 로드 (메인 스레드에서 실행)
        self.load_or_refresh_timeclock_ui()

    def load_or_refresh_timeclock_ui(self):
        """DB연결 및 OwnerPage 생성 (기존 연결 끊고 재연결)"""
        try:
            # 1. 기존 DB 연결이 있으면 닫기 (★중요: 파일 핸들 놓기)
            if self.timeclock_db:
                try:
                    self.timeclock_db.close()
                except:
                    pass

            # 2. 새 파일에 대해 DB 다시 연결
            self.timeclock_db = DB(DB_PATH)

            # 3. 세션 생성 (Owner)
            user = self.timeclock_db.get_user_by_username("owner")
            if not user:
                self.timeclock_db.create_user(DEFAULT_OWNER_USER, "owner", DEFAULT_OWNER_PASS)
                user = self.timeclock_db.get_user_by_username("owner")

            session = Session(
                user_id=user['id'], username=user['username'], role=user['role'],
                must_change_pw=(user['must_change_pw'] == 1), job_title=user['job_title']
            )

            # 4. 위젯 생성 또는 갱신
            if self.timeclock_widget is None:
                self.timeclock_widget = OwnerPage(self.timeclock_db, session)

                tab_w = self.tabs.widget(self.timeclock_tab_index)
                if tab_w.layout():
                    QtWidgets.QWidget().setLayout(tab_w.layout())  # 기존 레이아웃 제거

                lay = QtWidgets.QVBoxLayout(tab_w)
                lay.setContentsMargins(0, 0, 0, 0)
                lay.addWidget(self.timeclock_widget)
            else:
                # [핵심] 기존 위젯에 새 DB 연결 주입
                self.timeclock_widget.db = self.timeclock_db
                self.timeclock_widget.session = session

                # 화면 새로고침 (새 DB에서 데이터를 읽어옴)
                self.timeclock_widget.refresh_work_logs()
                self.timeclock_widget.refresh_disputes()
                self.timeclock_widget.refresh_signup_requests()
                self.timeclock_widget.update_badges()

            # 로드 완료 후 배지 업데이트 한 번 실행
            self.run_badge_check()

        except Exception as e:
            print(f"[UI Load Error] {e}")
            import traceback
            traceback.print_exc()

    # ----------------------------------------------------
    # 로직: 배지 (알림)
    # ----------------------------------------------------
    def run_badge_check(self):
        # 현재 타임클락 탭을 보고 있다면 UI에서 직접 카운트 가져와서 탭 이름 갱신
        if self.tabs.currentIndex() == self.timeclock_tab_index:
            if self.timeclock_widget:
                try:
                    cnts = self.timeclock_db.get_pending_counts()
                    total = sum(cnts.values())
                    self.update_badge_ui(total)
                except:
                    pass
            return

        # 다른 탭을 보고 있다면 백그라운드 체크 시작
        if not self.badge_worker.isRunning():
            self.badge_worker.start()

    def update_badge_ui(self, count):
        """탭 타이틀 변경"""
        if count > 0:
            self.tabs.setTabText(self.timeclock_tab_index, f"⏰ 근태 관리 (🔴 {count})")
        else:
            self.tabs.setTabText(self.timeclock_tab_index, "⏰ 근태 관리")


# -------------------------------------------------------------------------
# 유틸리티 (스타일, 덤프 등)
# -------------------------------------------------------------------------
def _resource_path(rel: str) -> Path:
    """pyinstaller(onefile) 리소스 경로 계산"""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / rel


def _apply_app_style(app) -> None:
    """QSS 스타일 적용"""
    qss_path = _resource_path("app_style.qss")
    base_qss = r"""
    /* --------- 전체 기본 --------- */
    QWidget {
        font-size: 10pt;
        font-family: 'Malgun Gothic';
    }

    /* --------- 탭 글자 잘림 방지 --------- */
    QTabWidget::pane {
        border: 1px solid #e6e6e6;
        top: -1px;
        background: #fbfbfd;
    }
    QTabBar::tab {
        min-height: 34px;            
        padding: 6px 14px;           
        margin-right: 6px;
        border: 1px solid #e6e6e6;
        border-bottom: none;
        border-top-left-radius: 10px;
        border-top-right-radius: 10px;
        background: #f2f3f7;         
        color: #222;
    }
    QTabBar::tab:selected {
        background: #fff1f3;         
        border-color: #f0b6c1;
        color: #111;
        font-weight: 600;
    }
    QTabBar::tab:hover {
        background: #ffe4e8;
    }

    /* 탭별 테마 (속성 기반) */
    QTabBar::tab[theme="pink"]:selected { background: #fff1f3; border-color: #f0b6c1; color: #d81b60; border-top: 3px solid #d81b60; }
    QTabBar::tab[theme="dark"]:selected { background: #e0e0e0; border-color: #999; color: #333; border-top: 3px solid #333; }
    QTabBar::tab[theme="blue"]:selected { background: #e3f2fd; border-color: #90caf9; color: #1976d2; border-top: 3px solid #1976d2; }
    QTabBar::tab[theme="orange"]:selected { background: #fff3e0; border-color: #ffcc80; color: #f57c00; border-top: 3px solid #f57c00; }
    QTabBar::tab[theme="green"]:selected { background: #e8f5e9; border-color: #a5d6a7; color: #2e7d32; border-top: 3px solid #2e7d32; }

    /* --------- 입력창 스타일 --------- */
    QLineEdit, QDateEdit, QComboBox, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox {
        min-height: 28px;            
        padding: 4px 10px;           
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        background: #ffffff;
        color: #111111;
    }
    QLineEdit:focus, QDateEdit:focus, QComboBox:focus, QTextEdit:focus, QPlainTextEdit:focus {
        border: 1px solid #f0b6c1;
        background: #fffafb;
    }

    /* --------- 그룹박스/라벨 기본 --------- */
    QGroupBox {
        border: 1px solid #ececec;
        border-radius: 10px;
        margin-top: 12px;
        background: #fbfbfd;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 0 6px;
        left: 10px;
        color: #333;
        font-weight: 600;
    }
    QLabel {
        color: #222;
    }

    /* --------- 버튼 --------- */
    QPushButton {
        min-height: 30px;
        padding: 6px 14px;
        border-radius: 10px;
        border: 1px solid #f0b6c1;
        background: #fff1f3;
        color: #111;
        font-weight: 600;
    }
    QPushButton:hover {
        background: #ffe4e8;
    }
    QPushButton:pressed {
        background: #ffd3da;
    }
    """

    if qss_path.is_file():
        try:
            file_qss = qss_path.read_text(encoding="utf-8")
        except Exception:
            file_qss = qss_path.read_text(encoding="utf-8", errors="replace")
        app.setStyleSheet(file_qss)
    else:
        app.setStyleSheet(base_qss)


def install_global_exception_dump(log_dir: str) -> None:
    """프로그램이 시작하자마자 꺼지는 경우를 파일로 남긴다"""
    import os
    import sys
    import traceback
    from datetime import datetime
    from pathlib import Path

    Path(log_dir).mkdir(parents=True, exist_ok=True)
    dump_path = Path(log_dir) / "exception_dump.log"

    def _write_dump(prefix: str, exc_type, exc, tb):
        try:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            msg = "".join(traceback.format_exception(exc_type, exc, tb))
            with dump_path.open("a", encoding="utf-8") as f:
                f.write("\n" + "=" * 80 + "\n")
                f.write(f"[{ts}] {prefix}\n")
                f.write(msg + "\n")
        except Exception:
            pass

    sys.excepthook = lambda t, e, tb: _write_dump("sys.excepthook", t, e, tb)


def force_tab_size_only(tabs: QtWidgets.QTabWidget, *, tab_w: int = 700, tab_h: int = 50) -> None:
    """탭 크기만 강제로 적용"""
    old = tabs.tabBar()
    new_bar = FixedSizeTabBar(tab_w=tab_w, tab_h=tab_h, parent=tabs)

    try:
        new_bar.setMovable(old.isMovable())
        new_bar.setTabsClosable(old.tabsClosable())
        new_bar.setDrawBase(old.drawBase())
    except Exception:
        pass

    tabs.setTabBar(new_bar)
    new_bar.setMinimumHeight(tab_h)


class FixedSizeTabBar(QtWidgets.QTabBar):
    """탭 크기 고정 TabBar"""

    def __init__(self, tab_w: int = 350, tab_h: int = 40, parent=None):
        super().__init__(parent)
        self._tab_w = int(tab_w)
        self._tab_h = int(tab_h)
        self.setUsesScrollButtons(True)
        self.setExpanding(False)

    def set_tab_size(self, tab_w: int, tab_h: int) -> None:
        self._tab_w = int(tab_w)
        self._tab_h = int(tab_h)
        self.updateGeometry()
        self.update()

    def tabSizeHint(self, index: int) -> QtCore.QSize:
        base = super().tabSizeHint(index)
        w = max(base.width(), self._tab_w)
        h = max(base.height(), self._tab_h)
        return QtCore.QSize(w, h)


def main():
    import sys
    from pathlib import Path
    from PyQt5 import QtWidgets

    # ✅ 로그 폴더 경로 (네가 쓰는 경로로 고정)
    log_dir = r"C:\my_games\excel_cal\log"
    install_global_exception_dump(log_dir)

    app = QtWidgets.QApplication(sys.argv)

    # 스타일시트 적용
    _apply_app_style(app)

    try:
        win = MainTabbedWindow()
        win.show()
        sys.exit(app.exec_())
    except Exception as e:
        raise


if __name__ == "__main__":
    main()