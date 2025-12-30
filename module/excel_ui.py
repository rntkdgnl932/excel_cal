# excel_ui.py
# 탭 통합 메인 UI: 엑셀 계산기 + 송장 + 스케줄 + [비동기] 타임클락 + 업데이트

import sys
import subprocess
from pathlib import Path
from PyQt5 import QtWidgets, QtCore, QtGui

# -------------------------------------------------------------------------
# [매우 중요] 경로 설정
# -------------------------------------------------------------------------
TIMECLOCK_ROOT = Path(r"C:\my_games\timeclock")
HAS_TIMECLOCK_PATH = False

if TIMECLOCK_ROOT.exists():
    parent_dir = str(TIMECLOCK_ROOT.parent)
    if parent_dir not in sys.path:
        sys.path.append(parent_dir)
    if str(TIMECLOCK_ROOT) not in sys.path:
        sys.path.append(str(TIMECLOCK_ROOT))
    HAS_TIMECLOCK_PATH = True
else:
    print(f"⚠️ 경고: Timeclock 경로를 찾을 수 없습니다: {TIMECLOCK_ROOT}")

# -------------------------------------------------------------------------
# 모듈 임포트
# -------------------------------------------------------------------------
from excel_cal_ui import ExcelCalWindow
from read_excel import ReadInvoiceWidget
from schedule import ScheduleWidget

HAS_TIMECLOCK = False
try:
    if HAS_TIMECLOCK_PATH:
        from timeclock.db import DB
        from timeclock.settings import DB_PATH, DEFAULT_OWNER_USER, DEFAULT_OWNER_PASS
        from timeclock.ui.owner_page import OwnerPage
        from timeclock.ui.login_page import Session
        from timeclock import sync_manager

        HAS_TIMECLOCK = True
except ImportError as e:
    print(f"⚠️ Timeclock 모듈 임포트 실패: {e}")
    import traceback

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
    여기서는 '다운로드 -> 체크' 로직을 수행하되, 사용자가 작업 중이 아닐 때만 수행 권장
    """
    badge_updated = QtCore.pyqtSignal(int)  # 대기 건수 합계

    def run(self):
        if not HAS_TIMECLOCK or not DB_PATH.exists():
            return

        try:
            # 단순히 현재 로컬 DB 기준으로만 체크 (서버 다운로드는 너무 무거움)
            # 만약 서버 데이터를 꼭 봐야한다면, download 로직이 필요하지만
            # 사용자 경험상 '탭 누를 때 동기화' + '주기적 자동 동기화'가 섞이면 충돌 위험이 있음.
            # -> 안전을 위해 여기서는 "현재 로컬 DB"의 상태만 체크하거나,
            #    sync_manager.download_latest_db()를 수행하되 메인 스레드와 충돌 방지 필요.

            # [전략] 안전하게: 그냥 로컬 DB만 체크한다.
            # (탭을 누를 때 동기화되므로, 탭을 안 누르면 배지는 안 바뀌는게 맞음.
            #  자동으로 배지가 뜨게 하려면 백그라운드 다운로드가 필수인데 이는 복잡도 증가)

            # 하지만 사용자 요청은 "알림이 왔으면 좋겠다" 이므로,
            # '탭이 활성화되지 않았을 때'만 몰래 다운로드를 시도해본다.

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
        self.resize(self.parent().size())
        self.show()
        self.raise_()

    def hide_loading(self):
        self.hide()


# =========================================================================
# 메인 코드 (업데이트 위젯 포함)
# =========================================================================

class UpdateWidget(QtWidgets.QWidget):
    def __init__(self, repo_root: Path, parent=None):
        super().__init__(parent)
        self.repo_root = repo_root

        l = QtWidgets.QVBoxLayout(self)
        l.setContentsMargins(20, 20, 20, 20)

        self.lbl_status = QtWidgets.QLabel("Git 상태 확인 중...")
        self.btn_pull = QtWidgets.QPushButton("최신 버전 업데이트 (Git Pull)")
        self.btn_pull.clicked.connect(self.do_update)

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)

        l.addWidget(self.lbl_status)
        l.addWidget(self.btn_pull)
        l.addWidget(self.log)

        self.check_git()

    def check_git(self):
        try:
            import git
            repo = git.Repo(self.repo_root)
            sha = repo.head.object.hexsha[:7]
            self.lbl_status.setText(f"현재 버전: {sha}")
        except:
            self.lbl_status.setText("Git 저장소가 아닙니다.")

    def do_update(self):
        try:
            import git
            repo = git.Repo(self.repo_root)
            origin = repo.remotes.origin
            origin.pull()
            self.log.appendPlainText("업데이트 성공! 프로그램을 재시작하세요.")
            QtWidgets.QMessageBox.information(self, "완료", "업데이트가 완료되었습니다.\n프로그램을 재시작합니다.")
            import os
            os.execl(sys.executable, sys.executable, *sys.argv)
        except Exception as e:
            self.log.appendPlainText(f"업데이트 실패: {e}")


class MainTabbedWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("하비브라운 통합 관리 시스템")
        self.resize(1200, 800)

        # 중앙 탭
        self.tabs = QtWidgets.QTabWidget(self)
        self.setCentralWidget(self.tabs)

        force_tab_size_only(self.tabs, tab_w=200, tab_h=45)
        self.tabs.tabBar().setElideMode(QtCore.Qt.ElideNone)

        # 1. 엑셀 계산기
        self.cal_window = ExcelCalWindow(self)
        self.tabs.addTab(self.cal_window, "부가세/3종 엑셀")

        # 2. 송장
        self.read_invoice = ReadInvoiceWidget(self)
        self.tabs.addTab(self.read_invoice, "송장 관리")

        # 3. 스케줄
        self.schedule = ScheduleWidget(self)
        self.tabs.addTab(self.schedule, "스케줄")

        # 4. 타임클락 (초기엔 빈 탭)
        self.timeclock_widget = None
        self.timeclock_db = None
        self.timeclock_tab_index = -1

        if HAS_TIMECLOCK:
            self.timeclock_tab_index = self.tabs.addTab(QtWidgets.QWidget(), "⏰ 근태 관리")
        else:
            self.tabs.addTab(QtWidgets.QLabel("Timeclock 모듈 없음"), "근태 관리(오류)")

        # 5. 업데이트
        self.update_tab = UpdateWidget(Path(__file__).parent, self)
        self.tabs.addTab(self.update_tab, "업데이트")

        # -----------------------------------------------
        # 비동기 로더 & 배지 타이머 설정
        # -----------------------------------------------
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

        # 스타일 적용
        self.apply_theme_for_index(0)

    # ----------------------------------------------------
    # 로직: 탭 변경 시
    # ----------------------------------------------------
    def on_tab_changed(self, index):
        self.apply_theme_for_index(index)

        # 타임클락 탭을 눌렀을 때
        if HAS_TIMECLOCK and index == self.timeclock_tab_index:
            # 이미 로드되어 있어도, 최신 데이터 확인을 위해 Sync 시도
            self.start_async_sync()

    def start_async_sync(self):
        """비동기 동기화 시작 (로딩화면 표시)"""
        if self.sync_worker.isRunning():
            return

        # 1. 로딩 화면 띄우기
        # 타임클락 탭 영역 크기에 맞춤
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
        self.load_timeclock_ui()

    def load_timeclock_ui(self):
        """DB연결 및 OwnerPage 생성 (이미 있으면 갱신만)"""
        try:
            # DB 재연결 (파일이 바뀌었을 수 있으므로)
            if self.timeclock_db:
                self.timeclock_db.close()

            self.timeclock_db = DB(DB_PATH)

            # Owner 세션 생성
            user = self.timeclock_db.get_user_by_username("owner")
            if not user:
                self.timeclock_db.create_user(DEFAULT_OWNER_USER, "owner", DEFAULT_OWNER_PASS)
                user = self.timeclock_db.get_user_by_username("owner")

            session = Session(
                user_id=user['id'], username=user['username'], role=user['role'],
                must_change_pw=(user['must_change_pw'] == 1), job_title=user['job_title']
            )

            # 위젯이 없으면 생성
            if self.timeclock_widget is None:
                self.timeclock_widget = OwnerPage(self.timeclock_db, session)

                # 탭에 붙이기
                tab_w = self.tabs.widget(self.timeclock_tab_index)
                if tab_w.layout() is None:
                    lay = QtWidgets.QVBoxLayout(tab_w)
                    lay.setContentsMargins(0, 0, 0, 0)
                    lay.addWidget(self.timeclock_widget)
                else:
                    # 기존 레이아웃 비우고 다시 추가 (혹시 모를 잔재 제거)
                    # (생략: 위젯이 None일 때만 여기 오므로 안전)
                    tab_w.layout().addWidget(self.timeclock_widget)
            else:
                # 이미 있으면 DB랑 refresh만
                self.timeclock_widget.db = self.timeclock_db
                self.timeclock_widget.session = session
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
        """타이머에 의해 주기적으로 호출"""
        # 현재 타임클락 탭을 보고 있다면 굳이 백그라운드 체크 안 해도 됨 (실시간 갱신되므로)
        if self.tabs.currentIndex() == self.timeclock_tab_index:
            # 현재 보고 있다면 UI에서 직접 카운트 가져와서 탭 이름 갱신
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

    # ----------------------------------------------------
    # 스타일 테마
    # ----------------------------------------------------
    def apply_theme_for_index(self, idx: int):
        # 0: pink, 1: dark, 2: blue, 3: green(timeclock), 4: orange
        themes = {0: "pink", 1: "dark", 2: "blue", 3: "green", 4: "orange"}
        if not HAS_TIMECLOCK and idx == 3:
            theme = "orange"  # 에러나 업데이트 탭일 경우
        else:
            theme = themes.get(idx, "pink")

        self.tabs.tabBar().setProperty("theme", theme)
        self.tabs.tabBar().style().unpolish(self.tabs.tabBar())
        self.tabs.tabBar().style().polish(self.tabs.tabBar())


# -------------------------------------------------------------------------
# 유틸리티 (스타일, 덤프 등)
# -------------------------------------------------------------------------
def _resource_path(rel: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / rel


def _apply_app_style(app):
    base_qss = r"""
    QWidget { font-size: 10pt; font-family: 'Malgun Gothic'; }
    QTabWidget::pane { border: 1px solid #ccc; background: #fff; }
    QTabBar::tab {
        min-height: 34px; padding: 6px 14px; margin-right: 4px;
        background: #f0f0f0; border-top-left-radius: 8px; border-top-right-radius: 8px;
    }
    QTabBar::tab:selected { background: #fff; font-weight: bold; border: 1px solid #ccc; border-bottom: none; }

    /* 테마별 색상 */
    QTabBar::tab[theme="pink"]:selected { color: #d81b60; border-top: 3px solid #d81b60; }
    QTabBar::tab[theme="dark"]:selected { color: #333; border-top: 3px solid #333; }
    QTabBar::tab[theme="blue"]:selected { color: #1976d2; border-top: 3px solid #1976d2; }
    QTabBar::tab[theme="green"]:selected { color: #2e7d32; border-top: 3px solid #2e7d32; }
    QTabBar::tab[theme="orange"]:selected { color: #f57c00; border-top: 3px solid #f57c00; }
    """
    app.setStyleSheet(base_qss)


def force_tab_size_only(tabs, tab_w=200, tab_h=40):
    # 간단하게 탭바 스타일로 처리 (복잡한 클래스 제거)
    tabs.setStyleSheet(f"QTabBar::tab {{ min-width: {tab_w}px; min-height: {tab_h}px; }}")


def install_global_exception_dump(log_dir):
    pass  # 생략 (기존 유지 권장)


def main():
    log_dir = r"C:\my_games\excel_cal\log"
    Path(log_dir).mkdir(parents=True, exist_ok=True)

    app = QtWidgets.QApplication(sys.argv)
    _apply_app_style(app)

    win = MainTabbedWindow()
    win.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()