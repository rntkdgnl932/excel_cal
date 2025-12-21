# excel_ui.py
# 탭 통합 메인 UI: 기존 엑셀 계산기 + 네이버/쿠팡 송장 읽기 + git 업데이트 탭

import sys
import subprocess
from pathlib import Path

from PyQt5 import QtWidgets, QtCore

from excel_cal_ui import ExcelCalWindow      # 기존 부가세/3종 엑셀 UI
from read_excel import ReadInvoiceWidget     # 송장 읽기 탭
from schedule import ScheduleWidget


class UpdateWidget(QtWidgets.QWidget):
    """
    Git 저장소 기준으로 현재 githash 확인 + git pull 실행하는 탭.
    - repo_root 아래/상위에 .git 폴더가 있어야 정상 동작
    """
    def __init__(self, repo_root: Path, parent=None):
        super().__init__(parent)
        self.repo_root = repo_root

        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        # 상단 설명 및 경로 / 해시 표시
        lbl_title = QtWidgets.QLabel("엑셀 도구 Git 업데이트")
        lbl_title.setStyleSheet("font-weight: bold; font-size: 14px;")
        main_layout.addWidget(lbl_title)

        self.lbl_repo = QtWidgets.QLabel(f"저장소 경로: {str(self.repo_root)}")
        main_layout.addWidget(self.lbl_repo)

        self.lbl_hash = QtWidgets.QLabel("현재 githash: (알 수 없음)")
        main_layout.addWidget(self.lbl_hash)

        # 버튼 영역
        btn_layout = QtWidgets.QHBoxLayout()
        self.btn_refresh = QtWidgets.QPushButton("상태 새로고침 (해시/상태)")
        self.btn_pull = QtWidgets.QPushButton("업데이트 실행 (git pull)")
        btn_layout.addWidget(self.btn_refresh)
        btn_layout.addWidget(self.btn_pull)
        btn_layout.addStretch(1)
        main_layout.addLayout(btn_layout)

        # 로그 영역
        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(250)
        main_layout.addWidget(self.log, 1)

        # 시그널 연결
        self.btn_refresh.clicked.connect(self.on_refresh_clicked)
        self.btn_pull.clicked.connect(self.on_pull_clicked)


    # ---------------------------
    # 내부 유틸: 로그/경로/실행
    # ---------------------------
    def _append_log(self, text: str) -> None:
        self.log.appendPlainText(text)

    def _find_git_root(self) -> Path | None:
        """
        repo_root 기준으로 상위로 올라가며 .git 폴더를 찾는다.
        (exe로 빌드된 경우에도 경로만 맞으면 동작)
        """
        cur = self.repo_root
        for _ in range(5):  # 너무 멀리는 안 감
            if (cur / ".git").is_dir():
                return cur
            if cur.parent == cur:
                break
            cur = cur.parent
        return None

    def _run_git(self, args: list[str]) -> subprocess.CompletedProcess | None:
        """
        git 명령을 실행하고 결과를 반환.
        - Git 미설치 / .git 없음 등은 로그에 메시지 출력.
        """
        git_root = self._find_git_root()
        if git_root is None:
            self._append_log("[오류] 현재 경로 기준으로 .git 폴더를 찾지 못했습니다.")
            self._append_log("       엑셀 도구 폴더가 Git 저장소인지 확인해 주세요.")
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

        if proc.returncode != 0:
            self._append_log(f"[오류] git 명령이 실패했습니다. (returncode={proc.returncode})")

        return proc

    # ---------------------------
    # 버튼 핸들러
    # ---------------------------
    def on_refresh_clicked(self):
        """
        현재 HEAD 해시 + git status를 화면에 표시.
        """
        self._append_log("=" * 60)
        self._append_log("[정보] Git 상태 새로고침 시작")

        # 해시
        proc_hash = self._run_git(["rev-parse", "--short", "HEAD"])
        if proc_hash and proc_hash.returncode == 0 and proc_hash.stdout:
            hash_str = proc_hash.stdout.strip().splitlines()[0]
            self.lbl_hash.setText(f"현재 githash: {hash_str}")
        else:
            self.lbl_hash.setText("현재 githash: (읽기 실패)")

        # status
        self._run_git(["status", "-sb"])

        self._append_log("[정보] 상태 새로고침 완료")
        self._append_log("")

    def on_pull_clicked(self):
        """
        git pull 실행 (origin 기준, 기본 브랜치).
        - pull 성공 시: 프로그램 자동 재시작
        """
        import git
        import os
        my_repo = git.Repo()
        my_repo.remotes.origin.pull()
        # 실행 후 재시작 부분
        os.execl(sys.executable, sys.executable, *sys.argv)

    # ---------------------------
    # 재시작 로직
    # ---------------------------
    def _restart_app(self):
        """
        현재 프로세스를 종료하고, 동일한 명령줄로 새 프로세스를 띄움.
        - pyinstaller exe든, python 스크립트든 sys.executable + sys.argv 사용.
        """
        try:
            python = sys.executable
            args = sys.argv[:]  # 현재 인자 그대로
            self._append_log(f"[정보] 재시작: {python} {' '.join(args)}")
            subprocess.Popen([python] + args)
        except Exception as e:
            self._append_log(f"[예외] 재시작 실패: {e}")
            # 재시작까지 실패하면 그냥 여기서 끝.
            return

        # 새 프로세스를 띄웠으니, 현재 앱 종료
        app = QtWidgets.QApplication.instance()
        if app is not None:
            app.quit()
        else:
            # 혹시나 앱 인스턴스가 없으면 프로세스를 직접 종료
            sys.exit(0)


class MainTabbedWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("하비 브라운 엑셀 도구 모음")
        self.resize(1200, 800)

        # 중앙에 탭 위젯 배치
        tabs = QtWidgets.QTabWidget(self)
        self.setCentralWidget(tabs)

        force_tab_size_only(tabs, tab_w=250, tab_h=45)  # ✅ 탭 크기만 강제

        # ✅ 탭 글자 잘림 방지 (탭 바/폰트/패딩)
        tabbar = tabs.tabBar()
        tabbar.setElideMode(QtCore.Qt.ElideNone)      # "..." 생략 금지
        tabbar.setUsesScrollButtons(True)            # 탭 많으면 스크롤 버튼
        tabbar.setExpanding(False)                   # 탭이 균등 확장되며 잘리는 현상 방지

        # 탭/입력창이 너무 큰 기본 폰트가 잡혀있으면 강제로 정리
        # (QSS가 적용되더라도 기본 폰트가 크면 내부 위젯들이 같이 커지는 경우가 있음)
        app = QtWidgets.QApplication.instance()
        if app is not None:
            f = app.font()
            # 너무 커져 있으면 내려줌 (원하면 숫자만 조절)
            if f.pointSize() <= 0 or f.pointSize() > 11:
                f.setPointSize(10)
                app.setFont(f)

        # 1. 기존 부가세/3종 엑셀 생성기 탭
        self.cal_window = ExcelCalWindow()
        tabs.addTab(self.cal_window, "부가세 계산 / 3종 엑셀")

        # 2. 네이버/쿠팡 송장 엑셀 읽기 탭
        self.read_invoice_widget = ReadInvoiceWidget(self)
        tabs.addTab(self.read_invoice_widget, "네이버·쿠팡 송장 엑셀 읽기")

        # 3. 스케쥴 탭
        self.schedule_widget = ScheduleWidget(self)
        tabs.addTab(self.schedule_widget, "스케쥴")

        # 4. Git 업데이트 탭
        base_dir = Path(__file__).resolve().parent  # 보통 C:\my_games\excel_cal
        self.update_widget = UpdateWidget(base_dir, self)
        tabs.addTab(self.update_widget, "업데이트 (git pull)")


        # 여기서 한 번만 상태 새로고침 호출 → "도달할 수 없습니다" 경고 안 뜸
        self.update_widget.on_refresh_clicked()


def _resource_path(rel: str) -> Path:
    """
    pyinstaller(onefile)에서도 동작하도록 리소스 경로를 계산
    - 개발 실행: excel_ui.py가 있는 폴더 기준
    - exe 실행: sys._MEIPASS 기준
    """
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / rel


def _apply_app_style(app) -> None:
    """
    1) app_style.qss가 있으면 우선 적용
    2) 없어도 기본 스타일(탭/입력창 최소 높이, 과도한 폰트 방지)을 강제로 적용
    """
    qss_path = _resource_path("app_style.qss")
    base_qss = r"""
    /* --------- 전체 기본 --------- */
    QWidget {
        font-size: 10pt;
    }

    /* --------- 탭 글자 잘림 방지 --------- */
    QTabWidget::pane {
        border: 1px solid #e6e6e6;
        top: -1px;
        background: #fbfbfd;
    }
    QTabBar::tab {
        min-height: 34px;            /* ✅ 탭 높이 확보 */
        padding: 6px 14px;           /* ✅ 글자 좌우/상하 여백 */
        margin-right: 6px;
        border: 1px solid #e6e6e6;
        border-bottom: none;
        border-top-left-radius: 10px;
        border-top-right-radius: 10px;
        background: #f2f3f7;         /* 흰색 눈부심 줄임 */
        color: #222;
    }
    QTabBar::tab:selected {
        background: #fff1f3;         /* 은은한 핑크/레드 계열 */
        border-color: #f0b6c1;
        color: #111;
        font-weight: 600;
    }
    QTabBar::tab:hover {
        background: #ffe4e8;
    }

    /* --------- 입력창(placeholder/값) 안 보임 방지 --------- */
    QLineEdit, QDateEdit, QComboBox, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox {
        min-height: 28px;            /* ✅ 글자 높이 확보 */
        padding: 4px 10px;           /* ✅ 내부 여백 */
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

    /* --------- 버튼(은은한 레드 포인트) --------- */
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

    # 1) 파일 QSS가 있으면 읽어서 적용 + base_qss를 뒤에 덧붙여 "안전장치"로 강제
    if qss_path.is_file():
        try:
            file_qss = qss_path.read_text(encoding="utf-8")
        except Exception:
            file_qss = qss_path.read_text(encoding="utf-8", errors="replace")
        app.setStyleSheet(file_qss)

    else:
        # 2) qss 파일이 없으면 base_qss만 적용
        app.setStyleSheet(base_qss)



def install_global_exception_dump(log_dir: str) -> None:
    """
    프로그램이 시작하자마자 꺼지는 경우(콘솔이 닫혀서 안 보이는 예외)를 파일로 남긴다.
    - exception_dump.log에 traceback 기록
    """
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

    def excepthook(exc_type, exc, tb):
        _write_dump("sys.excepthook", exc_type, exc, tb)

    sys.excepthook = excepthook

    # Qt 쪽 예외도 최대한 남기기(일부 환경에서만 동작)
    try:
        from PyQt5 import QtCore

        def qt_message_handler(mode, context, message):
            # Qt 내부 경고/에러도 파일로
            try:
                ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with dump_path.open("a", encoding="utf-8") as f:
                    f.write("\n" + "-" * 80 + "\n")
                    f.write(f"[{ts}] QtMessage: {message}\n")
            except Exception:
                pass

        QtCore.qInstallMessageHandler(qt_message_handler)
    except Exception:
        pass

def force_tab_size_only(tabs: QtWidgets.QTabWidget, *, tab_w: int = 700, tab_h: int = 50) -> None:
    """
    '탭 크기만' 강제로 적용.
    - QSS가 있어도 무조건 커짐
    - 다른 위젯/스타일 안 건드림
    """
    old = tabs.tabBar()
    new_bar = FixedSizeTabBar(tab_w=tab_w, tab_h=tab_h, parent=tabs)

    # 기존 탭바의 일부 설정 승계(선택사항)
    try:
        new_bar.setMovable(old.isMovable())
        new_bar.setTabsClosable(old.tabsClosable())
        new_bar.setDrawBase(old.drawBase())
    except Exception:
        pass

    tabs.setTabBar(new_bar)

    # 탭바 높이도 확실히 확보(세로 잘림 방지)
    new_bar.setMinimumHeight(tab_h)



class FixedSizeTabBar(QtWidgets.QTabBar):
    """
    탭 크기만 강제로 고정하는 TabBar.
    QSS/레이아웃이 뭐라고 하든 tabSizeHint를 고정해서 '진짜로' 탭 크기가 커진다.
    """
    def __init__(self, tab_w: int = 350, tab_h: int = 40, parent=None):
        super().__init__(parent)
        self._tab_w = int(tab_w)
        self._tab_h = int(tab_h)

        # 탭이 많아지면 스크롤 버튼
        self.setUsesScrollButtons(True)
        # 균등분배로 폭 쪼개지는 것 방지
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

    # (선택) 스타일 적용은 여기서 하되, 실패해도 절대 죽지 않게
    try:
        base_dir = Path(__file__).resolve().parent
        qss_path = base_dir / "app_style.qss"
        if qss_path.exists():
            app.setStyleSheet(qss_path.read_text(encoding="utf-8"))
    except Exception:
        pass

    try:
        win = MainTabbedWindow()
        win.show()
        sys.exit(app.exec_())
    except Exception as e:
        # 여기서 죽어도 exception_dump.log에 남게 된다
        raise





if __name__ == "__main__":
    main()



# python -m PyInstaller `
#   --noconfirm `
#   --clean `
#   --name excel_cal `
#   --icon "icon.ico" `
#   --add-data "icon.ico;." `
#   --add-data "app_style.qss;." `
#   --hidden-import PyQt5 `
#   main.py

