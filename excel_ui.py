# excel_ui.py
# 탭 통합 메인 UI: 기존 엑셀 계산기 + 네이버/쿠팡 송장 읽기 + git 업데이트 탭

import sys
import subprocess
from pathlib import Path

from PyQt5 import QtWidgets

from excel_cal_ui import ExcelCalWindow      # 기존 부가세/3종 엑셀 UI
from read_excel import ReadInvoiceWidget     # 송장 읽기 탭


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

        # NOTE:
        # 여기서 self.on_refresh_clicked()를 바로 호출하면
        # 일부 검사기에서 "도달할 수 없습니다" 경고가 뜰 수 있으니
        # MainTabbedWindow 쪽에서 초기 1회 호출하도록 함.

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
        self._append_log("=" * 60)
        self._append_log("[정보] git pull 실행 시작")

        # 버튼 잠깐 비활성화 (중복 클릭 방지)
        self.btn_refresh.setEnabled(False)
        self.btn_pull.setEnabled(False)

        try:
            proc = self._run_git(["pull", "--ff-only"])
            if proc and proc.returncode == 0:
                self._append_log("[성공] git pull이 정상적으로 완료되었습니다.")
                # 알림창
                QtWidgets.QMessageBox.information(
                    self,
                    "업데이트 완료",
                    "코드 업데이트가 완료되었습니다.\n프로그램을 다시 시작합니다."
                )
                # 재시작 로직
                self._restart_app()
            else:
                self._append_log("[주의] git pull 도중 문제가 발생했습니다. 위 로그를 확인하세요.")
        finally:
            # pull 성공 시에는 _restart_app()에서 종료되므로
            # 여기 버튼 재활성화는 실패/예외 케이스용
            self.btn_refresh.setEnabled(True)
            self.btn_pull.setEnabled(True)

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

        # 1. 기존 부가세/3종 엑셀 생성기 탭
        self.cal_window = ExcelCalWindow()
        tabs.addTab(self.cal_window, "부가세 계산 / 3종 엑셀")

        # 2. 네이버/쿠팡 송장 엑셀 읽기 탭
        self.read_invoice_widget = ReadInvoiceWidget(self)
        tabs.addTab(self.read_invoice_widget, "네이버·쿠팡 송장 엑셀 읽기")

        # 3. Git 업데이트 탭
        base_dir = Path(__file__).resolve().parent  # 보통 C:\my_games\excel_cal
        self.update_widget = UpdateWidget(base_dir, self)
        tabs.addTab(self.update_widget, "업데이트 (git pull)")

        # 여기서 한 번만 상태 새로고침 호출 → "도달할 수 없습니다" 경고 안 뜸
        self.update_widget.on_refresh_clicked()


def main():
    app = QtWidgets.QApplication(sys.argv)
    win = MainTabbedWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
