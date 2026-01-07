from PyQt5 import QtWidgets, QtCore


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
        on_progress_ui, finalize_ui, dlg = _mk_progress(owner, title, tail_file=tail_file)
        setattr(owner, "_progress_ctx", (on_progress_ui, finalize_ui, dlg))
        reused = False

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

        if not reused:
            try:
                finalize_ui(ok, payload, err)
            except Exception:
                pass
        else:
            try:
                on_progress_ui({"stage": "done", "msg": "[ui] 작업 완료 (창 유지)"})
            except Exception:
                pass

        if callable(on_done):
            try:
                on_done(ok, payload, err)
            except Exception:
                pass

        try:
            th.quit()
            th.wait(100)
        except Exception:
            pass

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

    try:
        jobs = getattr(owner, "_progress_jobs", None)
        if not isinstance(jobs, list):
            jobs = []
        jobs.append(th)
        setattr(owner, "_progress_jobs", jobs)
        setattr(th, "_worker_ref", obj)
    except Exception:
        pass

    try:
        th.start()
    except Exception as start_exc:
        try:
            on_progress_ui({"stage": "error", "msg": f"[error] thread start failed: {start_exc}"})
        except Exception:
            pass
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
    dlg = QtWidgets.QDialog(owner)
    dlg.setWindowTitle(title)
    dlg.setModal(True)
    dlg.resize(500, 320)
    dlg.setWindowFlags(dlg.windowFlags() & ~QtCore.Qt.WindowContextHelpButtonHint)

    dlg.setStyleSheet("""
        QDialog { background-color: #ffffff; }
        QLabel#TitleLabel { font-size: 15px; font-weight: bold; color: #333333; }
        QProgressBar { border: none; background-color: #f1f3f5; border-radius: 4px; height: 6px; }
        QProgressBar::chunk { background-color: #74c0fc; border-radius: 4px; }
        QPlainTextEdit { background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px; padding: 10px; font-family: '맑은 고딕', sans-serif; font-size: 12px; color: #555555; }
        QPushButton { background-color: #ff6b6b; color: white; border-radius: 6px; padding: 6px 14px; font-weight: bold; }
        QPushButton:hover { background-color: #fa5252; }
    """)

    layout = QtWidgets.QVBoxLayout(dlg)
    layout.setContentsMargins(20, 20, 20, 20)
    layout.setSpacing(12)

    lbl = QtWidgets.QLabel(title)
    lbl.setObjectName("TitleLabel")
    layout.addWidget(lbl)

    pbar = QtWidgets.QProgressBar()
    pbar.setRange(0, 0)
    pbar.setTextVisible(False)
    layout.addWidget(pbar)

    log = QtWidgets.QPlainTextEdit()
    log.setReadOnly(True)
    layout.addWidget(log, 1)

    btn_area = QtWidgets.QHBoxLayout()
    btn_close = QtWidgets.QPushButton("닫기")
    btn_close.setVisible(False)
    btn_area.addStretch(1)
    btn_area.addWidget(btn_close)
    layout.addLayout(btn_area)

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
        pbar.setRange(0, 100)
        if ok:
            pbar.setValue(100)
            pbar.setStyleSheet("QProgressBar::chunk { background-color: #a9e34b; }")
            _append("\n[성공] 모든 작업이 완료되었습니다.")
            _append("잠시 후 창이 닫힙니다...")
            QtCore.QTimer.singleShot(1000, dlg.accept)
        else:
            pbar.setValue(150)
            pbar.setStyleSheet("QProgressBar::chunk { background-color: #ff6b6b; }")
            _append(f"\n[오류] 작업 중 문제가 발생했습니다.\n{err}")
            btn_close.setVisible(True)

    btn_close.clicked.connect(dlg.reject)
    dlg.show()
    return on_progress_ui, finalize_ui, dlg

