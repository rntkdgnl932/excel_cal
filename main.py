# C:\my_games\excel_cal\main.py
# -*- coding: utf-8 -*-
from __future__ import annotations
import sys, os
sys.path.append('C:/my_games/excel_cal/module')

venv_path = os.path.join(os.getcwd(), ".venv", "Lib", "site-packages")
if os.path.exists(venv_path) and venv_path not in sys.path:
    sys.path.insert(0, venv_path)

import traceback
from pathlib import Path
import importlib.util

from PyQt5 import QtWidgets

###################################모두 붙여 버리기###############################
import os
import pandas as pd  # pandas 추가
from datetime import datetime # 날짜용 추가
import msoffcrypto

import io
from PyQt5.QtWidgets import QFileDialog # 파일 탐색기
from PyQt5.QtTest import QTest # 대기(wait)용

import re
from pathlib import Path
from typing import List

from PyQt5 import QtWidgets, QtCore

import subprocess
from pathlib import Path

from PyQt5 import QtWidgets, QtCore
import time
from pathlib import Path

import requests
import shutil
from pathlib import Path
from typing import Optional, List, Dict
import subprocess
import pandas as pd
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import Qt, QDateTime, QTimer
from PyQt5.QtWidgets import *


import json
import uuid
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from PyQt5.QtCore import QLocale
from PyQt5.QtWidgets import QCalendarWidget
from typing import Dict, List, Optional, Callable


import faulthandler
from pathlib import Path
from PyQt5 import QtWidgets, QtCore, QtGui
from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.cell import MergedCell as _MC
from openpyxl.styles import Alignment, Border, Side
from openpyxl.worksheet.worksheet import Worksheet

# --- PyInstaller 강제 인식을 위한 더미 임포트 구역 ---
import google_auth_oauthlib
import google_auth_oauthlib.flow
import googleapiclient.discovery
import google.oauth2.credentials
import pydrive.auth
import pydrive.drive
import pandas
import openpyxl
import wsgiref
import wsgiref.simple_server
import google_auth_oauthlib
import google_auth_oauthlib.flow
import googleapiclient.discovery
import pydrive.auth
import pydrive.drive
import pandas
import openpyxl
# --- 데이터베이스 및 유틸리티 ---
import sqlite3      # DB 조작 필수
import threading    # 백그라운드 비동기 작업 필수
import logging      # 에러 및 상태 로깅 필수
import csv          # 근무 기록 내보내기용

# --- 외부 연동 및 동기화 ---
import git          # 시스템 업데이트(GitHub 연동) 필수
from pydrive.auth import GoogleAuth      # 구글 인증
from pydrive.drive import GoogleDrive    # 구글 드라이브 조작

# --- PyQt5 세부 신호 (비동기 처리 시 필수) ---
from PyQt5.QtCore import pyqtSignal, pyqtSlot # 커스텀 신호 전달용
# --
##################################################################################



MODULE_DIRNAME = "module"
ENTRY_FILENAME = "excel_ui.py"      # module 폴더 안 엔트리 파일
ENTRY_MODNAME = "excel_ui"          # 로딩할 모듈 이름(고정)
QSS_FILENAME = "app_style.qss"      # module 폴더 안에 둔다


def _runtime_root() -> Path:
    """PyInstaller exe 실행이면 exe 폴더, 개발 실행이면 이 파일 폴더."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _apply_qss(app: QtWidgets.QApplication, qss_path: Path) -> None:
    """QSS가 있으면 적용(없으면 무시)."""
    try:
        if qss_path.is_file():
            app.setStyleSheet(qss_path.read_text(encoding="utf-8"))
    except Exception:
        # QSS 오류는 치명적이지 않게 무시
        pass


def _load_module_from_file(module_name: str, file_path: Path):
    """
    module\excel_ui.py를 '파일 경로 기준'으로 강제 로딩.
    - module 폴더를 sys.path 최우선(0)로 넣어, module 내부 모듈 import가 안정적으로 되게 함.
    - 동일 모듈명이 이미 로드돼 있으면 제거(혼종/캐시 방지).
    """
    module_root = str(file_path.parent)
    if not sys.path or sys.path[0] != module_root:
        sys.path.insert(0, module_root)

    if module_name in sys.modules:
        del sys.modules[module_name]

    spec = importlib.util.spec_from_file_location(module_name, str(file_path))
    if spec is None or spec.loader is None:
        raise RuntimeError(f"spec 생성 실패: {file_path}")

    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _fatal_dialog(title: str, message: str) -> None:
    QtWidgets.QMessageBox.critical(None, title, message)


def main():
    root = _runtime_root()
    module_dir = (root / MODULE_DIRNAME).resolve()

    entry_file = (module_dir / ENTRY_FILENAME).resolve()
    qss_file = (module_dir / QSS_FILENAME).resolve()   # ✅ QSS는 module에서 읽는다

    # 1) 경로 체크
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)

    if not module_dir.is_dir():
        _fatal_dialog("실행 오류", f"'{MODULE_DIRNAME}' 폴더가 없습니다.\n\n{module_dir}")
        return

    if not entry_file.is_file():
        _fatal_dialog("실행 오류", f"엔트리 파일이 없습니다.\n\n{entry_file}")
        return

    # 2) QSS 적용
    _apply_qss(app, qss_file)

    # 3) 외부 엔트리 로드 + 실행
    try:
        entry = _load_module_from_file(ENTRY_MODNAME, entry_file)

        if not hasattr(entry, "main"):
            raise AttributeError(f"{ENTRY_FILENAME}에 main() 함수가 없습니다.")

        entry.main()

    except Exception:
        _fatal_dialog("치명적 오류", traceback.format_exc())
        return

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
