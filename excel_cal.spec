# excel_cal.spec
# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules, collect_all


def add_pkg(pkg: str):
    try:
        return collect_submodules(pkg)
    except Exception:
        return []


hiddenimports = []
datas = [("icon.ico", ".")]
binaries = []

# -------------------------------------------------
# 외부 module 폴더에서 사용하는 서드파티 라이브러리 포함
# (새 라이브러리 추가 시, 여기에 pkg 추가 후 재빌드 1회)
# -------------------------------------------------

# ✅ Google API / OAuth
for pkg in [
    "google",
    "googleapiclient",
    "google_auth_oauthlib",
    "google.auth",
]:
    hiddenimports += add_pkg(pkg)

# ✅ 일반 패키지들
for pkg in [
    "requests",
    "pandas",
    "openpyxl",
    "cryptography",
    "git",
]:
    hiddenimports += add_pkg(pkg)

# ✅ msoffcrypto: 모듈+데이터+숨김임포트까지 싹 수집
msoff_bins, msoff_datas, msoff_hidden = collect_all("msoffcrypto")
binaries += msoff_bins
datas += msoff_datas
hiddenimports += msoff_hidden

# ✅ olefile: msoffcrypto 내부 의존성(설치돼 있어도 번들 누락될 수 있어 강제 포함)
ole_bins, ole_datas, ole_hidden = collect_all("olefile")
binaries += ole_bins
datas += ole_datas
hiddenimports += ole_hidden

# (선택) 그래도 안전빵으로 문자열 강제 포함하고 싶으면 켜도 됨
# hiddenimports += ["msoffcrypto", "olefile"]

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="excel_cal",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=["icon.ico"],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="excel_cal",
)
