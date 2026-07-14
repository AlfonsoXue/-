# -*- mode: python ; coding: utf-8 -*-
# ============================================================
# EVA 旅客滿意度儀表板 — PyInstaller 打包設定(單機免安裝、單一 .exe)
# 用法： pyinstaller --noconfirm EVA_Dashboard.spec
#       (或) python -m PyInstaller --noconfirm EVA_Dashboard.spec
# 產出： dist\EVA滿意度儀表板.exe
# 說明：pandas / numpy 改用 PyInstaller 內建掛勾自動處理(不再掃測試模組，
#       可避免 "No module named 'pytest'" 警告，並縮小體積)。
# ============================================================
from PyInstaller.utils.hooks import collect_all

# openpyxl 需要其資料檔，整包收集；build_dashboard 為引擎(同資料夾)
datas, binaries, hiddenimports = [], [], ['build_dashboard']
_d, _b, _h = collect_all('openpyxl')
datas += _d; binaries += _b; hiddenimports += _h

a = Analysis(
    ['EVA_儀表板GUI.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    # 排除測試/未用的大型套件 → 去除警告、縮小體積
    excludes=['matplotlib', 'PyQt5', 'PySide2', 'pytest', '_pytest',
              'pandas.tests', 'numpy.tests', 'numpy.f2py.tests', 'test', 'tests'],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name='EVA滿意度儀表板',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                 # 純 GUI；除錯時可暫改 True
    disable_windowed_traceback=False,
    icon='eva.ico',                # exe 圖示(置於與 .spec 同資料夾)；改用其他檔改此處
)
