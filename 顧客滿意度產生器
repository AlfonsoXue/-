# -*- coding: utf-8 -*-
"""
長榮航空 旅客滿意度互動分析儀表板產生器
EVA Air Passenger Satisfaction Interactive Analysis Dashboard Generator
------------------------------------------------------------------------
讀取問卷原始匯出檔，分析 會員卡別/國籍/搭機年月/班機編號/機型/艙等/起飛城市 等
維度之對應關係，計算「滿意度分數、回覆人數、3分(含)以下原因解析」，
並輸出具備「資料驗證下拉式選單」之互動 Excel 檔，方便同仁快速撈取與洞悉資料。

可重複使用(每月作業)：
  1. 把當月的 DataExport*.xlsx(例：DataExportMAY2026.xlsx)放進
     「滿意度資料庫 ＼ RAW DATA」子資料夾，歷月檔案皆保留於該資料夾。
  2. 直接執行本程式(免帶任何參數)：python build_dashboard.py
     → 自動讀取 RAW DATA 內【全部】DataExport 檔，合併所有年份/月份一起分析，
       將互動 Excel 輸出至「滿意度資料庫」根目錄。
       單月時檔名為該期間(例 …_2026年5月.xlsx)；跨多月時檔名為 …_全期間.xlsx。
     儀表板提供「搭機年度」「搭機月份」下拉選單，可隨時回切到單一月別檢視。
  進階：亦可手動指定單一檔 → python build_dashboard.py 輸入檔.xlsx [輸出檔.xlsx]
"""
import sys, os, glob, re, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.workbook.defined_name import DefinedName


def main():
    """建立/更新儀表板的主流程(可被 GUI 以 in-process 方式重複呼叫)。
    執行行為由環境變數控制：EVA_DB_ROOT(資料庫根)、EVA_OUTPUT_FILE(匯出檔)、
    EVA_UPDATE_MODE(FULL/DATA_ONLY)、EVA_REBUILD(=1 強制全重建)。"""

    # ============================ 路徑設定 ============================
    # 資料庫根目錄。原始匯出檔放在其下的「RAW DATA」子資料夾(檔名 DataExportMMMYYYY.xlsx，
    # 例：DataExportMAY2026.xlsx)；互動分析成品則輸出到根目錄本身。
    # 如需暫時改用其他資料夾，可設定環境變數 EVA_DB_ROOT 覆蓋(免改程式)。
    _DB_ROOT   = os.environ.get("EVA_DB_ROOT") or \
        r"\\omdfile\omd\PMS\17.WS服務檢討會-顧客滿意度與OCSS\01 全球滿意度彙整\01 顧客滿意度統計\滿意度資料庫"
    INPUT_DIR  = os.path.join(_DB_ROOT, "RAW DATA")   # 原始匯出檔(DataExport*.xlsx)所在子資料夾
    OUTPUT_DIR = _DB_ROOT                              # 互動分析成品輸出位置(根目錄)

    def _resolve_inputs(arg):
        """決定要分析的輸入檔(可多檔)：
           (0) 環境變數 EVA_INPUT_FILES 明確指定(以系統路徑分隔)→ 直接採用該些檔(GUI 自選檔案，免資料夾)。
           (1) 命令列有給 → 只分析該檔(完整路徑直接用；只給檔名則接到 INPUT_DIR)。
           (2) 未給 → 抓 INPUT_DIR 內【全部】DataExport*.xlsx，合併所有年份/月份一起分析。"""
        env_files = os.environ.get("EVA_INPUT_FILES")
        if env_files:
            files = [p for p in env_files.split(os.pathsep) if p.strip()]
            _miss = [p for p in files if not os.path.exists(p)]
            if _miss:
                raise FileNotFoundError("找不到指定的輸入檔：" + "; ".join(_miss))
            if files:
                return files
        if arg:
            return [arg if os.path.dirname(arg) else os.path.join(INPUT_DIR, arg)]
        cands = sorted(glob.glob(os.path.join(INPUT_DIR, "DataExport*.xlsx")))
        if not cands:
            raise FileNotFoundError(
                f"在 {INPUT_DIR} 找不到 DataExport*.xlsx，請確認匯出檔已放入 RAW DATA 子資料夾，"
                f"或於命令列指定輸入檔：python build_dashboard.py 完整檔名.xlsx")
        return cands

    # ============================ 更新行為開關 (可自行切換) ============================
    # 當「成品檔已存在」且未加 --rebuild 時，重跑本程式的行為由 UPDATE_MODE 決定：
    #   "DATA_ONLY" → 僅刷新『滿意度明細』『原因明細』兩張明細，並同步更新『清單』類別欄與
    #                 各分頁下拉選單範圍；其餘分析頁之人工維護(備註、目標值編修等)一律保留。
    #   "FULL"      → 以最新資料【重建所有頁籤】(等同全新建檔)；目標值等取程式內建設定，
    #                 不保留 Excel 內對分析頁的人工編輯。日後若需改目標值，請改本程式設定。
    # 亦可用環境變數覆蓋(免改程式)：EVA_UPDATE_MODE=DATA_ONLY 或 FULL。
    UPDATE_MODE = "FULL"
    UPDATE_MODE = (os.environ.get("EVA_UPDATE_MODE") or UPDATE_MODE).strip().upper().replace(" ", "_")
    if UPDATE_MODE not in ("DATA_ONLY", "FULL"):
        print(f"　⚠ UPDATE_MODE 設定值『{UPDATE_MODE}』無效，已改用預設 FULL(重建所有頁籤)。")
        UPDATE_MODE = "FULL"

    # ---- 模式旗標：--rebuild 或環境變數 EVA_REBUILD=1 → 強制完整重建所有分頁(等同 FULL) ----
    _REBUILD = bool(os.environ.get("EVA_REBUILD"))
    if "--rebuild" in sys.argv:
        _REBUILD = True
        sys.argv = [a for a in sys.argv if a != "--rebuild"]

    INPUT_FILES = _resolve_inputs(sys.argv[1] if len(sys.argv) > 1 else None)
    print(f"共偵測到 {len(INPUT_FILES)} 個輸入檔：")
    for _f in INPUT_FILES:
        print("   -", os.path.basename(_f))
    # OUTPUT_FILE 於下方取得資料期間後再決定，輸出至 OUTPUT_DIR(資料庫根目錄)。

    # ============================ 1. 讀取與整理資料 ============================
    SAT_SHEET = "滿意度題目(Satisfaction Question)"
    RSN_SHEET = "非滿意度題目(Non Satisfaction Questio"

    # ---- 資料清理模式(整合自 survey_cleaner)：EVA_CLEAN = auto / 1(force) / 0(off) ----
    #   auto  ：逐檔自動偵測——僅當有 CONFIDENTIAL 浮水印或標題不在第 1 列時才清理；
    #           已清理乾淨(或偵測不到需清理之處)者直接讀取產製。
    #   force ：一律以第 CLEAN_HEADER1 列為標題清理。
    #   off   ：完全不清理，直接讀取。
    #   清理動作：以標題列為準、移除無標題空白欄、只保留指定題號(預設 14/16/17/18/26)；
    #   找不到「題號」欄的工作表(如非滿意度表)原樣保留。題號/分數維持原數值型別。
    _clean_raw = str(os.environ.get("EVA_CLEAN", "")).strip().lower()
    if _clean_raw == "auto":
        CLEAN_MODE = "auto"
    elif _clean_raw in ("1", "true", "on", "yes", "force"):
        CLEAN_MODE = "force"
    else:
        CLEAN_MODE = "off"
    CLEAN_ENABLED = CLEAN_MODE == "force"          # 向後相容
    CLEAN_KEEP = {p for p in re.split(r"[,\s、，;；]+", (os.environ.get("EVA_CLEAN_KEEP") or "14 16 17 18 26").strip()) if p} or {"14", "16", "17", "18", "26"}
    CLEAN_QCOL = (os.environ.get("EVA_CLEAN_QCOL") or "題號").strip()
    try:
        CLEAN_HEADER1 = int(os.environ.get("EVA_CLEAN_HEADER") or 3)
        CLEAN_HEADER = max(0, CLEAN_HEADER1 - 1)
    except ValueError:
        CLEAN_HEADER1, CLEAN_HEADER = 3, 2
    CLEAN_SUBITEMS = str(os.environ.get("EVA_CLEAN_SUBITEMS", "")).strip().lower() in ("1", "true", "on", "yes")
    # 清理後另存乾淨檔：預設存到網路資料夾，檔名依資料月份自動命名為「全球顧客滿意度 (MMMYYYY)_cleaned.xlsx」
    _DEFAULT_CLEAN_DIR = (r"\\omdfile\omd\PMS\17.WS服務檢討會-顧客滿意度與OCSS"
                          r"\01 全球滿意度彙整\01 顧客滿意度統計\滿意度資料庫\RAW DATA")
    CLEAN_SAVE = str(os.environ.get("EVA_CLEAN_SAVE", "1")).strip().lower() in ("1", "true", "on", "yes")
    CLEAN_SAVE_DIR = (os.environ.get("EVA_CLEAN_SAVE_DIR") or _DEFAULT_CLEAN_DIR).strip()
    _MON3 = {1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN",
             7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"}
    # 自動偵測標題列用的「已知欄名」(出現足夠多即視為標題列)
    _KNOWN_HEADERS = {"問卷編號", "會員卡別", "國籍", "搭機年度", "搭機月份", "班機編號",
                      "機型", "艙等", "起飛城市", "填答來源", "性別", "構面", "服務項目",
                      "題號", "分數", "關聯滿意度分數", "選項", CLEAN_QCOL}

    def _find_qcol(df, qcol):
        target = str(qcol).strip()
        for c in df.columns:
            if str(c).strip() == target:
                return c
        return None

    def _clean_df(df):
        """清理單一工作表：去除標題前後空白、移除無標題空白欄、依題號保留指定題目。
        無題號欄的工作表(如非滿意度表)原樣保留(僅去空白欄)。題號/分數等維持原數值型別。"""
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        col = _find_qcol(df, CLEAN_QCOL)
        if col is None:
            return df
        qn = df[col].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)  # 16.0→16，保留 16.1
        if CLEAN_SUBITEMS:
            qn = qn.str.split(".").str[0]                                          # 含子題：16.1→16
        return df[qn.isin(CLEAN_KEEP)].copy()

    def _period_tag(d1):
        """由清理後的滿意度資料推導 MMMYYYY(如 APR2026)；無法判定回傳 UNKNOWN。"""
        try:
            yr = pd.to_numeric(d1.get("搭機年度"), errors="coerce").dropna()
            mo = pd.to_numeric(d1.get("搭機月份"), errors="coerce").dropna()
            if len(yr) and len(mo):
                y = int(yr.mode().iloc[0]); m = int(mo.mode().iloc[0])
                return f"{_MON3.get(m, str(m))}{y}"
        except Exception:
            pass
        return "UNKNOWN"

    def _save_cleaned(d1, d2):
        """將清理後的兩個工作表另存為乾淨 DataExport 檔到 CLEAN_SAVE_DIR。
        檔名：全球顧客滿意度 (MMMYYYY)_cleaned.xlsx。回傳路徑或 None(未啟用/失敗)。"""
        if not CLEAN_SAVE:
            return None
        fname = f"全球顧客滿意度 ({_period_tag(d1)})_cleaned.xlsx"
        try:
            os.makedirs(CLEAN_SAVE_DIR, exist_ok=True)
        except Exception as e:
            print(f"　⚠ 無法建立/存取清理檔資料夾，略過另存：{CLEAN_SAVE_DIR}（{e}）")
            return None
        out = os.path.join(CLEAN_SAVE_DIR, fname)
        try:
            with pd.ExcelWriter(out, engine="openpyxl") as w:
                d1.to_excel(w, sheet_name=SAT_SHEET[:31], index=False)
                d2.to_excel(w, sheet_name=RSN_SHEET[:31], index=False)
            print(f"　💾 已另存清理檔：{out}")
            return out
        except Exception as e:
            print(f"　⚠ 清理檔另存失敗(不影響儀表板產製)：{out}（{e}）")
            return None

    # ---- 受保護/加密檔案自動解鎖：密碼解密(優先 msoffcrypto)或 Excel COM 解除 IRM/權限 ----
    FILE_PASSWORD = os.environ.get("EVA_FILE_PASSWORD") or ""   # 開啟密碼(常為人事代號)，由 GUI 帶入
    _xl = {"app": None}

    def _make_excel():
        """建立 Excel.Application COM 物件，能容忍 gen_py 快取毀損。
        先試早期繫結(EnsureDispatch)；若快取毀損(如 CLSIDToClassMap 錯誤)則清快取重試，
        仍失敗則改用純動態晚期繫結(dynamic.Dispatch)，完全不依賴 gen_py 快取。"""
        import win32com.client as win32           # 需 Windows + Excel + pywin32
        try:
            return win32.gencache.EnsureDispatch("Excel.Application")
        except Exception:
            pass
        # 清除毀損的 gen_py 快取後再試一次早期繫結
        try:
            import win32com as _w32, shutil
            gp = getattr(_w32, "__gen_path__", None)
            if gp:
                shutil.rmtree(gp, ignore_errors=True)
            # 移除已載入的 gen_py 模組參照，迫使重建
            for _mod in [m for m in list(sys.modules) if "win32com.gen_py" in m]:
                sys.modules.pop(_mod, None)
        except Exception:
            pass
        try:
            import importlib
            import win32com.client as win32b
            importlib.reload(win32b.gencache)
            return win32b.gencache.EnsureDispatch("Excel.Application")
        except Exception:
            pass
        # 最後手段：純動態晚期繫結(不使用 gen_py)
        import win32com.client.dynamic as _dyn
        return _dyn.Dispatch("Excel.Application")

    def _get_excel():
        if _xl["app"] is None:
            app = _make_excel()
            app.Visible = False
            app.DisplayAlerts = False
            for _attr, _val in (("AskToUpdateLinks", False), ("EnableEvents", False),
                                ("ScreenUpdating", False)):
                try:
                    setattr(app, _attr, _val)               # 盡量抑制各種互動式提示視窗
                except Exception:
                    pass
            _xl["app"] = app
        return _xl["app"]

    def _quit_excel():
        app = _xl.get("app")
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
            _xl["app"] = None

    def _excel_com_unprotect(src_path, password=None):
        """以背景 Excel 開啟受保護檔，解除 IRM/權限後另存乾淨 .xlsx 暫存檔。
        關鍵：一律帶入 Password 參數(沒有真密碼就帶佔位字串)，使受密碼保護的檔案
        會丟出可攔截的錯誤，而非跳出 Excel 的輸入密碼視窗。"""
        import tempfile
        nm = os.path.basename(src_path)
        try:
            excel = _get_excel()
        except Exception as e:
            raise RuntimeError(
                f"需透過 Excel 解鎖，但本機無法啟動 Excel COM（需 Windows + 已安裝 Excel + pywin32）。原因：{e}")
        tmp = tempfile.mktemp(suffix=".xlsx", prefix="eva_unlocked_")
        # 沒有真密碼時帶入一個不可能正確的佔位密碼 → Excel 不會跳視窗，而是回報錯誤讓我們攔截
        open_pw = password if password else "__eva_no_password__"
        try:
            # 位置參數：Open(FileName, UpdateLinks=0, ReadOnly=True, Format=None, Password)
            # 一律帶 Password(第5個)，使受密碼保護的檔不會跳出輸入密碼視窗；在早/晚期繫結皆可用
            wb = excel.Workbooks.Open(os.path.abspath(src_path), 0, True, None, open_pw)
        except Exception as e:
            if password:
                raise RuntimeError(
                    f"以 Excel 開啟「{nm}」失敗：密碼(人事代號)可能不正確"
                    f"（注意大小寫與前置 0）。原因：{e}")
            raise RuntimeError(
                f"檔案「{nm}」受密碼保護，需要密碼才能開啟。\n"
                f"　→ 請在『清理設定…』的「檔案密碼」輸入此檔密碼(通常為人事代號)後重試；"
                f"程式會背景自動帶入，不會跳出 Excel 密碼視窗。")
        try:
            wb.Permission.Enabled = False           # 解除 IRM 權限保護(若有)
        except Exception:
            pass
        wb.SaveAs(os.path.abspath(tmp), 51)         # 位置參數：FileFormat=51 (.xlsx)
        wb.Close(False)                             # 位置參數：SaveChanges=False
        return tmp

    def _is_xlsx_zip(p):
        """嚴謹檢查是否為『完整可解析』的 ZIP(xlsx)：不只看開頭 PK，
        而是驗證中央目錄結尾紀錄是否存在(zipfile.is_zipfile)。"""
        try:
            import zipfile
            return zipfile.is_zipfile(p)
        except Exception:
            try:
                with open(p, "rb") as fh:
                    return fh.read(4) == b"PK\x03\x04"
            except Exception:
                return False

    def _is_encrypted(p):
        """判斷檔案是否為『密碼加密』的 Office 檔。回傳 True/False；無法判斷時回傳 None。"""
        # 1) msoffcrypto(若有)最準
        try:
            import msoffcrypto
            with open(p, "rb") as fh:
                return bool(msoffcrypto.OfficeFile(fh).is_encrypted())
        except ImportError:
            pass
        except Exception:
            pass
        # 2) 後援：掃描 OLE2 內是否含加密串流名(UTF-16LE)；掃整個檔以免大檔漏判
        try:
            with open(p, "rb") as fh:
                blob = fh.read()
            return (("EncryptedPackage".encode("utf-16-le") in blob) or
                    ("EncryptionInfo".encode("utf-16-le") in blob))
        except Exception:
            return None

    def _decrypt_with_password(src_path, password):
        """用開啟密碼解密受保護檔，另存乾淨 .xlsx 暫存檔並回傳路徑。
        優先用純 Python 的 msoffcrypto(免 Excel)；輸出非有效 xlsx 或未安裝時，改用 Excel COM 帶密碼開啟。"""
        import tempfile
        name = os.path.basename(src_path)
        try:
            import msoffcrypto
        except ImportError:
            msoffcrypto = None
        if msoffcrypto is not None:
            tmp = tempfile.mktemp(suffix=".xlsx", prefix="eva_dec_")
            decrypted = False
            try:
                with open(src_path, "rb") as fin:
                    off = msoffcrypto.OfficeFile(fin)
                    off.load_key(password=password)
                    with open(tmp, "wb") as fout:
                        off.decrypt(fout)
                decrypted = True
            except Exception as e:
                try:
                    os.remove(tmp)
                except Exception:
                    pass
                _m = str(e).lower()
                if any(k in _m for k in ("password", "key", "incorrect", "verifier")):
                    raise RuntimeError(
                        f"檔案「{name}」解鎖失敗：密碼(人事代號)不正確，請確認後重試"
                        f"（注意大小寫與前置 0）。")
                # 其他原因(可能非密碼加密/格式特殊)→ 落到 Excel COM 後援
            if decrypted:
                if _is_xlsx_zip(tmp):
                    return tmp                      # 解密成功且為有效 xlsx
                try:
                    os.remove(tmp)
                except Exception:
                    pass
                # 解密「成功」但內容非有效 xlsx：多半是密碼不符(部分加密法不報錯)或舊格式
                # → 改用 Excel COM 帶密碼開啟(較相容)，若仍失敗會丟出清楚訊息
        return _excel_com_unprotect(src_path, password=password)

    def _resolve_readable(path):
        """回傳 (可讀路徑, 是否為暫存檔)：原檔可直接讀就用原檔；若受保護/毀損/為 OLE2 外殼，
        則自動以密碼解鎖或 Excel 重新另存乾淨 xlsx。"""
        name = os.path.basename(path)
        try:
            sz = os.path.getsize(path)
        except Exception:
            sz = -1
        try:
            with open(path, "rb") as fh:
                head = fh.read(8)
        except OSError as e:
            raise RuntimeError(f"無法開啟檔案「{name}」：{e}")

        is_zip = head[:4] == b"PK\x03\x04"                               # 真正的 .xlsx/.xlsm(ZIP)
        is_ole2 = head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"        # OLE2(舊式 .xls 或受保護的 xlsx 外殼)
        _t = head.lstrip().lower()
        is_html = _t.startswith(b"<") or _t.startswith(b"<?xml") or _t.startswith(b"<!doctype")

        try:
            import msoffcrypto  # noqa
            _has_mso = True
        except Exception:
            _has_mso = False
        _kind = ("ZIP/xlsx" if is_zip else "OLE2(加密或舊xls)" if is_ole2
                 else "HTML/XML" if is_html else "其他/未知")
        print(f"　🔍 檔案診斷「{name}」：前綴={head[:4].hex(' ')} 大小={sz} bytes "
              f"類型={_kind} msoffcrypto={'有' if _has_mso else '無'} "
              f"已輸入密碼={'是' if FILE_PASSWORD else '否'}")

        if is_zip:
            # 開頭是 PK，但仍須確認是『完整』ZIP；毀損/特殊壓縮 → 嘗試以 Excel 救援
            import zipfile
            if zipfile.is_zipfile(path):
                return path, False
            print(f"　⚠ 「{name}」開頭為 ZIP 但結構不完整(可能毀損或特殊壓縮)，"
                  f"嘗試以背景 Excel 重新另存為乾淨 xlsx…")
            tmp = _excel_com_unprotect(path, password=(FILE_PASSWORD or None))
            print(f"　　✓ 已重新另存(暫存檔)，讀取中…")
            return tmp, True
        if is_ole2:
            # (a) 先試是否為可直接讀取的舊式 .xls
            try:
                pd.ExcelFile(path, engine="xlrd").sheet_names
                return path, False
            except Exception:
                pass
            # (b) 有提供開啟密碼(人事代號)→ 以密碼解鎖(背景帶入，不跳視窗)
            if FILE_PASSWORD:
                print(f"　🔓 「{name}」受密碼保護，使用所輸入之密碼(人事代號)背景解鎖…")
                tmp = _decrypt_with_password(path, FILE_PASSWORD)
                print(f"　　✓ 已解鎖(暫存檔)，讀取中…")
                return tmp, True
            # (c) 沒給密碼：先判斷是否為『密碼加密』。是 → 直接擋下並提示(避免 Excel 跳出輸入密碼視窗)
            if _is_encrypted(path):
                raise RuntimeError(
                    f"檔案「{name}」受密碼保護，但尚未提供密碼。\n"
                    f"　→ 請在『清理設定…』的「檔案密碼」輸入此檔密碼(通常為人事代號)後再產生；"
                    f"程式會背景自動帶入解鎖，不會跳出 Excel 的密碼視窗。")
            # (d) 非密碼加密(視為 IRM/權限保護)→ 以 Excel COM 自動解除
            print(f"　🔓 「{name}」疑似受權限/IRM 保護，改用背景 Excel 自動解除保護後讀取…")
            tmp = _excel_com_unprotect(path)
            print(f"　　✓ 已解除保護(暫存檔)，讀取中…")
            return tmp, True
        if is_html:
            raise RuntimeError(
                f"檔案「{name}」其實是 HTML 表格（副檔名雖為 Excel），無法區分"
                f"『{SAT_SHEET}』與『{RSN_SHEET}』兩個工作表。\n"
                f"　→ 請在 Excel 開啟後另存為真正的 .xlsx（需含上述兩個工作表）再匯入。")
        # 其他/未知：嘗試以 Excel 開啟另存(說不定可救)，否則明確擋下
        try:
            print(f"　🔓 「{name}」格式未知，嘗試以背景 Excel 開啟並另存為 xlsx…")
            tmp = _excel_com_unprotect(path)
            return tmp, True
        except RuntimeError:
            raise RuntimeError(
                f"無法讀取檔案「{name}」（非有效 Excel/可能毀損），且無法以 Excel 自動解除。\n"
                f"　→ 請確認為有效的 .xlsx；必要時在 Excel 重新另存後再匯入。")

    def _eng(rp):
        return "openpyxl" if rp.lower().endswith((".xlsx", ".xlsm")) else None

    def _read_sheet(rp, src_name, sheet_name, header=0):
        try:
            return pd.read_excel(rp, sheet_name=sheet_name, engine=_eng(rp), header=header)
        except ValueError as e:
            if "not found" in str(e).lower() or "worksheet" in str(e).lower():
                try:
                    avail = "、".join(pd.ExcelFile(rp, engine=_eng(rp)).sheet_names)
                except Exception:
                    avail = "(無法列出)"
                raise RuntimeError(
                    f"檔案「{src_name}」中找不到工作表『{sheet_name}』。\n"
                    f"　該檔實際工作表：{avail}\n"
                    f"　→ 請確認匯出的是正確的 DataExport 檔（需同時含滿意度／非滿意度兩個工作表）。")
            raise
        except Exception as e:
            _m = str(e).lower()
            if "not a zip" in _m or e.__class__.__name__ == "BadZipFile":
                _hint = ("已提供密碼但仍無法讀取，請確認密碼(人事代號)正確"
                         if FILE_PASSWORD else
                         "若此檔受密碼保護，請在『清理設定…』輸入正確的人事代號當密碼")
                raise RuntimeError(
                    f"檔案「{src_name}」不是有效的 .xlsx（ZIP）。\n"
                    f"　常見原因：(1) 受密碼保護但密碼不符；(2) 檔案毀損；(3) 為『假 Excel』/HTML/舊式 .xls。\n"
                    f"　→ {_hint}；否則請在 Excel 另存為 .xlsx 後再匯入。")
            raise

    def _detect_sheet(rp, sheet_name, max_scan=15):
        """掃描某工作表前 max_scan 列(不設標題)，回傳 (標題列 0-based 索引 或 None, 是否有 CONFIDENTIAL 浮水印)。
        標題列判定：該列包含『題號欄名』，或包含 >=3 個已知欄名。"""
        try:
            raw = pd.read_excel(rp, sheet_name=sheet_name, header=None,
                                nrows=max_scan, engine=_eng(rp))
        except Exception:
            return None, False
        header_row, watermark = None, False
        for i in range(len(raw)):
            vals = [str(v).strip() for v in raw.iloc[i].tolist() if pd.notna(v)]
            if any("confidential" in v.lower() for v in vals):
                watermark = True
            sval = set(vals)
            if header_row is None and (CLEAN_QCOL in sval or len(sval & _KNOWN_HEADERS) >= 3):
                header_row = i
        return header_row, watermark

    def _auto_plan(rp, name):
        """自動判定某檔是否需要清理，以及兩個工作表的標題列。
        回傳 (need_clean, hdr_sat, hdr_rsn)。"""
        hr1, wm1 = _detect_sheet(rp, SAT_SHEET)
        hr2, wm2 = _detect_sheet(rp, RSN_SHEET)
        watermark = wm1 or wm2
        sat_off = (hr1 is not None and hr1 >= 1)     # 標題不在第 1 列
        rsn_off = (hr2 is not None and hr2 >= 1)
        need = bool(watermark or sat_off or rsn_off)
        hdr1 = hr1 if (hr1 is not None) else 0
        hdr2 = hr2 if (hr2 is not None) else 0
        _why = []
        if watermark:
            _why.append("偵測到 CONFIDENTIAL 浮水印")
        if sat_off or rsn_off:
            _why.append(f"標題不在第 1 列(滿意度第 {hdr1+1} 列、非滿意度第 {hdr2+1} 列)")
        if need:
            print(f"　🔎 「{name}」自動判定：需要清理（{'；'.join(_why)}）")
        else:
            print(f"　🔎 「{name}」自動判定：已是乾淨格式(標題在第 1 列、無浮水印)，直接讀取產製，不清理")
        return need, hdr1, hdr2

    # 逐檔解析→(必要時自動解保護)→讀取兩個工作表→(視模式清理)；最後關閉背景 Excel
    if CLEAN_MODE == "auto":
        print(f"　🧹 清理模式：自動偵測（僅在有浮水印或標題不在第 1 列時清理；保留題號 {'/'.join(sorted(CLEAN_KEEP))}）")
    elif CLEAN_ENABLED:
        print(f"　🧹 清理模式：一律清理（標題第 {CLEAN_HEADER1} 列、保留題號 {'/'.join(sorted(CLEAN_KEEP))}、{'含' if CLEAN_SUBITEMS else '不含'}子題）")
    _frames1, _frames2 = [], []
    try:
        for _f in INPUT_FILES:
            _name = os.path.basename(_f)
            _rp, _is_tmp = _resolve_readable(_f)
            try:
                # 決定本檔是否清理、及兩工作表標題列
                if CLEAN_MODE == "auto":
                    _do_clean, _hdr1, _hdr2 = _auto_plan(_rp, _name)
                elif CLEAN_MODE == "force":
                    _do_clean, _hdr1, _hdr2 = True, CLEAN_HEADER, CLEAN_HEADER
                else:
                    _do_clean, _hdr1, _hdr2 = False, 0, 0
                _d1 = _read_sheet(_rp, _name, SAT_SHEET, header=_hdr1)
                _d2 = _read_sheet(_rp, _name, RSN_SHEET, header=_hdr2)
                if _do_clean:
                    _b1, _b2 = len(_d1), len(_d2)
                    _d1, _d2 = _clean_df(_d1), _clean_df(_d2)
                    print(f"　🧹 {_name}：滿意度 {_b1:,}→{len(_d1):,} 列、非滿意度 {_b2:,}→{len(_d2):,} 列")
                    _save_cleaned(_d1, _d2)        # 清理後另存乾淨檔到指定資料夾
                _frames1.append(_d1)
                _frames2.append(_d2)
            finally:
                if _is_tmp:
                    try:
                        os.remove(_rp)
                    except Exception:
                        pass
    finally:
        _quit_excel()

    df1 = pd.concat(_frames1, ignore_index=True)
    df2 = pd.concat(_frames2, ignore_index=True)
    print(f"合併後：滿意度題目 {len(df1):,} 列、非滿意度題目 {len(df2):,} 列")

    # 構面(題號 -> 友善中文名稱)
    QNUM_TO_NAME = {14: "報到櫃檯服務", 16: "貴賓室人員服務", 17: "貴賓室環境",
                    18: "登機服務", 26: "客艙清潔"}
    DIM_ORDER = ["報到櫃檯服務", "貴賓室人員服務", "貴賓室環境", "登機服務", "客艙清潔"]

    # 年度趨勢／目標達成用：服務項目(2026 起新增合併項目) -> 對應之原始構面(可一或多)
    #   機場臨櫃報到+登機服務 = 報到櫃檯服務 + 登機服務
    #   貴賓室服務            = 貴賓室人員服務 + 貴賓室環境
    YEAR_ITEMS = [
        ("機場報到櫃台",          ["報到櫃檯服務"]),
        ("登機服務",              ["登機服務"]),
        ("貴賓室人員服務",        ["貴賓室人員服務"]),
        ("貴賓室環境",            ["貴賓室環境"]),
        ("客艙清潔",              ["客艙清潔"]),
        ("機場臨櫃報到+登機服務",  ["報到櫃檯服務", "登機服務"]),
        ("貴賓室服務",            ["貴賓室人員服務", "貴賓室環境"]),
    ]
    YEAR_ROWS = ["2024", "2025", "2026"]
    # 各年度各項目「目標值」(成品中可自行編修；None=該年度尚未適用，例如合併項目自2026起)
    YEAR_TARGETS = {
        "2024": [4.37, 4.37, 4.37, 4.37, 4.37, None, None],
        "2025": [4.37, 4.37, 4.37, 4.37, 4.37, None, None],
        "2026": [4.49, 4.49, 4.25, 4.25, 4.41, 4.49, 4.25],
    }
    MIN_N = 5   # 回覆人數 <= MIN_N 不列計滿意度平均(但回覆人數仍顯示)

    # 各構面/合併項目之「當年度目標值」(供所有頁籤分數依目標雙向著色)。取自 YEAR_TARGETS 最新年度。
    _TGT_YEAR = YEAR_ROWS[-1]                       # "2026"
    TARGET_NOW = {}
    for (_nm, _cs), _tg in zip(YEAR_ITEMS, YEAR_TARGETS.get(_TGT_YEAR, [])):
        if _tg is None:
            continue
        TARGET_NOW[_cs[0] if len(_cs) == 1 else _nm] = _tg
    # 結果：報到櫃檯服務/登機服務=4.49、貴賓室人員服務/貴賓室環境=4.25、客艙清潔=4.41、
    #       機場臨櫃報到+登機服務=4.49、貴賓室服務=4.25

    # ---- 「年度目標達成」頁版面列位(提前計算，供所有頁籤引用第三點『目標值』儲存格著色) ----
    #   各頁分數儲存格之漸層著色一律參照此頁第三點「目標值」之當年度該列，改一處、全簿色彩連動。
    YSHEET = "年度目標達成"
    _Y_SC_HR, _Y_SC_R0 = 5, 6                                              # 平均分數
    _Y_CN_HR, _Y_CN_R0 = _Y_SC_R0 + len(YEAR_ROWS) + 2, _Y_SC_R0 + len(YEAR_ROWS) + 3   # 回覆人數
    _Y_TG_HR, _Y_TG_R0 = _Y_CN_R0 + len(YEAR_ROWS) + 2, _Y_CN_R0 + len(YEAR_ROWS) + 3   # 目標值(第三點)
    _Y_TG_CUR_ROW = _Y_TG_R0 + (len(YEAR_ROWS) - 1)                        # 當年度(最新)目標值該列
    # 構面/合併項目 → 該頁第三點當年度「目標值」儲存格參照(供 add_target_gradient 引用)
    ITEM_TGT_CELL = {}
    ITEM_TGT_COL = {}        # 構面/項目 → 在第三點之欄字母(供互動查詢依年度動態取目標)
    for _j, (_nm, _cs) in enumerate(YEAR_ITEMS):
        _key = _cs[0] if len(_cs) == 1 else _nm
        _colL = get_column_letter(3 + _j)
        ITEM_TGT_COL[_key] = _colL
        ITEM_TGT_CELL[_key] = f"'{YSHEET}'!${_colL}${_Y_TG_CUR_ROW}"

    # 區域(依起飛城市歸併)。順序：東北亞、東南亞、歐洲/澳洲、北美、台灣、中國/香港/澳門
    REGION_ORDER = ["東北亞", "東南亞", "歐洲/澳洲", "北美", "台灣", "中國/香港/澳門"]
    _REGION_CITIES = {
        "台灣":          ["TPE", "TSA", "KHH"],
        "東北亞":        ["NRT", "HND", "KIX", "CTS", "FUK", "OKA", "SDJ", "KMQ", "AOJ",
                         "UKB", "MYJ", "ICN", "GMP", "PUS"],                       # 日本+韓國
        "東南亞":        ["BKK", "CNX", "HAN", "SGN", "DAD", "SIN", "KUL", "MNL",
                         "CEB", "CRK", "CGK", "DPS", "KTI"],                       # 含柬埔寨金邊(KTI)
        "歐洲/澳洲":     ["AMS", "CDG", "LHR", "MUC", "MXP", "VIE", "BNE"],         # 歐洲+澳洲(布里斯本)
        "北美":          ["LAX", "SFO", "SEA", "YVR", "ORD", "IAH", "JFK", "DFW", "YYZ"],
        "中國/香港/澳門": ["CAN", "PEK", "PVG", "SHA", "SZX", "XMN", "HGH", "TFU",   # 中國大陸
                         "HKG", "MFM"],                                            # 香港、澳門
    }
    CITY_TO_REGION = {c: reg for reg, cs in _REGION_CITIES.items() for c in cs}

    # 艙等(代碼 -> 友善標籤)
    CABIN_MAP = {"C": "商務艙(C)", "K": "豪華經濟艙(K)", "Y": "經濟艙(Y)"}

    df1["構面"] = df1["題號"].map(QNUM_TO_NAME)
    df1["艙等"] = df1["艙等"].map(CABIN_MAP).fillna(df1["艙等"])

    # 題目英文 -> 題號 -> 構面(供非滿意度表對應)
    text_to_qnum = df1.drop_duplicates("題號").set_index("題目")["題號"].to_dict()
    df2["構面"] = df2["關聯滿意度題目"].map(text_to_qnum).map(QNUM_TO_NAME)
    # 對應不到 5 大構面之原因(其關聯題目非保留題號，如 Wi-Fi/餐飲等機上題項)構面會是空值；
    # 填為「其他」使其在「全部」檢視仍可計數(選特定構面時自動排除)，避免待改善原因 TOP 提及次數歸零。
    _n_face_na = int(df2["構面"].isna().sum())
    df2["構面"] = df2["構面"].fillna("其他").replace(r"^\s*$", "其他", regex=True)
    if _n_face_na:
        print(f"　ℹ 原因明細：{_n_face_na:,} 列關聯題目非 5 大構面，構面歸為「其他」(於『全部』檢視仍計入待改善原因次數)")
    df2["艙等"] = df2["艙等"].map(CABIN_MAP).fillna(df2["艙等"])

    # ---- 年度/月份正規化：避免被讀成浮點(2025.0)後 astype(int) 崩潰 ----
    def _to_int_str(series):
        """將年度/月份等欄位正規化為乾淨整數字串：2025.0 / '2025.0' / 2025 → '2025'。
        無法轉為數值者保留原字串(僅去除結尾 .0)。"""
        s = series.astype(str).str.replace(r"\.0+$", "", regex=True)
        num = pd.to_numeric(series, errors="coerce")
        mask = num.notna()
        if mask.any():
            s = s.copy()
            s.loc[mask] = num[mask].round().astype("int64").astype(str)
        return s

    def _numkey(x):
        """排序鍵：能轉數值就以數值排序，否則以字串排序，永不崩潰。"""
        try:
            return (0, float(x))
        except Exception:
            return (1, str(x))

    # 整理後的明細資料表
    # ---- 年度別統計規則：2025 年僅計入「填答來源為 EDM 連結」之旅客；2026(含)以後維持全部 ----
    #   各年度資料命名可能有差異(2026=「EDM」、2025=「EMD連結」等)，故同時比對多個關鍵字。
    EDM_ONLY_YEARS = {y for y in re.split(r"[,\s、，;；]+", (os.environ.get("EVA_EDM_ONLY_YEARS") or "2025").strip()) if y}
    EDM_SOURCE_KWS = [k for k in re.split(r"[,\s、，;；/|]+", (os.environ.get("EVA_EDM_SOURCE") or "EDM EMD").strip()) if k]
    if not EDM_SOURCE_KWS:
        EDM_SOURCE_KWS = ["EDM", "EMD"]
    _EDM_PAT = "|".join(re.escape(k) for k in EDM_SOURCE_KWS)
    _EDM_LABEL = "/".join(EDM_SOURCE_KWS)

    def _apply_year_source_rule(df, label=""):
        """指定年度(EDM_ONLY_YEARS)僅保留『填答來源含任一 EDM 關鍵字(EDM/EMD…)』之列；其餘年度全部保留。"""
        if df is None or "搭機年度" not in df.columns or "填答來源" not in df.columns or not EDM_ONLY_YEARS:
            return df
        yr = df["搭機年度"].astype(str).str.strip()
        src = df["填答來源"].astype(str)
        tgt = yr.isin(EDM_ONLY_YEARS)
        is_edm = src.str.contains(_EDM_PAT, case=False, na=False, regex=True)
        # 診斷：列出各目標年度的『填答來源』分布與套規則後保留列數(歸零時警告)
        if tgt.any():
            for y in sorted(set(yr[tgt])):
                m = tgt & (yr == y)
                dist = src[m].value_counts().to_dict()
                kept = int((m & is_edm).sum())
                print(f"　🎯 {y} 年（{label}）EDM 規則：原始來源分布 {dist}；含「{_EDM_LABEL}」保留 {kept} 列")
                if kept == 0 and int(m.sum()) > 0:
                    print(f"　⚠ {y} 年（{label}）套用 EDM 規則後為 0 列！請確認該年度『填答來源』實際字樣；"
                          f"若非 {_EDM_LABEL}，可設定環境變數 EVA_EDM_SOURCE=實際字樣(多個以空白/逗號分隔)。")
        keep = (~tgt) | (tgt & is_edm)
        _drop = int((~keep).sum())
        if _drop:
            print(f"　🎯 年度統計規則：{'/'.join(sorted(EDM_ONLY_YEARS))} 年僅計入填答來源含「{_EDM_LABEL}」者，"
                  f"{label}排除 {_drop:,} 列")
        return df[keep].copy()

    sat = df1[["問卷編號", "會員卡別", "國籍", "搭機年度", "搭機月份", "班機編號",
               "機型", "艙等", "起飛城市", "填答來源", "性別", "構面", "分數"]].copy()
    sat["搭機年度"] = _to_int_str(sat["搭機年度"])
    sat["搭機月份"] = _to_int_str(sat["搭機月份"])
    sat = _apply_year_source_rule(sat, "滿意度")          # 2025 僅計入 EDM 連結填答
    try:
        _yc = sat["搭機年度"].astype(str).value_counts().to_dict()
        print(f"　📅 各年度滿意度列數：{ {k: int(v) for k, v in sorted(_yc.items())} }")
    except Exception:
        pass
    # 依起飛城市歸併「區域」(置於最末欄，不影響既有欄序)；未列入對照者歸「其他」
    sat["區域"] = sat["起飛城市"].astype(str).map(CITY_TO_REGION).fillna("其他")
    _unmapped = sorted(set(sat.loc[sat["區域"] == "其他", "起飛城市"].astype(str).unique()))
    if _unmapped:
        print(f"※ 注意：下列起飛城市尚未列入區域對照，暫歸『其他』，請於 _REGION_CITIES 補上：{_unmapped}")

    # ---- 由資料自動判定「資料期間」範圍(可跨多年月)，供檔名與標題動態使用 ----
    _EN_MONTHS = {"1":"JAN","2":"FEB","3":"MAR","4":"APR","5":"MAY","6":"JUN",
                  "7":"JUL","8":"AUG","9":"SEP","10":"OCT","11":"NOV","12":"DEC"}
    _per = sat[["搭機年度", "搭機月份"]].drop_duplicates().copy()
    _per["_y"] = pd.to_numeric(_per["搭機年度"], errors="coerce").fillna(0).astype(int)
    _per["_m"] = pd.to_numeric(_per["搭機月份"], errors="coerce").fillna(0).astype(int)
    _per = _per.sort_values(["_y", "_m"])
    _first, _last = _per.iloc[0], _per.iloc[-1]
    N_PERIODS = len(_per)
    if N_PERIODS == 1:
        PERIOD_TW       = f"{_last['搭機年度']} 年 {_last['搭機月份']} 月"
        PERIOD_TW_TIGHT = f"{_last['搭機年度']}年{_last['搭機月份']}月"
        PERIOD_EN       = f"{_EN_MONTHS.get(str(_last['_m']), _last['搭機月份'])} {_last['搭機年度']}"
    else:
        PERIOD_TW       = (f"{_first['搭機年度']} 年 {_first['搭機月份']} 月 ～ "
                           f"{_last['搭機年度']} 年 {_last['搭機月份']} 月（共 {N_PERIODS} 個月別）")
        PERIOD_TW_TIGHT = f"{_first['搭機年度']}年{_first['搭機月份']}月～{_last['搭機年度']}年{_last['搭機月份']}月"
        PERIOD_EN       = (f"{_EN_MONTHS.get(str(_first['_m']), _first['搭機月份'])} {_first['搭機年度']} – "
                           f"{_EN_MONTHS.get(str(_last['_m']), _last['搭機月份'])} {_last['搭機年度']}")

    OUTPUT_NAME = "EVA_旅客滿意度互動分析_全期間.xlsx" if N_PERIODS > 1 else f"EVA_旅客滿意度互動分析_{PERIOD_TW_TIGHT}.xlsx"
    # 匯出路徑優先序：環境變數 EVA_OUTPUT_FILE(GUI 用) > 命令列第 2 參數 > 預設(資料庫根目錄 / 自動檔名)
    OUTPUT_FILE = (os.environ.get("EVA_OUTPUT_FILE")
                   or (sys.argv[2] if len(sys.argv) > 2 else os.path.join(OUTPUT_DIR, OUTPUT_NAME)))

    rsn = df2[["問卷編號", "會員卡別", "國籍", "班機編號", "機型", "艙等", "起飛城市",
               "填答來源", "性別", "構面", "關聯滿意度分數", "選項", "細項1選項",
               "搭機年度", "搭機月份"]].copy()
    rsn.columns = ["問卷編號", "會員卡別", "國籍", "班機編號", "機型", "艙等", "起飛城市",
                   "填答來源", "性別", "構面", "關聯分數", "原因選項", "細項", "搭機年度", "搭機月份"]
    rsn["搭機年度"] = _to_int_str(rsn["搭機年度"])
    rsn["搭機月份"] = _to_int_str(rsn["搭機月份"])
    rsn["區域"] = rsn["起飛城市"].astype(str).map(CITY_TO_REGION).fillna("其他")  # 末欄
    # 「待改善原因」欄：優先採用較具體之「細項」，細項為空白/無文字時改用「原因選項」；
    # 供原因清單(REASON_LIST)與各原因次數分析(TOP10／原因解析／問卷項目總覽最主要原因)統一使用。
    _xi_empty = rsn["細項"].isna() | rsn["細項"].astype(str).str.strip().isin(["", "nan", "None", "NaN", "-"])
    rsn["待改善原因"] = rsn["原因選項"].where(_xi_empty, rsn["細項"]).astype(str).str.strip()
    rsn = _apply_year_source_rule(rsn, "原因")            # 2025 僅計入 EDM 連結填答

    # 所有「下拉篩選用」維度欄一律轉文字，確保 COUNTIFS 文字比對與萬用字元(*)正常運作
    _SAT_TEXT = ["會員卡別", "國籍", "搭機年度", "班機編號", "機型", "艙等",
                 "起飛城市", "填答來源", "性別", "構面", "區域"]
    _RSN_TEXT = ["會員卡別", "國籍", "班機編號", "機型", "艙等", "起飛城市",
                 "填答來源", "性別", "構面", "原因選項", "搭機年度", "區域"]
    for c in _SAT_TEXT:
        sat[c] = sat[c].astype(str)
    for c in _RSN_TEXT:
        rsn[c] = rsn[c].astype(str)

    # 搭機月份改存為「數值」(1..12)：月份篩選一律走「介於 起月～迄月」之數值範圍(單月即 [n,n])，
    # 數值比較才正確(文字配 ">=" 會被當數值致比對落空)。期間顯示已於此前算好，不受影響。
    for _df in (sat, rsn):
        _df["搭機月份"] = pd.to_numeric(_df["搭機月份"], errors="coerce").fillna(0).astype(int)

    # 分數一律轉為「數值」：部分年度來源檔可能把分數存成文字(如 2025)，
    # 而 Excel 的 SUMIFS/AVERAGEIFS 只加總數值、會略過文字 → 導致平均顯示 0。
    # 故在寫入明細前強制數值化，確保各年度分數統計一致正確。
    sat["分數"] = pd.to_numeric(sat["分數"], errors="coerce")
    if "關聯分數" in rsn.columns:
        rsn["關聯分數"] = pd.to_numeric(rsn["關聯分數"], errors="coerce")
    _bad = int(sat["分數"].isna().sum())
    if _bad:
        print(f"　⚠ 滿意度分數有 {_bad:,} 列非數值，已視為空值不列入加總(請檢查來源檔分數欄)。")

    # 去除重複明細列：每份問卷(問卷編號)在同一(年/月/城市)對同一構面只計一次，
    # 使「回覆數量＝不重複問卷份數(份)」。可移除重複匯出/多列(多航段等)造成之灌水；
    # 因題號→構面為 1:1，同問卷同構面分數相同，移除重複列後平均不變。
    _bs, _br = len(sat), len(rsn)
    _SAT_DEDUP_KEYS = ["問卷編號", "構面", "搭機年度", "搭機月份", "起飛城市"]
    sat = sat.drop_duplicates(subset=_SAT_DEDUP_KEYS).reset_index(drop=True)
    rsn = rsn.drop_duplicates().reset_index(drop=True)
    if (_bs - len(sat)) or (_br - len(rsn)):
        print(f"　已忽略重複明細：滿意度 {_bs - len(sat)} 列(以問卷編號×構面×年×月×城市為準)、原因 {_br - len(rsn)} 列")

    n_sat, n_rsn = len(sat), len(rsn)
    # 公式範圍上界固定為 Excel 最大列(1048576)，使各分頁公式「與資料列數無關」：
    # 日後只更新明細資料(不重建儀表板)時，列數增減公式仍正確涵蓋。Excel 對 COUNTIFS/
    # AVERAGEIFS/SUMIFS 之整欄範圍會自動只算到「已使用範圍」，效能不受影響。
    FORMULA_MAX_ROW = 1048576
    sat_lo, sat_hi = 2, FORMULA_MAX_ROW
    rsn_lo, rsn_hi = 2, FORMULA_MAX_ROW
    print(f"滿意度明細 {n_sat} 列；原因明細 {n_rsn} 列")

    # 各維度選項清單(含「全部」)
    def opts(series, manual_order=None):
        vals = manual_order if manual_order else sorted(series.dropna().astype(str).unique())
        return ["全部"] + list(vals)

    def build_lists(sdf, rdf):
        """由明細資料(sdf/rdf)推導各維度下拉清單與原因清單；可用於全新或累積後資料。"""
        L = {
            "構面": opts(sdf["構面"], DIM_ORDER),
            "會員卡別": opts(sdf["會員卡別"]),
            "艙等": opts(sdf["艙等"], [CABIN_MAP[k] for k in ["C", "K", "Y"] if CABIN_MAP[k] in set(sdf["艙等"])]),
            "機型": opts(sdf["機型"]),
            "起飛城市": opts(sdf["起飛城市"]),
            "區域": opts(sdf["區域"], [r for r in REGION_ORDER if r in set(sdf["區域"])]
                         + (["其他"] if (sdf["區域"] == "其他").any() else [])),
            "國籍": opts(sdf["國籍"]),
            "班機編號": opts(sdf["班機編號"]),
            "填答來源": opts(sdf["填答來源"]),
            "性別": opts(sdf["性別"]),
            "搭機年度": opts(sdf["搭機年度"], sorted(sdf["搭機年度"].dropna().unique(), key=_numkey)),
            "搭機月份": opts(sdf["搭機月份"], sorted(sdf["搭機月份"].dropna().unique(), key=_numkey)),
        }
        _rcol = "待改善原因" if "待改善原因" in rdf.columns else "原因選項"
        R = sorted(x for x in rdf[_rcol].dropna().astype(str).str.strip().unique() if x not in ("", "nan", "None"))
        return L, R

    LISTS, REASON_LIST = build_lists(sat, rsn)

    # ============================ 2. 樣式定義 (EVA 品牌) ============================
    FN = "Microsoft JhengHei"
    EVA_GREEN   = "004D33"   # 主色 深綠
    EVA_MID     = "00875A"   # 次級綠
    EVA_LIGHT   = "E3F0EB"   # 淺綠底
    EVA_LIGHT2  = "F2F8F5"   # 更淺綠底(隔行)
    EVA_GOLD    = "B7903B"   # 金
    EVA_ORANGE  = "D9534F"   # 警示/低分(沉穩紅)
    WHITE, BLACK, GREY, LGREY = "FFFFFF", "1A1A1A", "595959", "D9D9D9"
    INPUT_FILL  = "FFF7E0"   # 可輸入欄(淡黃)

    def F(sz=11, b=False, color=BLACK, italic=False):
        return Font(name=FN, size=sz, bold=b, color=color, italic=italic)
    def fill(c):  return PatternFill("solid", fgColor=c)
    def side(c=LGREY, st="thin"): return Side(style=st, color=c)
    def box(c=LGREY, st="thin"):
        s = side(c, st); return Border(left=s, right=s, top=s, bottom=s)
    C = Alignment(horizontal="center", vertical="center", wrap_text=True)
    L = Alignment(horizontal="left", vertical="center", wrap_text=True)
    Rr = Alignment(horizontal="right", vertical="center")
    thin_box = box(); mid_box = box(EVA_MID)

    # ---- 共用「依目標值雙向漸層」著色：達標(≥目標)綠、越高越深綠；未達標(<目標)紅、越低越深紅(無黃色) ----
    _GRAD_GREEN = [(0.20, "1B7E3F", "FFFFFF"), (0.10, "4CAE6A", "06402B"),
                   (0.03, "86CF9C", "06402B"), (0.00, "CDEBD6", "06402B")]
    _GRAD_RED   = [(-0.03, "F7D6D4", "7A1B17"), (-0.10, "ED9C97", "7A1B17"),
                   (-0.20, "DE6A63", "FFFFFF"), (None,  "B5261E", "FFFFFF")]
    def _grad_fill_rule(cond, bg, fg):
        # 條件式格式須以 start/end color 設定實心「儲存格底色」(僅 fgColor 在部分 Excel 不顯示)
        return FormulaRule(formula=[cond],
                           fill=PatternFill(start_color=bg, end_color=bg, fill_type="solid"),
                           font=Font(name=FN, color=fg, bold=True), stopIfTrue=True)
    def add_target_gradient(ws, cell_range, anchor, target, target_is_cell=False):
        """對 cell_range 依「分數−目標」雙向著色。
        anchor=範圍左上格(相對基準)；target=固定數值字串(如 '4.49') 或 目標儲存格參照。"""
        tgt = target if target_is_cell else f"{float(target):.4f}"
        guard = f"ISNUMBER({anchor})" + (f',{tgt}<>""' if target_is_cell else "")
        for thr, bg, fg in _GRAD_GREEN:
            ws.conditional_formatting.add(cell_range, _grad_fill_rule(
                f"AND({guard},{anchor}-{tgt}>={thr})", bg, fg))
        for thr, bg, fg in _GRAD_RED:
            cond = (f"AND({guard},{anchor}-{tgt}>={thr})" if thr is not None
                    else f"AND({guard},{anchor}<{tgt})")
            ws.conditional_formatting.add(cell_range, _grad_fill_rule(cond, bg, fg))

    # ============================ 3. 建立工作簿 ============================
    # ---- 明細資料表寫入器(來源資料，供公式引用)；可用於全新建立或更新模式 ----
    def write_data_sheet(wb, name, dataframe, tab=EVA_MID):
        ws = wb.create_sheet(name)
        ws.sheet_properties.tabColor = tab
        cols = list(dataframe.columns)
        ws.append(cols)
        for cell in ws[1]:
            cell.font = F(10, True, WHITE); cell.fill = fill(EVA_GREEN)
            cell.alignment = C; cell.border = thin_box
        for row in dataframe.itertuples(index=False):
            ws.append(list(row))
        # 寬度
        widths = {"問卷編號": 11, "會員卡別": 9, "國籍": 8, "搭機年度": 9, "搭機月份": 9,
                  "班機編號": 11, "機型": 7, "艙等": 13, "起飛城市": 9, "填答來源": 10,
                  "性別": 8, "構面": 15, "分數": 7, "關聯分數": 9, "原因選項": 34, "細項": 30,
                  "待改善原因": 36}
        for i, c in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 12)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(dataframe)+1}"
        ws.sheet_view.showGridLines = False
        return ws

    # ---- 模式判斷 ----
    # 成品不存在、或加 --rebuild / EVA_REBUILD=1、或 UPDATE_MODE="FULL" → 「全新建立」(重建所有分頁)。
    # 成品已存在、未指定 --rebuild，且 UPDATE_MODE="DATA_ONLY" → 「資料更新模式」(只換明細+清單下拉，保留人工維護)。
    _DATA_ONLY = (not _REBUILD) and (UPDATE_MODE == "DATA_ONLY") and os.path.exists(OUTPUT_FILE)
    if _DATA_ONLY:
        from openpyxl import load_workbook
        wb = load_workbook(OUTPUT_FILE)

        # ---- 增量更新：以現有明細為準，只把「新問卷編號」的資料往下新增；重複者忽略 ----
        def _read_detail_df(sheet, columns):
            rows = list(sheet.iter_rows(min_row=2, values_only=True))
            import pandas as _pd
            df = _pd.DataFrame(rows, columns=[c.value for c in sheet[1]]) if rows else _pd.DataFrame(columns=columns)
            return df

        ws_sat = wb["滿意度明細"]; ws_rsn = wb["原因明細"]
        _ex_sat = _read_detail_df(ws_sat, list(sat.columns))
        _ex_rsn = _read_detail_df(ws_rsn, list(rsn.columns))
        _ex_sat_ids = set(_ex_sat["問卷編號"].astype(str)) if len(_ex_sat) else set()
        _ex_rsn_ids = set(_ex_rsn["問卷編號"].astype(str)) if len(_ex_rsn) else set()
        # 新資料 = 問卷編號尚未存在於現有明細者(整份問卷視為一筆，已存在即略過)
        _new_sat = sat[~sat["問卷編號"].astype(str).isin(_ex_sat_ids)].copy()
        _new_rsn = rsn[~rsn["問卷編號"].astype(str).isin(_ex_rsn_ids)].copy()
        _skip_q = sat["問卷編號"].astype(str).isin(_ex_sat_ids).sum()
        # 往下新增(不重建、不更動既有列)
        for _row in _new_sat.itertuples(index=False):
            ws_sat.append(list(_row))
        for _row in _new_rsn.itertuples(index=False):
            ws_rsn.append(list(_row))
        ws_sat.auto_filter.ref = f"A1:{get_column_letter(len(sat.columns))}{ws_sat.max_row}"
        ws_rsn.auto_filter.ref = f"A1:{get_column_letter(len(rsn.columns))}{ws_rsn.max_row}"
        _added_sat, _added_rsn = len(_new_sat), len(_new_rsn)

        # 累積後資料(現有 + 新增)→ 重新推導清單，使下拉含全部累積期間之分類
        import pandas as _pd
        _comb_sat = _pd.concat([_ex_sat, _new_sat], ignore_index=True) if len(_ex_sat) else _new_sat
        _comb_rsn = _pd.concat([_ex_rsn, _new_rsn], ignore_index=True) if len(_ex_rsn) else _new_rsn
        for _df_c in (_comb_sat, _comb_rsn):
            for _c in ("搭機年度", "搭機月份"):
                if _c in _df_c.columns:
                    _df_c[_c] = _to_int_str(_df_c[_c])
        LISTS, REASON_LIST = build_lists(_comb_sat, _comb_rsn)

        # ---- 同步刷新「清單」分頁類別欄 + 各分頁下拉選單範圍(依累積後資料) ----
        # 僅改寫類別值欄(A~L、R)與下拉來源範圍；完全不動條件轉換欄(AB~AF)與
        # 最主要原因格線公式(U~Y)，分析分頁之人工維護亦不受影響。
        _list_cols = list(LISTS.keys())
        _cat_changes = _dv_changes = 0
        _CLR_HI = 1200   # 類別欄清除上限(預留班機編號等長清單之成長空間)
        if "清單" in wb.sheetnames:
            wl = wb["清單"]
            # (1) 重寫類別欄 A..L(含「全部」)；先清殘列再寫新值，標題列(第1列)保留
            for j, key in enumerate(_list_cols, 1):
                col = get_column_letter(j)
                for r in range(2, _CLR_HI + 1):
                    wl[f"{col}{r}"] = None
                for i, v in enumerate(LISTS[key], 2):
                    wl[f"{col}{i}"] = v
                _cat_changes += 1
            # (2) 原因清單(R 欄)
            for r in range(2, _CLR_HI + 1):
                wl[f"R{r}"] = None
            for i, v in enumerate(REASON_LIST, 2):
                wl[f"R{i}"] = v
            # (3) 原因格線來源(T 欄)：項目數不變→就地更新；有變動→保留並提示改用 --rebuild
            _oldT = sum(1 for r in range(2, _CLR_HI + 1) if wl[f"T{r}"].value not in (None, ""))
            if _oldT == len(REASON_LIST):
                for i, v in enumerate(REASON_LIST, 2):
                    wl[f"T{i}"] = v
            else:
                print("　⚠ 原因選項項目數有異動，『問卷項目總覽』最主要原因格線未自動延展；")
                print("　　如需同步，請執行：python build_dashboard.py --rebuild")
            # (4) 重新指向所有分頁下拉(資料驗證)範圍 → 對應新清單長度
            _new_end = {get_column_letter(j): len(LISTS[key]) + 1
                        for j, key in enumerate(_list_cols, 1)}
            _new_end["R"] = len(REASON_LIST) + 1
            for ws_any in wb.worksheets:
                dvs = getattr(ws_any, "data_validations", None)
                if not dvs:
                    continue
                for dv in dvs.dataValidation:
                    if dv.type != "list" or not dv.formula1 or "清單" not in dv.formula1:
                        continue
                    try:
                        rhs = dv.formula1.split("!", 1)[1]            # $E$2:$E$58
                        first = rhs.split(":")[0]                     # $E$2
                        letters = "".join(c for c in first.lstrip("$") if c.isalpha())
                    except Exception:
                        continue
                    if letters in _new_end:
                        dv.formula1 = f"'清單'!${letters}$2:${letters}${_new_end[letters]}"
                        _dv_changes += 1

        wb.save(OUTPUT_FILE)
        print(f"【更新模式：DATA_ONLY 增量更新】新增明細：滿意度 {_added_sat} 列、原因 {_added_rsn} 列"
              f"(略過已存在問卷 {_skip_q} 列)。")
        print(f"　以現有明細為準往下新增，重複問卷自動忽略；累積明細現為 滿意度 {ws_sat.max_row - 1} 列、"
              f"原因 {ws_rsn.max_row - 1} 列。")
        print(f"　已依累積資料刷新『清單』類別欄({_cat_changes} 欄)與下拉範圍({_dv_changes} 處)；"
              f"分析分頁與人工維護維持原樣。")
        print(f"　註：依區域連動之下拉(城市/班機/機型…)其選項清單於『重建(FULL)』時更新；如新增了"
              f"區域別的新分類，請以 UPDATE_MODE=\"FULL\" 或 --rebuild 重建一次。")
        print(f"已儲存：{OUTPUT_FILE}")
        sys.exit(0)

    # 全新建立模式(首次建檔 / UPDATE_MODE=FULL 重建 / --rebuild)
    if os.path.exists(OUTPUT_FILE):
        print(f"【重建模式：{'--rebuild' if _REBUILD else 'FULL'}】成品已存在，將以最新資料重建所有頁籤"
              f"(目標值等取程式內建設定，不保留 Excel 內人工編輯)。")
    else:
        print(f"【首次建檔】將建立全部頁籤。")
    wb = Workbook()

    ws_sat = write_data_sheet(wb, "滿意度明細", sat)
    ws_rsn = write_data_sheet(wb, "原因明細", rsn)

    # ---- 3b. 清單表(下拉式選單來源；隱藏) ----
    ws_list = wb.create_sheet("清單")
    ws_list.sheet_state = "hidden"
    wsT = wb.create_sheet("目標參照")      # 目標著色輔助頁(獨立，避免與「清單」下拉來源欄衝突)
    wsT.sheet_state = "hidden"
    TGT = "目標參照"
    list_cols = list(LISTS.keys())
    for j, key in enumerate(list_cols, 1):
        col = get_column_letter(j)
        ws_list[f"{col}1"] = key
        ws_list[f"{col}1"].font = F(10, True, WHITE); ws_list[f"{col}1"].fill = fill(EVA_GREEN)
        for i, v in enumerate(LISTS[key], 2):
            ws_list[f"{col}{i}"] = v
    # 原因清單放在 R 欄
    ws_list["R1"] = "原因選項清單"; ws_list["R1"].font = F(10, True, WHITE); ws_list["R1"].fill = fill(EVA_GREEN)
    for i, v in enumerate(REASON_LIST, 2):
        ws_list[f"R{i}"] = v
    LIST_RANGE = {key: f"'清單'!${get_column_letter(j)}$2:${get_column_letter(j)}${len(LISTS[key])+1}"
                  for j, key in enumerate(list_cols, 1)}

    # ============================ 4. 公式產生器 + 各分頁篩選基礎建設 ============================
    SATq = "'滿意度明細'"; RSNq = "'原因明細'"
    def sat_col(c): return f"{SATq}!${c}${sat_lo}:${c}${sat_hi}"
    def rsn_col(c): return f"{RSNq}!${c}${rsn_lo}:${c}${rsn_hi}"

    # 維度 -> 明細欄位字母(滿意度明細 / 原因明細)；區域為新增之末欄(滿意度明細 N、原因明細 P)
    SAT_COL = {"構面":"L","會員卡別":"B","國籍":"C","搭機年度":"D","搭機月份":"E",
               "班機編號":"F","機型":"G","艙等":"H","起飛城市":"I","填答來源":"J","性別":"K","區域":"N"}
    RSN_COL = {"構面":"J","會員卡別":"B","國籍":"C","班機編號":"D","機型":"E","艙等":"F",
               "起飛城市":"G","填答來源":"H","性別":"I","搭機年度":"N","搭機月份":"O","區域":"P"}
    # 各分頁篩選面板專用：在「清單」表以一欄存放「條件轉換」儲存格(全部→*)，列號固定如下
    # 篩選條件主排序(所有頁籤一致)：構面→年度→月份→區域→城市→班機→機型→艙等→卡別→國籍→性別→來源
    MASTER_ORDER = ["構面", "搭機年度", "搭機月份", "區域", "起飛城市", "班機編號",
                    "機型", "艙等", "會員卡別", "國籍", "性別", "填答來源"]
    # 「區域以下」的維度 → 依所選區域連動(靜態具名清單 + INDIRECT)
    DEP_DIMS = ["起飛城市", "班機編號", "機型", "艙等", "會員卡別", "國籍", "性別", "填答來源"]
    DIMKEY = {"起飛城市": "城", "班機編號": "機", "機型": "型", "艙等": "艙",
              "會員卡別": "卡", "國籍": "國", "性別": "性", "填答來源": "源"}   # 具名清單前綴
    def _ord(dims):  # 依主排序排列維度子集
        return [d for d in MASTER_ORDER if d in dims]

    ALL_DIMS = list(MASTER_ORDER)
    DIM_CRIT_ROW = {d: i + 2 for i, d in enumerate(ALL_DIMS)}       # 構面=2 … 來源=13
    DIM_LABEL = {"構面": "滿意度構面"}                                 # 顯示用標籤(其餘同名)
    # 各分頁的條件轉換欄(互動查詢沿用既有 AB；其餘分頁各自一欄)
    CRIT_OV, CRIT_CR, CRIT_CT, CRIT_RS = "AC", "AD", "AE", "AF"
    # 各分頁的篩選維度集合(均含「區域」)，一律依主排序呈現
    OV_DIMS = _ord(["會員卡別","艙等","機型","區域","起飛城市","國籍","班機編號","填答來源","性別","搭機年度","搭機月份"])
    CR_DIMS = _ord(["區域","起飛城市","國籍","班機編號","填答來源","性別","搭機年度","搭機月份"])
    CT_DIMS = _ord(["會員卡別","艙等","機型","區域","國籍","班機編號","填答來源","性別","搭機年度","搭機月份"])
    RS_DIMS = _ord(["構面","會員卡別","艙等","機型","區域","起飛城市","國籍","班機編號","填答來源","性別","搭機年度","搭機月份"])

    # ====== 搭機月份「季度/半年/單月」範圍篩選 ======================================
    # Excel 下拉無法真正複選；改以「預設期間」選項，將月份篩選一律視為「介於 起月～迄月」之範圍。
    # 月份於明細已存為兩位數文字("01".."12")，故可正確比大小。選項→(起,迄)以查表取得。
    _months_actual = [m for m in LISTS["搭機月份"] if m != "全部"]          # [4,5,…](數值，僅資料中出現者)
    MONTH_PRESETS = [("全部", 1, 12), ("全年", 1, 12),
                     ("Q1（1-3月）", 1, 3), ("Q2（4-6月）", 4, 6),
                     ("Q3（7-9月）", 7, 9), ("Q4（10-12月）", 10, 12),
                     ("上半年（1-6月）", 1, 6), ("下半年（7-12月）", 7, 12)]
    MONTH_DD = [p[0] for p in MONTH_PRESETS] + _months_actual               # 下拉顯示清單(含季度+單月)
    ws_list["M1"] = "搭機月份(含季度)"; ws_list["M1"].font = F(9, True, WHITE); ws_list["M1"].fill = fill(EVA_GREEN)
    for i, v in enumerate(MONTH_DD, 2):
        ws_list[f"M{i}"] = v
    MONTH_DV_RANGE = f"'清單'!$M$2:$M${len(MONTH_DD) + 1}"
    # 選項→起訖月對照表(N=選項, O=起月, P=迄月，皆數值)；單月對應自身
    _mtbl = MONTH_PRESETS + [(m, m, m) for m in _months_actual]
    ws_list["N1"] = "月份選項"; ws_list["O1"] = "起月"; ws_list["P1"] = "迄月"
    for col in ("N", "O", "P"):
        ws_list[f"{col}1"].font = F(9, True, WHITE); ws_list[f"{col}1"].fill = fill(EVA_GREEN)
    for i, (lbl, lo, hi) in enumerate(_mtbl, 2):
        ws_list[f"N{i}"] = lbl; ws_list[f"O{i}"] = lo; ws_list[f"P{i}"] = hi
    MONTH_TBL_RANGE = f"清單!$N$2:$P${len(_mtbl) + 1}"
    # 各頁(含互動查詢)之「起訖月」查找格：Q=起月, Z=迄月(以該頁月份輸入格查表)
    _MONTH_LH = {}                          # 索引(各頁 crit_col / "互動查詢") -> (起月參照, 迄月參照)
    _mlh_row = [1]
    def _alloc_month_lh(month_input_ref):
        _mlh_row[0] += 1; r = _mlh_row[0]
        ws_list[f"Q{r}"] = f"=IFERROR(VLOOKUP({month_input_ref},{MONTH_TBL_RANGE},2,FALSE),1)"
        ws_list[f"Z{r}"] = f"=IFERROR(VLOOKUP({month_input_ref},{MONTH_TBL_RANGE},3,FALSE),12)"
        return f"清單!$Q${r}", f"清單!$Z${r}"
    # 預先為各分頁面板與互動查詢配置起訖月查找格(於任何配對公式產生前)
    _PANEL_MONTH = [(CRIT_OV, "問卷項目總覽", OV_DIMS, 10, 4),
                    (CRIT_CR, "交叉分析", CR_DIMS, 9, 4),
                    (CRIT_CT, "城市排行", CT_DIMS, 11, 4),
                    (CRIT_RS, "原因解析", RS_DIMS, 7, 4)]
    for _cc, _sn, _dims, _lcol, _trow in _PANEL_MONTH:
        if "搭機月份" in _dims:
            _mref = f"'{_sn}'!${get_column_letter(_lcol + 1)}${_trow + _dims.index('搭機月份')}"
            _MONTH_LH[_cc] = _alloc_month_lh(_mref)
    _MONTH_LH["互動查詢"] = _alloc_month_lh("互動查詢!$D$8")   # 互動查詢月份篩選格 D8

    def _month_range_pair(col_ref, lo_hi):
        lo, hi = lo_hi
        return f'{col_ref},">="&{lo},{col_ref},"<="&{hi}'

    def sat_pairs(dims, crit_col):   # 給 COUNTIFS/AVERAGEIFS 的「滿意度明細範圍,條件」配對
        out = []
        for d in dims:
            if d == "搭機月份" and crit_col in _MONTH_LH:
                out.append(_month_range_pair(sat_col(SAT_COL[d]), _MONTH_LH[crit_col]))
            else:
                out.append(f"{sat_col(SAT_COL[d])},清單!${crit_col}${DIM_CRIT_ROW[d]}")
        return ",".join(out)
    def rsn_pairs(dims, crit_col):   # 同上，對應原因明細
        out = []
        for d in dims:
            if d == "搭機月份" and crit_col in _MONTH_LH:
                out.append(_month_range_pair(rsn_col(RSN_COL[d]), _MONTH_LH[crit_col]))
            else:
                out.append(f"{rsn_col(RSN_COL[d])},清單!${crit_col}${DIM_CRIT_ROW[d]}")
        return ",".join(out)

    def build_filter_panel_v(ws, sheet_name, dims, crit_col, label_col, top_row=4, title_row=2):
        """於分頁右側建立垂直篩選面板：含標題、各維度下拉輸入格(資料驗證)，
           並在『清單』表 {crit_col} 欄建立『全部→*』條件轉換儲存格，供 COUNTIFS/AVERAGEIFS 引用。"""
        Lc, Ic = label_col, label_col + 1
        Lcl, Icl = get_column_letter(Lc), get_column_letter(Ic)
        ws.merge_cells(start_row=title_row, start_column=Lc, end_row=title_row, end_column=Ic)
        t = ws.cell(row=title_row, column=Lc, value="▍篩選條件(本頁)")
        t.font = F(10.5, True, WHITE); t.fill = fill(EVA_MID); t.alignment = C
        ws.cell(row=title_row, column=Ic).fill = fill(EVA_MID)
        # 本頁「區域」「起飛城市」輸入格 → 供區域以下維度 / 城市→班機 連動(Req3)
        _panel_regidx = _panel_reg_ref = _panel_city_ref = None
        if "區域" in dims:
            _panel_reg_ref = f"'{sheet_name}'!${Icl}${top_row + dims.index('區域')}"
            _panel_regidx = _region_idx(_panel_reg_ref)
        if "起飛城市" in dims:
            _panel_city_ref = f"'{sheet_name}'!${Icl}${top_row + dims.index('起飛城市')}"
        for i, dim in enumerate(dims):
            rr = top_row + i
            lab = ws.cell(row=rr, column=Lc, value=DIM_LABEL.get(dim, dim))
            lab.font = F(9, True, EVA_GREEN); lab.fill = fill(EVA_LIGHT2); lab.alignment = L; lab.border = thin_box
            inp = ws.cell(row=rr, column=Ic, value="全部")
            inp.font = F(9, True, BLACK); inp.alignment = C; inp.fill = fill(INPUT_FILL); inp.border = box(EVA_GOLD, "medium")
            if dim == "班機編號" and _panel_regidx:
                _src = _flight_source(_panel_reg_ref, _panel_regidx, _panel_city_ref)
            elif dim in DEP_DIMS and _panel_regidx:
                _src = _dep_source(dim, _panel_regidx)
            elif dim == "搭機月份":
                _src = MONTH_DV_RANGE          # 含季度/半年/單月
            else:
                _src = LIST_RANGE[dim]
            dv = DataValidation(type="list", formula1=_src, allow_blank=False)
            dv.prompt = "請點選下拉選單"; dv.promptTitle = DIM_LABEL.get(dim, dim)
            ws.add_data_validation(dv); dv.add(inp)
            ws_list[f"{crit_col}{DIM_CRIT_ROW[dim]}"] = (
                f"=IF('{sheet_name}'!${Icl}${rr}=\"全部\",\"*\",'{sheet_name}'!${Icl}${rr}&\"\")")
        ws.column_dimensions[Lcl].width = 12
        ws.column_dimensions[Icl].width = 13
        hint = ws.cell(row=top_row + len(dims), column=Lc, value="↑ 黃底可下拉篩選")
        hint.font = F(8, False, GREY)
        return top_row + len(dims)   # 面板最後一列

    # 構面 × 原因 次數矩陣(隱藏，供「問卷項目總覽」找出各構面最主要原因；隨該頁篩選連動)
    RSN_R_LO, RSN_R_HI = 2, FORMULA_MAX_ROW
    G_REASON_LO = 2
    G_REASON_HI = 1 + len(REASON_LIST)        # T2:T{...}
    for i, v in enumerate(REASON_LIST, G_REASON_LO):
        ws_list[f"T{i}"] = v
    _OV_GRID_PAIRS = rsn_pairs(OV_DIMS, CRIT_OV)   # 問卷項目總覽篩選 -> 套用到原因次數矩陣
    GRID_COL = {}  # 構面 -> 欄位字母
    for d_idx, dim in enumerate(DIM_ORDER):
        gcol = get_column_letter(21 + d_idx)  # U,V,W,X,Y
        GRID_COL[dim] = gcol
        ws_list[f"{gcol}1"] = dim
        for i in range(G_REASON_LO, G_REASON_HI + 1):
            ws_list[f"{gcol}{i}"] = (
                f"=COUNTIFS('原因明細'!$J${RSN_R_LO}:$J${RSN_R_HI},{gcol}$1,"
                f"'原因明細'!$Q${RSN_R_LO}:$Q${RSN_R_HI},$T{i},{_OV_GRID_PAIRS})")
    GRID_REASON_RNG = f"清單!$T${G_REASON_LO}:$T${G_REASON_HI}"

    # ============================ 4b. 互動查詢專用篩選條件 ============================
    # 篩選格 D6..D17 依 MASTER_ORDER 排列；(資料欄, 篩選格)由維度→欄位對照自動推導。
    SAT_CONDS = [(SAT_COL[d], f"$D${6 + i}") for i, d in enumerate(MASTER_ORDER)]
    RSN_CONDS = [(RSN_COL[d], f"$D${6 + i}") for i, d in enumerate(MASTER_ORDER)]
    _DASH_ROW = {d: 6 + i for i, d in enumerate(MASTER_ORDER)}   # 維度→互動查詢篩選列
    # ---- 效能關鍵：以 COUNTIFS/AVERAGEIFS 取代逐列 SUMPRODUCT ----
    # 在隱藏的「清單」表 AB 欄放 11 個「條件轉換」儲存格：選「全部」→ "*"(萬用字元，比對所有)，
    # 否則 → 該下拉值。COUNTIFS/AVERAGEIFS 原生函式遠快於 SUMPRODUCT，且不需逐列輔助公式。
    FILTER_ROWS = [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]   # 篩選儲存格 D6..D17(D17=區域)
    DASH = "互動查詢"
    ws_list["AB1"] = "條件轉換(全部→*)"; ws_list["AB1"].font = F(9, True, WHITE); ws_list["AB1"].fill = fill(EVA_GREEN)
    for fr in FILTER_ROWS:
        # 一律以 &"" 強制轉文字：避免 Excel 將數字外觀之下拉值(機型321、年度2026、班機等)
        # 存成數值，導致 COUNTIFS 以數值條件比對文字明細欄位而配不到(篩選後無資料)。
        ws_list[f"AB{fr}"] = f'=IF({DASH}!$D${fr}="全部","*",{DASH}!$D${fr}&"")'
    CRIT = {fr: f"清單!$AB${fr}" for fr in FILTER_ROWS}   # 篩選列 -> 條件儲存格參照

    # ---- 合併構面支援(僅互動查詢)：構面條件拆成 A/B 兩格，公式以兩段相加實作 OR ----
    # 選「全部」→A="*",B=哨符；選單一構面→A=該值,B=哨符；選合併項→A=第1構面,B=第2構面。
    _MERGED_FACE = [(nm, cs) for nm, cs in YEAR_ITEMS if len(cs) > 1]   # [(合併名,[構面1,構面2]),...]
    _NONE = "##NONE##"                                                  # 對不到任何列的哨符
    _d6 = f"{DASH}!$D$6"
    _faceA = f'IF({_d6}="全部","*",'
    _faceB = ""
    for nm, cs in _MERGED_FACE:
        _faceA += f'IF({_d6}="{nm}","{cs[0]}",'
        _faceB += f'IF({_d6}="{nm}","{cs[1]}",'
    _faceA += _d6 + ")" * (1 + len(_MERGED_FACE))
    _faceB += f'"{_NONE}"' + ")" * len(_MERGED_FACE)
    ws_list["AB6"] = "=" + _faceA                     # 構面A(覆寫原通用條件)
    ws_list["AG6"] = "=" + _faceB                     # 構面B(第2構面或哨符)
    FACE_A, FACE_B = "清單!$AB$6", "清單!$AG$6"
    # 互動查詢 D6 專用「構面(含合併)」下拉清單，放在清單 AH 欄
    _face_list = list(LISTS["構面"]) + [nm for nm, _ in _MERGED_FACE]
    ws_list["AH1"] = "構面(含合併)"; ws_list["AH1"].font = F(9, True, WHITE); ws_list["AH1"].fill = fill(EVA_GREEN)
    for i, v in enumerate(_face_list, 2):
        ws_list[f"AH{i}"] = v
    FACE_PLUS_RANGE = f"'清單'!$AH$2:$AH${len(_face_list) + 1}"

    def _frow(fc):                 # "$D$6" -> 6
        return int(fc.split("$")[-1])
    _MROW = 6 + MASTER_ORDER.index("搭機月份")    # 互動查詢月份篩選列(=8)
    def crit_pairs(conds, colf):   # 產生 COUNTIFS/AVERAGEIFS 的「範圍,條件」配對字串
        out = []
        _lh = _MONTH_LH.get("互動查詢")
        for dc, fc in conds:
            if _frow(fc) == _MROW and _lh:        # 月份 → 介於 起月～迄月(範圍)
                out.append(_month_range_pair(colf(dc), _lh))
            else:
                out.append(f"{colf(dc)},{CRIT[_frow(fc)]}")
        return ",".join(out)

    SCORE = sat_col("M")           # 分數欄範圍(數值)

    # ---- 合併構面：移除構面後之基礎條件 + 兩段相加之度量產生器(供互動查詢) ----
    SAT_CONDS_BASE = [c for c in SAT_CONDS if c != ("L", "$D$6")]
    RSN_CONDS_BASE = [c for c in RSN_CONDS if c != ("J", "$D$6")]
    PAIRS_SAT_BASE = crit_pairs(SAT_CONDS_BASE, sat_col)
    PAIRS_RSN_BASE = crit_pairs(RSN_CONDS_BASE, rsn_col)
    SAT_FACE, RSN_FACE = sat_col("L"), rsn_col("J")   # 構面欄範圍(滿意度明細L / 原因明細J)

    def cnt_sat(extra=""):    # 兩段 COUNTIFS 相加(A 段 + B 段)，extra 為額外條件(如分數)
        return (f"(COUNTIFS({PAIRS_SAT_BASE},{SAT_FACE},{FACE_A}{extra})"
                f"+COUNTIFS({PAIRS_SAT_BASE},{SAT_FACE},{FACE_B}{extra}))")
    def sum_sat():            # 兩段 SUMIFS(分數)相加
        return (f"(SUMIFS({SCORE},{PAIRS_SAT_BASE},{SAT_FACE},{FACE_A})"
                f"+SUMIFS({SCORE},{PAIRS_SAT_BASE},{SAT_FACE},{FACE_B}))")
    def cnt_rsn(extra=""):    # 原因明細兩段 COUNTIFS 相加
        return (f"(COUNTIFS({PAIRS_RSN_BASE},{RSN_FACE},{FACE_A}{extra})"
                f"+COUNTIFS({PAIRS_RSN_BASE},{RSN_FACE},{FACE_B}{extra}))")

    # ============================ 5. 互動查詢儀表板 ============================
    ws = wb.create_sheet("互動查詢", 0)
    ws.sheet_properties.tabColor = EVA_GREEN
    ws.sheet_view.showGridLines = False
    # 欄寬
    for col, w in {"A": 2.5, "B": 14, "C": 9, "D": 20, "E": 2.5,
                   "F": 15, "G": 12, "H": 13, "I": 12, "J": 13, "K": 13, "L": 2.5}.items():
        ws.column_dimensions[col].width = w

    def merge_set(rng, value="", font=None, fillc=None, align=C, border=None, numfmt=None):
        ws.merge_cells(rng)
        top = rng.split(":")[0]
        cell = ws[top]
        cell.value = value
        if font: cell.font = font
        if fillc: cell.fill = fill(fillc)
        if align: cell.alignment = align
        if numfmt: cell.number_format = numfmt
        if border:
            for r in ws[rng]:
                for cc in r: cc.border = border
        return cell

    # 標題列
    merge_set("A1:L2", "長榮航空　旅客滿意度互動分析儀表板", F(20, True, WHITE), EVA_GREEN, C)
    merge_set("A3:L3", f"EVA AIR  Passenger Satisfaction Interactive Dashboard　│　資料期間：{PERIOD_TW} ({PERIOD_EN})",
              F(10, True, WHITE), EVA_MID, C)
    merge_set("A4:L4", f"操作說明：請於左側【篩選條件】點選下拉選單(可選「全部」)，右側分析結果將即時更新　│　"
                       f"問卷回覆 {df1['問卷編號'].nunique():,} 份，評分紀錄 {n_sat:,} 筆",
              F(9, False, GREY), WHITE, C)

    # ── 篩選條件區 (左) ──
    merge_set("B5:D5", "▍篩選條件", F(12, True, EVA_GREEN), EVA_LIGHT, L, mid_box)
    FILTERS = [(DIM_LABEL.get(d, d), d, 6 + i) for i, d in enumerate(MASTER_ORDER)]
    # ====================== Req3：依區域連動下拉(區域以下所有維度) ======================
    # 為「區域以下」每個維度，預先建立『每區域該維度實際出現的值』之靜態具名清單(<前綴>_<區域序號>)，
    # 並以 MATCH 取得所選區域之序號、INDIRECT 拼出對應清單。全部為一般公式(無動態陣列/無非法字元)，
    # Excel 各版本皆接受、openpyxl 可正確寫出，不會被修復移除。所有頁籤共用這些具名清單。
    _regions = [r for r in LISTS["區域"] if r != "全部"]              # 與清單 F 欄(rows3..)同序
    _BASE_COL = 40                  # 連動清單區塊起始欄(AN…)；避開既有 A–AK 欄
    _blk = [0]
    def _write_block(name, items):  # 於清單以一欄寫入「全部+items」並建立具名範圍
        col = get_column_letter(_BASE_COL + _blk[0]); _blk[0] += 1
        ws_list[f"{col}1"] = name; ws_list[f"{col}1"].font = F(8, True, WHITE); ws_list[f"{col}1"].fill = fill(EVA_MID)
        seq = ["全部"] + list(items)
        for i, v in enumerate(seq, 2):
            ws_list[f"{col}{i}"] = v
        wb.defined_names[name] = DefinedName(name, attr_text=f"清單!${col}$2:${col}${len(seq) + 1}")
    # 各依賴維度 × 區域 之具名清單：<前綴>_0=全部、<前綴>_<i>=第 i 區域實際出現的值
    for d in DEP_DIMS:
        k = DIMKEY[d]
        _write_block(f"{k}_0", [v for v in LISTS[d] if v != "全部"])
        for i, rg in enumerate(_regions, 1):
            vals = sorted(sat.loc[sat["區域"] == rg, d].dropna().astype(str).unique())
            _write_block(f"{k}_{i}", vals)
    # 班機編號【額外】依城市之清單(機城_0=全部、機城_<j>=第 j 城市之班機)，供「城市→班機」連動
    _all_cities2 = [c for c in LISTS["起飛城市"] if c != "全部"]
    _write_block("機城_0", [f for f in LISTS["班機編號"] if f != "全部"])
    for j, ct in enumerate(_all_cities2, 1):
        vals = sorted(sat.loc[sat["起飛城市"] == ct, "班機編號"].dropna().astype(str).unique())
        _write_block(f"機城_{j}", vals)

    # 連動輔助格配置器(清單 AK 欄)：每個下拉所需的序號/名稱字串各佔一格，確保唯一
    _REGION_LIST_RNG = f"清單!$F$3:$F${2 + len(_regions)}"
    _CITY_LIST_RNG = f"清單!$E$3:$E${2 + len(_all_cities2)}"
    _akrow = [1]
    def _alloc_cell(formula):
        _akrow[0] += 1
        ws_list[f"AK{_akrow[0]}"] = formula
        return f"清單!$AK${_akrow[0]}"
    def _region_idx(region_ref):     # 區域→序號(全部/找不到→0)
        return _alloc_cell(f"=IFERROR(MATCH({region_ref},{_REGION_LIST_RNG},0),0)")
    def _dep_source(dim, regidx_cell):   # 一般依賴維度：依區域連動
        return f'=INDIRECT("{DIMKEY[dim]}_"&{regidx_cell})'
    def _flight_source(region_ref, regidx_cell, city_ref):
        # 班機：選了城市→該城市班機；只選區域→該區域班機；皆全部→全部班機
        if city_ref is None:
            return f'=INDIRECT("機_"&{regidx_cell})'
        cityidx = _alloc_cell(f"=IFERROR(MATCH({city_ref},{_CITY_LIST_RNG},0),0)")
        namec = _alloc_cell(f'=IF({city_ref}<>"全部","機城_"&{cityidx},'
                            f'IF({region_ref}<>"全部","機_"&{regidx_cell},"機_0"))')
        return f'=INDIRECT({namec})'

    # 各篩選下拉來源：構面用「含合併」清單(Req1)；區域以下維度依區域連動；班機改依城市為主(Req3)。
    DV_SOURCE = {d: LIST_RANGE[d] for d in MASTER_ORDER}
    DV_SOURCE["構面"] = FACE_PLUS_RANGE
    DV_SOURCE["搭機月份"] = MONTH_DV_RANGE          # 月份下拉含季度/半年/單月
    _dash_reg_ref = f"{DASH}!$D${_DASH_ROW['區域']}"
    _dash_regidx = _region_idx(_dash_reg_ref)
    for d in DEP_DIMS:
        DV_SOURCE[d] = _dep_source(d, _dash_regidx)
    DV_SOURCE["班機編號"] = _flight_source(_dash_reg_ref, _dash_regidx, f"{DASH}!$D${_DASH_ROW['起飛城市']}")
    for label, key, r in FILTERS:
        ws[f"B{r}"] = label; ws[f"B{r}"].font = F(11, True, EVA_GREEN)
        ws[f"B{r}"].alignment = L; ws[f"B{r}"].fill = fill(EVA_LIGHT2); ws[f"B{r}"].border = thin_box
        ws[f"C{r}"].fill = fill(EVA_LIGHT2); ws[f"C{r}"].border = thin_box
        dcell = ws[f"D{r}"]
        dcell.value = "全部"; dcell.font = F(11, True, BLACK)
        dcell.alignment = C; dcell.fill = fill(INPUT_FILL); dcell.border = box(EVA_GOLD, "medium")
        dv = DataValidation(type="list", formula1=DV_SOURCE[key], allow_blank=False)
        dv.prompt = "請點選下拉選單"; dv.promptTitle = label
        ws.add_data_validation(dv); dv.add(dcell)

    # ── 分析結果區 (右) ──
    merge_set("F5:K5", "▍即時分析結果", F(12, True, EVA_GREEN), EVA_LIGHT, L, mid_box)

    # KPI 卡片 (三張)：平均分數 / 回覆筆數 / 滿意占比(4-5)
    def kpi(label_rng, val_rng, label, formula, numfmt, big=20, valcolor=EVA_GREEN, valfill=WHITE):
        merge_set(label_rng, label, F(10, True, WHITE), EVA_MID, C, thin_box)
        c = merge_set(val_rng, None, F(big, True, valcolor), valfill, C, thin_box, numfmt)
        c.value = formula
        return c

    kpi("F6:G6", "F7:G9", "平均滿意度分數",
        f'=IF($H$7<={MIN_N},IF($H$7=0,"無資料","—"),IFERROR({sum_sat()}/$H$7,"無資料"))', "0.00")
    kpi("H6:I6", "H7:I9", "回覆筆數",
        f"={cnt_sat()}", "#,##0", big=20, valcolor=EVA_MID)
    kpi("J6:K6", "J7:K9", "滿意占比 (4-5分)",
        f'=IFERROR({cnt_sat(f",{SCORE}," + chr(34) + ">=4" + chr(34))}/$H$7,"—")',
        "0.0%", big=20, valcolor=EVA_GREEN)

    # 第二排：待改善占比 / 綜合等級
    kpi("F10:G10", "F11:G12", "待改善占比 (≤3分)",
        f'=IFERROR({cnt_sat(f",{SCORE}," + chr(34) + "<=3" + chr(34))}/$H$7,"—")',
        "0.0%", big=18, valcolor=EVA_ORANGE)
    merge_set("H10:K10", "綜合滿意度等級", F(10, True, WHITE), EVA_MID, C, thin_box)
    avg_ref = "$F$7"
    grade_f = (f'=IFERROR(IF({avg_ref}>=4.5,"★★★★★  優異",'
               f'IF({avg_ref}>=4.2,"★★★★  良好",'
               f'IF({avg_ref}>=4.0,"★★★  普通",'
               f'IF({avg_ref}>=3.5,"★★  待加強","★  待改善")))),"—")')
    gc = merge_set("H11:K12", None, F(16, True, EVA_GOLD), WHITE, C, thin_box); gc.value = grade_f

    # 平均分數 F7：依「所選構面(D6)」取目標欄、「所選年度(D7)」取目標列，動態指向「年度目標達成」
    # 第三點對應目標儲存格(以 INDIRECT 組參照)；年度選「全部」或查無→回退當年度目標。
    # 所有輔助格放在獨立隱藏頁「目標參照」，避免覆蓋「清單」頁的下拉來源(城市/班機等)。
    _tgt_items = list(TARGET_NOW.items())
    for _i, (_nm, _tg) in enumerate(_tgt_items):
        wsT[f"A{2 + _i}"] = _nm                              # 構面/合併名稱
        wsT[f"B{2 + _i}"] = ITEM_TGT_COL.get(_nm, "")        # 第三點欄字母
    _ALAN_TBL = f"{TGT}!$A$2:$B${1 + len(_tgt_items)}"
    for _k, _yr in enumerate(YEAR_ROWS):
        wsT[f"D{2 + _k}"] = _yr                              # 年度
        wsT[f"E{2 + _k}"] = _Y_TG_R0 + _k                    # 對應第三點目標列
    _YR_TBL = f"{TGT}!$D$2:$E${1 + len(YEAR_ROWS)}"
    wsT["A12"] = f'=IFERROR(VLOOKUP({DASH}!$D$6,{_ALAN_TBL},2,FALSE),"")'                  # 構面→目標欄字母
    wsT["B12"] = f'=IFERROR(VLOOKUP({DASH}!$D$7,{_YR_TBL},2,FALSE),{_Y_TG_CUR_ROW})'       # 年度→目標列(全部/查無→當年度)
    wsT["C12"] = '=IF($A$12="","",INDIRECT("\'' + YSHEET + '\'!"&$A$12&$B$12))'             # 動態目標值
    add_target_gradient(ws, "F7", "$F$7", f"{TGT}!$C$12", target_is_cell=True)

    # ---- 問卷項目總覽/交叉分析/城市排行：著色目標亦隨『各頁自己的年度篩選』動態切換(同 F7 機制) ----
    #   於「目標參照」頁建立：年度→目標列(G欄) + 各構面動態目標(H..L欄，以 INDIRECT 取年度目標達成第三點)。
    #   各頁年度輸入格 = 標題欄(label_col)+1、列 = 面板首列(top_row)+該頁維度中『搭機年度』之索引。
    _DYN_SHEETS = [("問卷項目總覽", OV_DIMS, 10, 4, 2),
                   ("交叉分析",     CR_DIMS,  9, 4, 3),
                   ("城市排行",     CT_DIMS, 11, 4, 4)]
    SHEET_TGT = {}
    for _sn, _dims, _lcol, _trow, _hrow in _DYN_SHEETS:
        _yf = f"'{_sn}'!${get_column_letter(_lcol + 1)}${_trow + _dims.index('搭機年度')}"   # 該頁年度輸入格
        wsT[f"G{_hrow}"] = f'=IFERROR(VLOOKUP({_yf},{_YR_TBL},2,FALSE),{_Y_TG_CUR_ROW})'        # 年度→目標列
        _rowref = f"$G${_hrow}"
        _m = {}
        for _kk, _dim in enumerate(DIM_ORDER):                # 5 構面 → H..L 欄
            _hc = get_column_letter(8 + _kk)                  # H=8
            wsT[f"{_hc}{_hrow}"] = f'=INDIRECT("\'{YSHEET}\'!{ITEM_TGT_COL[_dim]}"&{_rowref})'
            _m[_dim] = f"{TGT}!${_hc}${_hrow}"
        SHEET_TGT[_sn] = _m

    # ── 分數分布 (左下；因新增「區域」篩選，整體下移一列至18起) ──
    merge_set("B18:D18", "▍分數分布", F(12, True, EVA_GREEN), EVA_LIGHT, L, mid_box)
    hdr = ["分數", "筆數", "占比"]
    for j, h in enumerate(hdr):
        cell = ws.cell(row=19, column=2 + j, value=h)
        cell.font = F(10, True, WHITE); cell.fill = fill(EVA_GREEN); cell.alignment = C; cell.border = thin_box
    score_labels = {5: "5分 非常滿意", 4: "4分 滿意", 3: "3分 普通", 2: "2分 不滿意", 1: "1分 非常不滿意"}
    for idx, sc in enumerate([5, 4, 3, 2, 1]):
        r = 20 + idx
        a = ws.cell(row=r, column=2, value=score_labels[sc]); a.font = F(10); a.alignment = L; a.border = thin_box
        b = ws.cell(row=r, column=3); b.value = f'={cnt_sat(f",{SCORE}," + chr(34) + "=" + str(sc) + chr(34))}'
        b.font = F(10); b.alignment = Rr; b.number_format = "#,##0"; b.border = thin_box
        c = ws.cell(row=r, column=4); c.value = f'=IFERROR(C{r}/$H$7,"—")'
        c.font = F(10); c.alignment = Rr; c.number_format = "0.0%"; c.border = thin_box
        a.fill = fill(EVA_LIGHT2 if idx % 2 else WHITE)
        b.fill = fill(EVA_LIGHT2 if idx % 2 else WHITE)
        c.fill = fill(EVA_LIGHT2 if idx % 2 else WHITE)
    # 合計列
    r = 25
    ws.cell(row=r, column=2, value="合計").font = F(10, True)
    ws.cell(row=r, column=2).alignment = L; ws.cell(row=r, column=2).border = thin_box
    ws.cell(row=r, column=2).fill = fill(EVA_LIGHT)
    tb = ws.cell(row=r, column=3, value="=SUM(C20:C24)"); tb.font = F(10, True); tb.alignment = Rr
    tb.number_format = "#,##0"; tb.border = thin_box; tb.fill = fill(EVA_LIGHT)
    tc = ws.cell(row=r, column=4, value="=IFERROR(SUM(D20:D24),0)"); tc.font = F(10, True); tc.alignment = Rr
    tc.number_format = "0.0%"; tc.border = thin_box; tc.fill = fill(EVA_LIGHT)
    # 筆數 資料條
    ws.conditional_formatting.add("C20:C24",
        DataBarRule(start_type="num", start_value=0, end_type="max",
                    color=EVA_MID, showValue=True, minLength=None, maxLength=None))

    # ── 待改善原因 TOP10 (右下) ──
    merge_set("F17:K17", "▍三分(含)以下 待改善原因 TOP 10", F(12, True, EVA_GREEN), EVA_LIGHT, L, mid_box)
    ws.merge_cells("G18:J18")
    for col, h in [("F", "排名"), ("G", "待改善原因"), ("K", "提及次數")]:
        cell = ws[f"{col}18"]; cell.value = h; cell.font = F(10, True, WHITE)
        cell.fill = fill(EVA_GREEN); cell.alignment = C; cell.border = thin_box
    ws["H18"].fill = fill(EVA_GREEN); ws["I18"].fill = fill(EVA_GREEN); ws["J18"].fill = fill(EVA_GREEN)
    ws["H18"].border = thin_box; ws["I18"].border = thin_box; ws["J18"].border = thin_box

    # 隱藏計算區(原因動態次數)放在 N:P 欄，rows 6.. (29 項)
    ws.column_dimensions["N"].hidden = True
    ws.column_dimensions["O"].hidden = True
    ws.column_dimensions["P"].hidden = True
    calc_start = 6
    for i, name in enumerate(REASON_LIST):
        rr = calc_start + i
        ws[f"N{rr}"] = name
        _rsn_opt = rsn_col("Q")          # 原因明細 Q 欄=待改善原因(細項優先，空白取原因選項)
        ws[f"O{rr}"] = "=" + cnt_rsn(f",{_rsn_opt},$N{rr}")
        ws[f"O{rr}"].number_format = "#,##0"
        # 加極小排序鍵以打破同分(以列序)
        ws[f"P{rr}"] = f"=O{rr}+ROW()/100000"
    calc_end = calc_start + len(REASON_LIST) - 1
    NRNG = f"$N${calc_start}:$N${calc_end}"
    ORNG = f"$O${calc_start}:$O${calc_end}"
    PRNG = f"$P${calc_start}:$P${calc_end}"
    # TOP10 顯示
    for k in range(1, 11):
        r = 18 + k
        ws.merge_cells(f"G{r}:J{r}")
        rank = ws[f"F{r}"]; rank.value = k; rank.font = F(10, True, EVA_GREEN); rank.alignment = C; rank.border = thin_box
        nm = ws[f"G{r}"]
        nm.value = (f'=IFERROR(INDEX({NRNG},MATCH(LARGE({PRNG},{k}),{PRNG},0)),"—")')
        nm.font = F(10); nm.alignment = L; nm.border = thin_box
        cnt = ws[f"K{r}"]
        cnt.value = (f'=IFERROR(INDEX({ORNG},MATCH(LARGE({PRNG},{k}),{PRNG},0)),"—")')
        cnt.font = F(10, True, EVA_ORANGE); cnt.alignment = Rr; cnt.number_format = "#,##0"; cnt.border = thin_box
        for cc in [f"H{r}", f"I{r}", f"J{r}"]:
            ws[cc].border = thin_box
        fillc = EVA_LIGHT2 if k % 2 == 0 else WHITE
        for col in ["F", "G", "H", "I", "J", "K"]:
            if ws[f"{col}{r}"].fill.fgColor.rgb in (None, "00000000"):
                ws[f"{col}{r}"].fill = fill(fillc)

    # 註腳（移到 TOP10 表格下方，避免覆蓋；TOP10 佔 19~28 列）
    merge_set("B30:L31",
        "【名詞說明】回覆筆數＝符合篩選之評分紀錄數；選擇單一「滿意度構面」時即等於該題回覆人數。"
        "「待改善原因」僅統計評分 3 分(含)以下之填答，可複選，故提及次數加總會大於人數。"
        "　艙等：商務艙(C)、豪華經濟艙(K)、經濟艙(Y)。詳細代碼對照請見「使用說明」分頁。",
        F(8.5, False, GREY), EVA_LIGHT2, L, thin_box)

    # 列高
    ws.row_dimensions[1].height = 20; ws.row_dimensions[2].height = 20
    for r in [7, 8, 9, 11, 12]:
        ws.row_dimensions[r].height = 22

    # ============================ 6. 問卷項目總覽 ============================
    ws2 = wb.create_sheet("問卷項目總覽")
    ws2.sheet_properties.tabColor = EVA_MID
    ws2.sheet_view.showGridLines = False
    for col, w in {"A": 3, "B": 18, "C": 12, "D": 12, "E": 13, "F": 13, "G": 13, "H": 24}.items():
        ws2.column_dimensions[col].width = w
    ws2.merge_cells("B2:H2")
    ws2["B2"] = f"各問卷項目滿意度總覽　({PERIOD_TW_TIGHT})"
    ws2["B2"].font = F(15, True, WHITE); ws2["B2"].fill = fill(EVA_GREEN); ws2["B2"].alignment = C
    for r in ws2["B2:H2"]:
        for cc in r: cc.fill = fill(EVA_GREEN)
    # 本頁專用篩選面板(右側 J/K 欄)；OV_P 套入各公式，最主要原因隨之連動(矩陣已含篩選)
    build_filter_panel_v(ws2, "問卷項目總覽", OV_DIMS, CRIT_OV, label_col=10, top_row=4)
    OV_P = sat_pairs(OV_DIMS, CRIT_OV)
    heads = ["問卷項目", "平均分數", "回覆人數", "滿意占比(4-5)", "普通占比(3)", "待改善占比(≤3)", "最主要待改善原因"]
    for j, h in enumerate(heads):
        cell = ws2.cell(row=4, column=2 + j, value=h)
        cell.font = F(10.5, True, WHITE); cell.fill = fill(EVA_MID); cell.alignment = C; cell.border = thin_box
    for i, dim in enumerate(DIM_ORDER):
        r = 5 + i
        a = ws2.cell(row=r, column=2, value=dim); a.font = F(11, True, EVA_GREEN); a.alignment = L
        smean = f'=IF($D{r}<={MIN_N},"—",IFERROR(AVERAGEIFS({sat_col("M")},{sat_col("L")},$B{r},{OV_P}),"—"))'
        cnt = f'=COUNTIFS({sat_col("L")},$B{r},{OV_P})'
        ws2.cell(row=r, column=3, value=smean).number_format = "0.00"
        ws2.cell(row=r, column=4, value=cnt).number_format = "#,##0"
        ws2.cell(row=r, column=5,
            value=f'=IFERROR(COUNTIFS({sat_col("L")},$B{r},{OV_P},{sat_col("M")},">=4")/$D{r},"—")').number_format = "0.0%"
        ws2.cell(row=r, column=6,
            value=f'=IFERROR(COUNTIFS({sat_col("L")},$B{r},{OV_P},{sat_col("M")},"=3")/$D{r},"—")').number_format = "0.0%"
        ws2.cell(row=r, column=7,
            value=f'=IFERROR(COUNTIFS({sat_col("L")},$B{r},{OV_P},{sat_col("M")},"<=3")/$D{r},"—")').number_format = "0.0%"
        # 最主要原因：該構面下原因次數最大者(取自隱藏矩陣，已隨本頁篩選連動)
        gcol = GRID_COL[dim]
        grid_rng = f"清單!${gcol}${G_REASON_LO}:${gcol}${G_REASON_HI}"
        top_reason = (f'=IFERROR(INDEX({GRID_REASON_RNG},'
                      f'MATCH(MAX({grid_rng}),{grid_rng},0)),"—")')
        tr = ws2.cell(row=r, column=8, value=top_reason); tr.font = F(9.5); tr.alignment = L
        for col in range(2, 9):
            cell = ws2.cell(row=r, column=col); cell.border = thin_box
            if cell.alignment is None or cell.alignment.horizontal is None:
                cell.alignment = C
            if col >= 3 and col != 8:
                cell.alignment = Rr if col == 4 else C
            cell.fill = fill(EVA_LIGHT2 if i % 2 else WHITE)
    # 各構面平均分數：依該構面之目標值雙向著色(達標綠/未達標紅)
    for i, dim in enumerate(DIM_ORDER):
        r = 5 + i
        if dim in SHEET_TGT["問卷項目總覽"]:
            add_target_gradient(ws2, f"C{r}", f"C{r}", SHEET_TGT["問卷項目總覽"][dim], target_is_cell=True)
    note2 = ws2.cell(row=5 + len(DIM_ORDER) + 1, column=2,
        value="說明：滿意占比＝評4或5分之比例；待改善占比＝評3分(含)以下之比例。「最主要待改善原因」取自該構面3分以下填答中提及最多者。")
    note2.font = F(8.5, False, GREY); ws2.merge_cells(f"B{5+len(DIM_ORDER)+1}:H{5+len(DIM_ORDER)+1}")

    # ============================ 7. 交叉分析 (卡別/艙等/機型 × 構面 平均分數) ============================
    ws3 = wb.create_sheet("交叉分析")
    ws3.sheet_properties.tabColor = EVA_MID
    ws3.sheet_view.showGridLines = False
    ws3.column_dimensions["A"].width = 3
    ws3.column_dimensions["B"].width = 16
    for col in ["C", "D", "E", "F", "G"]:
        ws3.column_dimensions[col].width = 14
    # 本頁專用篩選面板(右側 I/J 欄)；非交叉軸的維度才放入篩選
    build_filter_panel_v(ws3, "交叉分析", CR_DIMS, CRIT_CR, label_col=9, top_row=4)
    CR_P = sat_pairs(CR_DIMS, CRIT_CR)

    def cross_block(start_row, title, row_field_col, row_values, value_label="平均滿意度分數"):
        ws3.merge_cells(f"B{start_row}:G{start_row}")
        t = ws3[f"B{start_row}"]; t.value = title; t.font = F(13, True, WHITE)
        t.fill = fill(EVA_GREEN); t.alignment = L
        for r in ws3[f"B{start_row}:G{start_row}"]:
            for cc in r: cc.fill = fill(EVA_GREEN)
        hr = start_row + 1
        ws3.cell(row=hr, column=2, value="").border = thin_box
        h0 = ws3.cell(row=hr, column=2, value=value_label); h0.font = F(10, True, WHITE)
        h0.fill = fill(EVA_MID); h0.alignment = C; h0.border = thin_box
        for j, dim in enumerate(DIM_ORDER):
            cell = ws3.cell(row=hr, column=3 + j, value=dim)
            cell.font = F(9.5, True, WHITE); cell.fill = fill(EVA_MID); cell.alignment = C; cell.border = thin_box
        for i, rv in enumerate(row_values):
            r = hr + 1 + i
            a = ws3.cell(row=r, column=2, value=rv); a.font = F(10.5, True, EVA_GREEN); a.alignment = L; a.border = thin_box
            a.fill = fill(EVA_LIGHT2 if i % 2 else WHITE)
            for j, dim in enumerate(DIM_ORDER):
                col_letter = get_column_letter(3 + j)
                _cnt = (f'COUNTIFS({sat_col(row_field_col)},$B{r},'
                        f'{sat_col("L")},{col_letter}${hr},{CR_P})')
                f_ = (f'=IF({_cnt}<={MIN_N},"—",'
                      f'IFERROR(AVERAGEIFS({sat_col("M")},{sat_col(row_field_col)},$B{r},'
                      f'{sat_col("L")},{col_letter}${hr},{CR_P}),"—"))')
                cell = ws3.cell(row=r, column=3 + j, value=f_)
                cell.font = F(10); cell.alignment = C; cell.number_format = "0.00"; cell.border = thin_box
                cell.fill = fill(EVA_LIGHT2 if i % 2 else WHITE)
        last = hr + len(row_values)
        # 各構面欄(C..G)依該構面之目標值雙向著色(引用年度目標達成第三點，隨本頁年度篩選切換)
        for j, dim in enumerate(DIM_ORDER):
            if dim in SHEET_TGT["交叉分析"]:
                col = get_column_letter(3 + j)
                add_target_gradient(ws3, f"{col}{hr+1}:{col}{last}", f"{col}{hr+1}",
                                    SHEET_TGT["交叉分析"][dim], target_is_cell=True)
        return last + 2

    row = 2
    row = cross_block(row, "會員卡別 × 服務構面　平均滿意度分數", "B",
                      [v for v in LISTS["會員卡別"] if v != "全部"])
    row = cross_block(row, "艙等 × 服務構面　平均滿意度分數", "H",
                      [v for v in LISTS["艙等"] if v != "全部"])
    row = cross_block(row, "機型 × 服務構面　平均滿意度分數", "G",
                      [v for v in LISTS["機型"] if v != "全部"])
    row = cross_block(row, "區域 × 服務構面　平均滿意度分數", "N",
                      [v for v in LISTS["區域"] if v != "全部"])
    ws3.cell(row=row, column=2,
        value="色彩說明：綠色代表分數較高、紅色代表分數較低，可快速辨識相對弱項。「—」表示該組合無樣本。").font = F(8.5, False, GREY)
    ws3.merge_cells(f"B{row}:G{row}")

    # ============================ 8. 起飛城市排行 ============================
    ws4 = wb.create_sheet("城市排行")
    ws4.sheet_properties.tabColor = EVA_MID
    ws4.sheet_view.showGridLines = False
    ws4.column_dimensions["A"].width = 3
    ws4.column_dimensions["B"].width = 11
    ws4.column_dimensions["C"].width = 14   # 區域
    ws4.column_dimensions["D"].width = 11   # 總回覆人數
    for col in ["E", "F", "G", "H", "I"]:
        ws4.column_dimensions[col].width = 13
    ws4.merge_cells("B2:I2")
    ws4["B2"] = "起飛城市 × 服務構面　平均滿意度分數排行　(含區域；可點欄位篩選/排序)"
    ws4["B2"].font = F(14, True, WHITE); ws4["B2"].alignment = L
    for r in ws4["B2:I2"]:
        for cc in r: cc.fill = fill(EVA_GREEN)
    heads4 = ["起飛城市", "區域", "總回覆人數"] + DIM_ORDER
    for j, h in enumerate(heads4):
        cell = ws4.cell(row=4, column=2 + j, value=h)
        cell.font = F(9.5, True, WHITE); cell.fill = fill(EVA_MID); cell.alignment = C; cell.border = thin_box
    # 本頁專用篩選面板(右側 K/L 欄)；城市為列軸，故篩選不含起飛城市(但含區域，可只看某區域內各站)
    build_filter_panel_v(ws4, "城市排行", CT_DIMS, CRIT_CT, label_col=11, top_row=4)
    CT_P = sat_pairs(CT_DIMS, CRIT_CT)
    cities = [v for v in LISTS["起飛城市"] if v != "全部"]
    for i, city in enumerate(cities):
        r = 5 + i
        a = ws4.cell(row=r, column=2, value=city); a.font = F(10, True, EVA_GREEN); a.alignment = C; a.border = thin_box
        # 區域(靜態對應，依起飛城市)
        rg = ws4.cell(row=r, column=3, value=CITY_TO_REGION.get(city, "其他"))
        rg.font = F(9.5, True, EVA_MID); rg.alignment = C; rg.border = thin_box
        tot = ws4.cell(row=r, column=4, value=f'=COUNTIFS({sat_col("I")},$B{r},{CT_P})')
        tot.font = F(10); tot.alignment = Rr; tot.number_format = "#,##0"; tot.border = thin_box
        for j, dim in enumerate(DIM_ORDER):
            col_letter = get_column_letter(5 + j)
            _cnt = (f'COUNTIFS({sat_col("I")},$B{r},'
                    f'{sat_col("L")},{col_letter}$4,{CT_P})')
            f_ = (f'=IF({_cnt}<={MIN_N},"—",'
                  f'IFERROR(AVERAGEIFS({sat_col("M")},{sat_col("I")},$B{r},'
                  f'{sat_col("L")},{col_letter}$4,{CT_P}),"—"))')
            cell = ws4.cell(row=r, column=5 + j, value=f_)
            cell.font = F(10); cell.alignment = C; cell.number_format = "0.00"; cell.border = thin_box
        for col in range(2, 10):
            ws4.cell(row=r, column=col).fill = fill(EVA_LIGHT2 if i % 2 else WHITE)
    last4 = 4 + len(cities)
    ws4.auto_filter.ref = f"B4:I{last4}"
    ws4.freeze_panes = "B5"
    # 各構面欄(E..I)依該構面之目標值雙向著色(引用年度目標達成第三點，隨本頁年度篩選切換)
    for j, dim in enumerate(DIM_ORDER):
        if dim in SHEET_TGT["城市排行"]:
            col = get_column_letter(5 + j)
            add_target_gradient(ws4, f"{col}5:{col}{last4}", f"{col}5",
                                SHEET_TGT["城市排行"][dim], target_is_cell=True)
    ws4.cell(row=last4 + 2, column=2,
        value="提示：點選標題列篩選箭頭可依任一構面或【區域】排序/篩選，快速找出各站、各區域表現相對較弱之服務環節。").font = F(8.5, False, GREY)
    ws4.merge_cells(f"B{last4+2}:I{last4+2}")

    # ============================ 9. 原因解析 (本頁可獨立篩選，全 29 項) ============================
    ws5 = wb.create_sheet("原因解析")
    ws5.sheet_properties.tabColor = EVA_ORANGE
    ws5.sheet_view.showGridLines = False
    ws5.column_dimensions["A"].width = 3
    ws5.column_dimensions["B"].width = 6
    ws5.column_dimensions["C"].width = 40
    ws5.column_dimensions["D"].width = 12
    ws5.column_dimensions["E"].width = 12
    ws5.merge_cells("B2:E2")
    ws5["B2"] = "三分(含)以下　待改善原因完整解析"
    ws5["B2"].font = F(15, True, WHITE); ws5["B2"].alignment = L
    for r in ws5["B2:E2"]:
        for cc in r: cc.fill = fill(EVA_GREEN)
    ws5.merge_cells("B3:E3")
    ws5["B3"] = "※ 本頁可用右側【篩選條件】單獨篩選(含起飛城市、搭機年度/月份、構面等)；統計範圍即時連動，與互動查詢分頁互不影響。"
    ws5["B3"].font = F(9, False, GREY); ws5["B3"].alignment = L; ws5["B3"].fill = fill(EVA_LIGHT2)
    for r in ws5["B3:E3"]:
        for cc in r: cc.fill = fill(EVA_LIGHT2)
    heads5 = ["排名", "待改善原因", "提及次數", "占全部比"]
    for j, h in enumerate(heads5):
        cell = ws5.cell(row=5, column=2 + j, value=h)
        cell.font = F(10.5, True, WHITE); cell.fill = fill(EVA_MID); cell.alignment = C; cell.border = thin_box
    # 本頁專用篩選面板(右側 G/H 欄，全 11 維度)
    build_filter_panel_v(ws5, "原因解析", RS_DIMS, CRIT_RS, label_col=7, top_row=4)
    RS_RP = rsn_pairs(RS_DIMS, CRIT_RS)
    # 本頁專用隱藏計算區(原因動態次數，隨本頁篩選連動)放在 N/O/P 欄
    for _c in ("N", "O", "P"):
        ws5.column_dimensions[_c].hidden = True
    rs_calc_start = 6
    for i, name in enumerate(REASON_LIST):
        rr = rs_calc_start + i
        ws5[f"N{rr}"] = name
        ws5[f"O{rr}"] = f"=COUNTIFS({rsn_col('Q')},$N{rr},{RS_RP})"
        ws5[f"O{rr}"].number_format = "#,##0"
        ws5[f"P{rr}"] = f"=O{rr}+ROW()/100000"
    rs_calc_end = rs_calc_start + len(REASON_LIST) - 1
    NRNG2 = f"原因解析!$N${rs_calc_start}:$N${rs_calc_end}"
    ORNG2 = f"原因解析!$O${rs_calc_start}:$O${rs_calc_end}"
    PRNG2 = f"原因解析!$P${rs_calc_start}:$P${rs_calc_end}"
    total_mention = f"SUM({ORNG2})"
    for k in range(1, len(REASON_LIST) + 1):
        r = 5 + k
        rank = ws5.cell(row=r, column=2, value=k); rank.font = F(10, True, EVA_GREEN); rank.alignment = C; rank.border = thin_box
        nm = ws5.cell(row=r, column=3,
            value=f'=IFERROR(INDEX({NRNG2},MATCH(LARGE({PRNG2},{k}),{PRNG2},0)),"—")')
        nm.font = F(10); nm.alignment = L; nm.border = thin_box
        cnt = ws5.cell(row=r, column=4,
            value=f'=IFERROR(INDEX({ORNG2},MATCH(LARGE({PRNG2},{k}),{PRNG2},0)),"—")')
        cnt.font = F(10, True, EVA_ORANGE); cnt.alignment = Rr; cnt.number_format = "#,##0"; cnt.border = thin_box
        pct = ws5.cell(row=r, column=5,
            value=f'=IFERROR(D{r}/{total_mention},"—")')
        pct.font = F(10); pct.alignment = Rr; pct.number_format = "0.0%"; pct.border = thin_box
        for col in range(2, 6):
            ws5.cell(row=r, column=col).fill = fill(EVA_LIGHT2 if k % 2 == 0 else WHITE)
    last5 = 5 + len(REASON_LIST)
    # 合計
    tr = last5 + 1
    ws5.cell(row=tr, column=3, value="合計").font = F(10, True)
    ws5.cell(row=tr, column=3).alignment = Rr; ws5.cell(row=tr, column=3).border = thin_box
    ws5.cell(row=tr, column=3).fill = fill(EVA_LIGHT)
    tc = ws5.cell(row=tr, column=4, value=f"=SUM(D6:D{last5})"); tc.font = F(10, True); tc.alignment = Rr
    tc.number_format = "#,##0"; tc.border = thin_box; tc.fill = fill(EVA_LIGHT)
    ws5.cell(row=tr, column=5, value="100.0%").font = F(10, True); ws5.cell(row=tr, column=5).alignment = Rr
    ws5.cell(row=tr, column=5).number_format = "0.0%"; ws5.cell(row=tr, column=5).border = thin_box
    ws5.cell(row=tr, column=5).fill = fill(EVA_LIGHT)
    ws5.conditional_formatting.add(f"D6:D{last5}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color=EVA_ORANGE, showValue=True))

    # ============================ 9b. 年度目標達成 (加權平均 + 合併項目 + 目標色彩) ============================
    # 計算方式：滿意度平均 = Σ(分數ᵢ × 人數ᵢ) / Σ人數ᵢ。本資料每列為一筆評分(人數=1)，
    # 故等同 Σ加權分子(分數) / Σ回覆人數(筆數)，以 SUMIFS/COUNTIFS 明確呈現；
    # 回覆人數 <= 5 不列計平均(顯示「—」)，惟回覆人數一律照實顯示。
    wsY = wb.create_sheet("年度目標達成")
    wsY.sheet_properties.tabColor = EVA_GOLD
    wsY.sheet_view.showGridLines = False
    wsY.column_dimensions["A"].width = 3
    wsY.column_dimensions["B"].width = 9
    for j in range(len(YEAR_ITEMS)):
        wsY.column_dimensions[get_column_letter(3 + j)].width = 15
    NI = len(YEAR_ITEMS)
    last_col = get_column_letter(2 + NI)

    def _y_band(rrow, text, sub=None):
        wsY.merge_cells(f"B{rrow}:{last_col}{rrow}")
        c = wsY[f"B{rrow}"]; c.value = text; c.font = F(13, True, WHITE); c.alignment = L
        for rr in wsY[f"B{rrow}:{last_col}{rrow}"]:
            for cc in rr: cc.fill = fill(EVA_GREEN)

    def _y_header(hr):
        h0 = wsY.cell(row=hr, column=2, value="年度"); h0.font = F(10, True, WHITE)
        h0.fill = fill(EVA_MID); h0.alignment = C; h0.border = thin_box
        for j, (nm, _cs) in enumerate(YEAR_ITEMS):
            cell = wsY.cell(row=hr, column=3 + j, value=nm)
            cell.font = F(9, True, WHITE); cell.fill = fill(EVA_MID); cell.alignment = C; cell.border = thin_box

    # 標題
    wsY.merge_cells(f"B2:{last_col}2")
    wsY["B2"] = f"各年度服務項目滿意度　目標達成總覽　({PERIOD_TW_TIGHT})"
    wsY["B2"].font = F(15, True, WHITE); wsY["B2"].alignment = C
    for rr in wsY[f"B2:{last_col}2"]:
        for cc in rr: cc.fill = fill(EVA_GREEN)
    wsY.merge_cells(f"B3:{last_col}3")
    wsY["B3"] = ("※ 平均 = Σ(分數×人數)/Σ人數(本資料每筆人數=1，以 SUMIFS/COUNTIFS 加權呈現)；"
                 "回覆人數 ≤ 5 不列計平均(顯示「—」)；2026 起新增合併項目；達標綠/未達標紅。")
    wsY["B3"].font = F(9, False, GREY); wsY["B3"].alignment = L; wsY["B3"].fill = fill(EVA_LIGHT2)
    for rr in wsY[f"B3:{last_col}3"]:
        for cc in rr: cc.fill = fill(EVA_LIGHT2)

    # 區段列位(沿用前段提前計算之常數，確保其他頁籤引用之目標值儲存格位置一致)
    SC_HR, SC_R0 = _Y_SC_HR, _Y_SC_R0      # 平均分數：表頭列、資料首列
    CN_HR, CN_R0 = _Y_CN_HR, _Y_CN_R0      # 回覆人數
    TG_HR, TG_R0 = _Y_TG_HR, _Y_TG_R0      # 目標值

    # ---- 區段一：平均滿意度分數(加權，含合併項目，n<=5 -> 「—」) ----
    _y_band(SC_HR - 1, "一、平均滿意度分數（加權；回覆人數 ≤ 5 顯示「—」）")
    _y_header(SC_HR)
    for yi, yr in enumerate(YEAR_ROWS):
        r = SC_R0 + yi
        yb = wsY.cell(row=r, column=2, value=yr); yb.font = F(11, True, EVA_GREEN); yb.alignment = C; yb.border = thin_box
        yb.fill = fill(EVA_LIGHT2 if yi % 2 else WHITE)
        for j, (nm, cs) in enumerate(YEAR_ITEMS):
            denom = "+".join(f'COUNTIFS({sat_col("L")},"{c}",{sat_col("D")},$B{r})' for c in cs)
            num   = "+".join(f'SUMIFS({sat_col("M")},{sat_col("L")},"{c}",{sat_col("D")},$B{r})' for c in cs)
            f_ = f'=IF(({denom})<={MIN_N},"—",IFERROR(({num})/({denom}),"—"))'
            cell = wsY.cell(row=r, column=3 + j, value=f_)
            cell.font = F(10.5, True); cell.alignment = C; cell.number_format = "0.00"; cell.border = thin_box
            cell.fill = fill(EVA_LIGHT2 if yi % 2 else WHITE)
    SC_RN = SC_R0 + len(YEAR_ROWS) - 1

    # ---- 區段二：回覆人數(即使 n<=5 仍照實顯示) ----
    _y_band(CN_HR - 1, "二、回覆人數（資料照實顯示；供判讀樣本是否足夠）")
    _y_header(CN_HR)
    for yi, yr in enumerate(YEAR_ROWS):
        r = CN_R0 + yi
        yb = wsY.cell(row=r, column=2, value=yr); yb.font = F(11, True, EVA_GREEN); yb.alignment = C; yb.border = thin_box
        yb.fill = fill(EVA_LIGHT2 if yi % 2 else WHITE)
        for j, (nm, cs) in enumerate(YEAR_ITEMS):
            denom = "+".join(f'COUNTIFS({sat_col("L")},"{c}",{sat_col("D")},$B{r})' for c in cs)
            cell = wsY.cell(row=r, column=3 + j, value=f"=({denom})")
            cell.font = F(10); cell.alignment = C; cell.number_format = "#,##0"; cell.border = thin_box
            cell.fill = fill(EVA_LIGHT2 if yi % 2 else WHITE)

    # ---- 區段三：目標值(可自行編修；空白=該年度尚未適用) ----
    _y_band(TG_HR - 1, "三、目標值（可自行編修；達標綠、未達標紅之比對基準）")
    _y_header(TG_HR)
    for yi, yr in enumerate(YEAR_ROWS):
        r = TG_R0 + yi
        yb = wsY.cell(row=r, column=2, value=yr); yb.font = F(11, True, EVA_GREEN); yb.alignment = C; yb.border = thin_box
        yb.fill = fill(EVA_LIGHT2 if yi % 2 else WHITE)
        tgs = YEAR_TARGETS.get(yr, [None] * NI)
        for j in range(NI):
            v = tgs[j] if j < len(tgs) else None
            cell = wsY.cell(row=r, column=3 + j, value=v)
            cell.font = F(10.5, True, EVA_GOLD); cell.alignment = C; cell.number_format = "0.00"; cell.border = thin_box
            cell.fill = fill(INPUT_FILL)   # 黃底=可編修

    # ---- 區段四：合併項目對照(增列清單) ----
    LG_HR = TG_R0 + len(YEAR_ROWS) + 2
    wsY.merge_cells(f"B{LG_HR}:{last_col}{LG_HR}")
    lg = wsY[f"B{LG_HR}"]; lg.value = "四、2026 新增合併項目對照（清單）"; lg.font = F(11, True, EVA_GREEN)
    lg.fill = fill(EVA_LIGHT); lg.alignment = L
    for rr in wsY[f"B{LG_HR}:{last_col}{LG_HR}"]:
        for cc in rr: cc.fill = fill(EVA_LIGHT)
    merged = [(nm, cs) for nm, cs in YEAR_ITEMS if len(cs) > 1]
    for mi, (nm, cs) in enumerate(merged):
        r = LG_HR + 1 + mi
        a = wsY.cell(row=r, column=2, value=nm); a.font = F(10, True, EVA_MID); a.alignment = L; a.border = thin_box
        wsY.merge_cells(f"C{r}:{last_col}{r}")
        b = wsY.cell(row=r, column=3, value="＝ " + " ＋ ".join(cs)); b.font = F(10); b.alignment = L; b.border = thin_box

    # ---- 色彩管理：平均分數 vs 目標值(同欄、固定位移) ----
    # 雙向漸層：達標(≥目標)由淺到深【綠】，越高越綠；未達標(<目標)由淺到深【紅】，越低越紅。
    # 不使用黃色等中間色；以「分數−目標」之差距分段，stopIfTrue 由最極端往回判定。
    # ISNUMBER 排除「—」(樣本不足)；目標<>"" 排除尚未適用之合併項目(2024/2025) → 皆不著色。
    score_rng = f"C{SC_R0}:{last_col}{SC_RN}"
    _DARK_FONT = "FFFFFF"
    # (差距下限, 底色, 字色)；綠：由深到淺，紅：由淺到深
    GREEN_BANDS = [(0.20, "1B7E3F", _DARK_FONT), (0.10, "4CAE6A", "06402B"),
                   (0.03, "86CF9C", "06402B"), (0.00, "CDEBD6", "06402B")]
    RED_BANDS   = [(-0.03, "F7D6D4", "7A1B17"), (-0.10, "ED9C97", "7A1B17"),
                   (-0.20, "DE6A63", _DARK_FONT), (None,  "B5261E", _DARK_FONT)]
    def _grad_rule(cond, bg, fg):
        return FormulaRule(formula=[cond],
                           fill=PatternFill(start_color=bg, end_color=bg, fill_type="solid"),
                           font=Font(name=FN, color=fg, bold=True), stopIfTrue=True)
    _guard = f'ISNUMBER(C{SC_R0}),C{TG_R0}<>""'
    # 綠(達標)：差距 >= 門檻，由深(0.20)到淺(0.00)
    for thr, bg, fg in GREEN_BANDS:
        wsY.conditional_formatting.add(score_rng, _grad_rule(
            f'AND({_guard},C{SC_R0}-C{TG_R0}>={thr})', bg, fg))
    # 紅(未達標)：由淺(>= -0.03)到深；最後一段為其餘(catch-all，最深)
    for thr, bg, fg in RED_BANDS:
        cond = (f'AND({_guard},C{SC_R0}-C{TG_R0}>={thr})' if thr is not None
                else f'AND({_guard},C{SC_R0}<C{TG_R0})')
        wsY.conditional_formatting.add(score_rng, _grad_rule(cond, bg, fg))
    wsY.cell(row=LG_HR + len(merged) + 2, column=2,
        value="提示：上表「平均分數」依「目標值」雙向漸層著色——達標(≥目標)綠、越高越深綠；"
              "未達標(<目標)紅、越低越深紅；無中間色(不出現黃色)。目標留空或樣本不足(「—」)者不著色，"
              "修改目標值即時更新色彩。").font = F(8.5, False, GREY)
    wsY.merge_cells(f"B{LG_HR + len(merged) + 2}:{last_col}{LG_HR + len(merged) + 2}")

    # ============================ 10. 使用說明 / 封面 ============================
    ws0 = wb.create_sheet("使用說明", 0)
    ws0.sheet_properties.tabColor = EVA_GOLD
    ws0.sheet_view.showGridLines = False
    ws0.column_dimensions["A"].width = 3
    ws0.column_dimensions["B"].width = 22
    ws0.column_dimensions["C"].width = 60
    ws0.column_dimensions["D"].width = 22
    ws0.merge_cells("B2:D3")
    ws0["B2"] = "長榮航空　旅客滿意度互動分析報告"
    ws0["B2"].font = F(20, True, WHITE); ws0["B2"].alignment = C
    for r in ws0["B2:D3"]:
        for cc in r: cc.fill = fill(EVA_GREEN)
    ws0.merge_cells("B4:D4")
    ws0["B4"] = f"資料期間：{PERIOD_TW}　│　運務管理部 客運運務課 (OMD/PMS)"
    ws0["B4"].font = F(11, True, WHITE); ws0["B4"].alignment = C
    for r in ws0["B4:D4"]:
        for cc in r: cc.fill = fill(EVA_MID)

    def section(r, title):
        ws0.merge_cells(f"B{r}:D{r}")
        c = ws0[f"B{r}"]; c.value = title; c.font = F(12, True, EVA_GREEN); c.fill = fill(EVA_LIGHT); c.alignment = L
        for rr in ws0[f"B{r}:D{r}"]:
            for cc in rr: cc.fill = fill(EVA_LIGHT)
        return r + 1

    r = 6
    r = section(r, "一、資料概況")
    overview = [
        ("資料期間", f"{PERIOD_TW}"),
        ("問卷回覆份數", f"{df1['問卷編號'].nunique():,} 份"),
        ("評分紀錄筆數", f"{n_sat:,} 筆"),
        ("3分以下原因填答", f"{n_rsn:,} 筆"),
        ("涵蓋服務構面", "報到櫃檯、貴賓室人員、貴賓室環境、登機服務、客艙清潔(共5項)"),
        ("涵蓋起飛城市", f"{sat['起飛城市'].nunique()} 個航站（歸併為 {sat['區域'].nunique()} 大區域）"),
        ("涵蓋班機編號", f"{sat['班機編號'].nunique()} 個班次"),
        ("涵蓋旅客國籍", f"{sat['國籍'].nunique()} 國"),
    ]
    for k, v in overview:
        ws0.cell(row=r, column=2, value=k).font = F(10.5, True)
        ws0.cell(row=r, column=2).fill = fill(EVA_LIGHT2); ws0.cell(row=r, column=2).border = thin_box
        ws0.cell(row=r, column=2).alignment = L
        cv = ws0.cell(row=r, column=3, value=v); cv.font = F(10.5); cv.alignment = L; cv.border = thin_box
        ws0.cell(row=r, column=4).border = thin_box
        r += 1

    r += 1
    r = section(r, "二、各分頁說明")
    sheets_desc = [
        ("互動查詢", "★主要工具：左側下拉選單篩選，右側即時顯示平均分數、回覆筆數、滿意占比、分數分布與待改善原因TOP10。"),
        ("年度目標達成", "各年度×服務項目加權平均(Σ分數/Σ人數)，回覆人數≤5不列計；含2026合併項目(臨櫃報到+登機、貴賓室服務)；達標綠/未達標紅。"),
        ("問卷項目總覽", "五大問卷項目之平均分數、回覆人數、滿意/待改善占比與最主要待改善原因一覽；右側可用本頁專屬篩選(含起飛城市、搭機年度/月份等)。"),
        ("交叉分析", "會員卡別 / 艙等 / 機型 / 區域 與問卷項目之平均分數交叉表，色階呈現相對高低；右側可用本頁專屬篩選(含區域)。"),
        ("城市排行", "各起飛城市(含所屬區域) × 問卷項目平均分數，支援標題列排序/篩選；右側可用本頁專屬篩選(含區域，可只看單一區域內各站)。"),
        ("原因解析", "3分以下待改善原因完整29項排名；右側可用本頁專屬篩選(含起飛城市、搭機年度/月份、構面等)，與互動查詢互不影響。"),
        ("滿意度明細 / 原因明細", "原始整理後資料(公式來源)，可自行樞紐分析；請勿刪除以免公式失效。"),
    ]
    for k, v in sheets_desc:
        ws0.cell(row=r, column=2, value=k).font = F(10.5, True, EVA_MID)
        ws0.cell(row=r, column=2).alignment = L; ws0.cell(row=r, column=2).border = thin_box
        ws0.cell(row=r, column=2).fill = fill(EVA_LIGHT2)
        cv = ws0.cell(row=r, column=3, value=v); cv.font = F(10); cv.alignment = L; cv.border = thin_box
        ws0.cell(row=r, column=4).border = thin_box
        r += 1

    r += 1
    r = section(r, "三、代碼對照")
    ws0.cell(row=r, column=2, value="類別").font = F(10.5, True, WHITE)
    ws0.cell(row=r, column=2).fill = fill(EVA_MID); ws0.cell(row=r, column=2).alignment = C; ws0.cell(row=r, column=2).border = thin_box
    ws0.merge_cells(f"C{r}:D{r}")
    ws0.cell(row=r, column=3, value="代碼說明").font = F(10.5, True, WHITE)
    ws0.cell(row=r, column=3).fill = fill(EVA_MID); ws0.cell(row=r, column=3).alignment = C; ws0.cell(row=r, column=3).border = thin_box
    ws0.cell(row=r, column=4).fill = fill(EVA_MID); ws0.cell(row=r, column=4).border = thin_box
    r += 1
    codes = [
        ("艙等", "商務艙(C)、豪華經濟艙(K)、經濟艙(Y)"),
        ("區域(依起飛城市)", "東北亞=日韓各站；東南亞=泰越星馬菲印柬等站；歐洲/澳洲=AMS/CDG/LHR/MUC/MXP/VIE/BNE；"
                           "北美=LAX/SFO/SEA/YVR/ORD/IAH/JFK/DFW/YYZ；台灣=TPE/TSA/KHH；中國/香港/澳門=CAN/PEK/PVG/SHA/SZX/XMN/HGH/TFU/HKG/MFM"),
        ("會員卡別", "CD / CE / CG / GC 為會員卡別代碼；NIL 為非會員或未提供 (實際對應請依公司定義)"),
        ("機型", "321(A321)、333(A330-300)、77A/77B/77M(B777系列)、781/78N/78P(B787系列)"),
        ("填答來源", "EDM(電子報問卷)、Website(官網問卷)"),
        ("分數", "1=非常不滿意，2=不滿意，3=普通，4=滿意，5=非常滿意"),
    ]
    for k, v in codes:
        ws0.cell(row=r, column=2, value=k).font = F(10.5, True)
        ws0.cell(row=r, column=2).alignment = L; ws0.cell(row=r, column=2).fill = fill(EVA_LIGHT2); ws0.cell(row=r, column=2).border = thin_box
        ws0.merge_cells(f"C{r}:D{r}")
        cv = ws0.cell(row=r, column=3, value=v); cv.font = F(9.5); cv.alignment = L; cv.border = thin_box
        ws0.cell(row=r, column=4).border = thin_box
        r += 1
    ws0.cell(row=r + 1, column=2,
        value="製表：本檔由 Python 自動產生，每月更換來源檔即可重跑。如需新增維度或指標歡迎再提出。").font = F(8.5, False, GREY)
    ws0.merge_cells(f"B{r+1}:D{r+1}")

    # ============================ 9. 專屬「分析圖表」頁籤(原生 Excel 圖表，集中於單一頁籤，隨篩選連動) ============================
    from openpyxl.chart import BarChart, DoughnutChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import DataPoint

    _PAL = ["004D33", "00875A", "2E9E72", "B7903B", "5FBE94", "86CF9C", "C9A227"]  # EVA 綠金系

    def _chart_base(ch, title, legend=True, legend_pos="b"):
        ch.title = title
        ch.height = 7.5; ch.width = 15
        if getattr(ch, "x_axis", None) is not None:
            ch.x_axis.delete = False; ch.x_axis.majorTickMark = "out"
        if getattr(ch, "y_axis", None) is not None:
            ch.y_axis.delete = False; ch.y_axis.majorTickMark = "out"
        if legend and ch.legend is not None:
            ch.legend.position = legend_pos; ch.legend.overlay = False
        else:
            ch.legend = None
        return ch

    def _series_colors(ch, colors):
        for i, s in enumerate(ch.series):
            c = colors[i % len(colors)]
            s.graphicalProperties.solidFill = c
            s.graphicalProperties.line.solidFill = c        # 線色(雷達/折線必需)
            s.graphicalProperties.line.width = 28000          # ~2.2pt

    def _value_labels(ch, num="0.00", pos="outEnd"):
        dl = DataLabelList(); dl.showVal = True; dl.numFmt = num
        if pos:
            dl.dLblPos = pos
        dl.showLegendKey = False; dl.showCatName = False; dl.showSerName = False
        dl.showPercent = False; dl.showBubbleSize = False
        ch.dataLabels = dl

    # 專屬「分析圖表」頁籤：卡片式儀表板(左側 KPI 數字卡 + 分數圓餅 + 各分析圖網格)
    from openpyxl.chart import PieChart
    import pandas as _pd
    wsC = wb.create_sheet("分析圖表")
    wsC.sheet_view.showGridLines = False
    try:
        wsC.sheet_properties.tabColor = "004D33"
    except Exception:
        pass
    for _c, _w in {"A": 7.5, "B": 7.5, "C": 7.5, "D": 2.5}.items():
        wsC.column_dimensions[_c].width = _w
    wsC.merge_cells("A1:U1")
    wsC["A1"] = "旅客滿意度　分析圖表"
    wsC["A1"].font = Font(name=FN, size=15, bold=True, color="FFFFFF")
    wsC["A1"].fill = PatternFill("solid", fgColor="004D33")
    wsC["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    wsC.row_dimensions[1].height = 30

    # ── 左側 KPI 數字卡(填寫人數 + 5 個分數等級) ──
    _scs = _pd.to_numeric(sat["分數"], errors="coerce").dropna().round().astype(int).value_counts().to_dict()
    _n = {k: int(_scs.get(k, 0)) for k in (5, 4, 3, 2, 1)}
    _n_resp = int(df1["問卷編號"].nunique())
    _cside = Side(style="thin", color="BFBFBF")
    _cardbox = Border(left=_cside, right=_cside, top=_cside, bottom=_cside)
    _cards = [("填寫人數", _n_resp, "004D33", "E7F2EC"),
              ("5分 非常滿意", _n[5], "1B7E3F", "EAF6EE"),
              ("4分 滿意", _n[4], "4CAE6A", "EFF8F1"),
              ("3分 普通", _n[3], "9C7B2E", "F7F1E4"),
              ("2分 不滿意", _n[2], "E08E45", "FBEEDF"),
              ("1分 非常不滿意", _n[1], "D9534F", "FBEAE9")]
    _r = 3
    for _lab, _val, _col, _bg in _cards:
        wsC.merge_cells(start_row=_r, start_column=1, end_row=_r, end_column=3)
        _lc = wsC.cell(row=_r, column=1, value=_lab)
        _lc.font = Font(name=FN, size=9, bold=True, color="404040")
        _lc.fill = PatternFill("solid", fgColor=_bg)
        _lc.alignment = Alignment(horizontal="center", vertical="center")
        wsC.merge_cells(start_row=_r + 1, start_column=1, end_row=_r + 1, end_column=3)
        _nc = wsC.cell(row=_r + 1, column=1, value=_val)
        _nc.font = Font(name=FN, size=20, bold=True, color=_col)
        _nc.fill = PatternFill("solid", fgColor="FFFFFF")
        _nc.alignment = Alignment(horizontal="center", vertical="center")
        _nc.number_format = "#,##0"
        for _rr in (_r, _r + 1):
            for _cc in (1, 2, 3):
                wsC.cell(row=_rr, column=_cc).border = _cardbox
        wsC.row_dimensions[_r].height = 16
        wsC.row_dimensions[_r + 1].height = 30
        _r += 2

    # ── 分數分布資料表(供圓餅參照；置於右側 V/W 欄，不隱藏以確保圖表能繪製) ──
    _pie_lab = {5: "5分 非常滿意", 4: "4分 滿意", 3: "3分 普通", 2: "2分 不滿意", 1: "1分 非常不滿意"}
    _gf = Font(name=FN, size=8, color="A6A6A6")
    wsC["V2"] = "分數"; wsC["W2"] = "筆數"
    wsC["V2"].font = _gf; wsC["W2"].font = _gf
    for _i, _k in enumerate((5, 4, 3, 2, 1)):
        _vc = wsC.cell(row=3 + _i, column=22, value=_pie_lab[_k]); _vc.font = _gf
        _wc = wsC.cell(row=3 + _i, column=23, value=_n[_k]); _wc.font = _gf
    wsC.column_dimensions["V"].width = 14
    wsC.column_dimensions["W"].width = 7

    # ── 分數分析(圓餅；5分→1分 綠至紅，含筆數標籤) ──
    cP = PieChart()
    cP.add_data(Reference(wsC, min_col=23, max_col=23, min_row=2, max_row=7), titles_from_data=True)
    cP.set_categories(Reference(wsC, min_col=22, max_col=22, min_row=3, max_row=7))
    _chart_base(cP, "分數分析（評分分布）", legend=True, legend_pos="r")
    cP.height = 8; cP.width = 11.5
    for _i, _col in enumerate(["1B7E3F", "4CAE6A", "B7903B", "E08E45", "D9534F"]):
        _dp = DataPoint(idx=_i); _dp.graphicalProperties.solidFill = _col
        cP.series[0].data_points.append(_dp)
    _dlp = DataLabelList(); _dlp.showVal = True; _dlp.showPercent = False; _dlp.numFmt = "#,##0"
    _dlp.showLegendKey = False; _dlp.showCatName = False; _dlp.showSerName = False
    cP.dataLabels = _dlp
    wsC.add_chart(cP, "E3")

    # ── 問卷項目總覽：各構面平均分數(直條) ──
    c1 = BarChart(); c1.type = "col"; c1.grouping = "clustered"
    c1.add_data(Reference(ws2, min_col=3, max_col=3, min_row=4, max_row=9), titles_from_data=True)
    c1.set_categories(Reference(ws2, min_col=2, max_col=2, min_row=5, max_row=9))
    _chart_base(c1, "各構面平均滿意度分數", legend=False)
    c1.y_axis.scaling.min = 3.5; c1.y_axis.scaling.max = 5; c1.y_axis.numFmt = "0.00"
    _series_colors(c1, ["00875A"]); _value_labels(c1, "0.00")
    wsC.add_chart(c1, "M3")

    # ── 交叉分析：艙等 × 構面(雷達圖；各艙等在 5 構面上的分數輪廓) ──
    from openpyxl.chart import RadarChart
    c2 = RadarChart(); c2.type = "marker"
    c2.add_data(Reference(ws3, min_col=2, max_col=7, min_row=12, max_row=14), titles_from_data=True, from_rows=True)
    c2.set_categories(Reference(ws3, min_col=3, max_col=7, min_row=11, max_row=11))
    _chart_base(c2, "艙等 × 服務構面　平均滿意度（雷達圖）", legend=True)
    c2.y_axis.scaling.min = 3.5; c2.y_axis.scaling.max = 5; c2.y_axis.numFmt = "0.00"
    c2.width = 16; c2.height = 11; _series_colors(c2, ["004D33", "B7903B", "00875A"])
    wsC.add_chart(c2, "E20")

    # ── 城市排行：各城市回覆人數(橫條，依城市數調整高度) ──
    c3 = BarChart(); c3.type = "bar"
    c3.add_data(Reference(ws4, min_col=4, max_col=4, min_row=4, max_row=last4), titles_from_data=True)
    c3.set_categories(Reference(ws4, min_col=2, max_col=2, min_row=5, max_row=last4))
    _chart_base(c3, "各起飛城市　總回覆人數", legend=False)
    c3.x_axis.scaling.orientation = "maxMin"      # 由上而下
    c3.height = max(8, min(42, 0.5 * len(cities) + 2)); c3.y_axis.numFmt = "0"
    _series_colors(c3, ["00875A"]); _value_labels(c3, "0", pos=None)
    wsC.add_chart(c3, "E83")

    # ── 原因解析：待改善原因 TOP10(橫條) ──
    c4 = BarChart(); c4.type = "bar"
    c4.add_data(Reference(ws5, min_col=4, max_col=4, min_row=5, max_row=15), titles_from_data=True)
    c4.set_categories(Reference(ws5, min_col=3, max_col=3, min_row=6, max_row=15))
    _chart_base(c4, "待改善原因 TOP 10（提及次數）", legend=False)
    c4.x_axis.scaling.orientation = "maxMin"; c4.height = 10; c4.width = 18; c4.y_axis.numFmt = "0"
    _series_colors(c4, ["D9534F"]); _value_labels(c4, "0", pos=None)
    wsC.add_chart(c4, "E62")

    # ── 年度目標達成：各項目 年度平均(群組直條，2024→2026 由淺至深) ──
    c5 = BarChart(); c5.type = "col"; c5.grouping = "clustered"
    c5.add_data(Reference(wsY, min_col=2, max_col=9, min_row=6, max_row=8), titles_from_data=True, from_rows=True)
    c5.set_categories(Reference(wsY, min_col=3, max_col=9, min_row=5, max_row=5))
    _chart_base(c5, "各服務項目　年度平均滿意度", legend=True)
    c5.y_axis.scaling.min = 3.5; c5.y_axis.scaling.max = 5; c5.y_axis.numFmt = "0.00"
    c5.width = 24; c5.height = 9; _series_colors(c5, ["A7E0C8", "2E9E72", "004D33"])
    wsC.add_chart(c5, "E42")

    # (分數分布已改為上方「分數分析」圓餅 cP)



    # 分頁順序
    # 移除 openpyxl 預設殘留的空白分頁「Sheet」(若存在且未使用)
    if "Sheet" in wb.sheetnames:
        try:
            del wb["Sheet"]
        except Exception:
            pass
    order = ["使用說明", "互動查詢", "分析圖表", "年度目標達成", "問卷項目總覽", "交叉分析", "城市排行", "原因解析", "滿意度明細", "原因明細", "清單", "目標參照"]
    wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)

    wb.active = wb.sheetnames.index("互動查詢")
    wb.save(OUTPUT_FILE)
    print("已儲存：", OUTPUT_FILE)



if __name__ == "__main__":
    main()
