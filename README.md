# EVA 旅客滿意度互動儀表板 — 桌面版使用與打包說明 (v2.1)

長榮航空 運航管理部 客運運務課
將問卷原始匯出檔（`DataExport*.xlsx`）轉成具「資料驗證下拉、依區域連動、交叉分析、原因解析」的互動 Excel 儀表板。

---

## 一、檔案清單（請全部放在同一資料夾）
| 檔案 | 必要 | 用途 |
|---|:--:|---|
| `EVA_儀表板GUI.py` | ✔ | 圖形介面前端（執行的主程式） |
| `build_dashboard.py` | ✔ | 建表引擎（GUI 內部直接呼叫，**檔名不可改**） |
| `EVA_Dashboard.spec` | 打包時 | PyInstaller 打包設定 |
| `eva.ico` | 打包時 | 產生 .exe 的圖示（地球圖標） |
| `build_exe.bat` | 打包時 | 一鍵打包成單機免安裝 .exe |
| `eva_logo_full.ico` | 選用 | 備選圖示（整條 logo 版） |

> 引擎與 GUI 必須同資料夾；打包時連同 `eva.ico` 一起放齊。

---

## 二、直接以 Python 執行（開發／測試）
```
python EVA_儀表板GUI.py
```
（認不得 `python` 時改用 `py EVA_儀表板GUI.py`）

需先安裝套件（僅這台）：
```
pip install pandas openpyxl pillow
```

---

## 三、介面操作
本版為「**純自選檔案**」模式，**全程不複製檔案、不建立任何資料夾**。

1. **① 選取要匯入的資料檔**
   - 按「＋ 新增檔案…」挑選一個或多個 `DataExport*.xlsx`（清單只記住檔案路徑）。
   - 可「移除選取／清空」。同一路徑不會重複加入。
2. **② 匯出設定**
   - 「選擇…」設定匯出路徑與檔名。
   - 產生模式：
     - **FULL（重建所有頁籤）**：以所選檔重建整份儀表板（目標值取程式內建設定）。
     - **DATA_ONLY（僅更新資料、保留人工維護）**：把所選檔的**新問卷**增量併入**既有匯出檔**，
       重複問卷自動略過，並保留分析頁的人工編修（須匯出檔已存在）。
3. **產生儀表板**：底部進度條顯示百分比與流光動畫；下方有狀態列與詳細記錄。
   完成後可選擇開啟匯出檔所在資料夾。

### 增量更新（DATA_ONLY）說明
- 以「匯出檔的明細」為累積基準：讀取既有明細的問卷編號 → 只把**新問卷**接在明細下方 → 重複者忽略。
- 下拉分類（年月、國籍…）依「累積後」資料刷新。
- **注意**：依區域連動的下拉（城市／班機／機型…）其「選項清單」於 **FULL（重建）** 時更新；
  若新月份帶進某區域的新城市／機型，請偶爾以 FULL 重建一次。

---

## 四、打包成單機免安裝 .exe
把 `EVA_儀表板GUI.py`、`build_dashboard.py`、`EVA_Dashboard.spec`、`eva.ico`、`build_exe.bat` 放同一資料夾，**雙擊 `build_exe.bat`**。
或手動執行（請用「模組」形式，避免 PATH 問題）：
```
python -m PyInstaller --noconfirm EVA_Dashboard.spec
```
完成後於 `dist\EVA滿意度儀表板.exe` 取得執行檔，可複製到任一台 Windows 免安裝直接雙擊。

### 更換圖示
- **.exe 圖示**：只吃 `.ico`。換圖時把新 `.ico` 放同資料夾，改 `EVA_Dashboard.spec` 內 `icon='你的檔.ico'`。
- **視窗／工作列圖示**：程式內已用內嵌地球 PNG（`EVA_ICON_B64`）並設定 Windows AppUserModelID，與 exe 圖示一致；
  表頭右側另顯示整條去背 EVA logo。

### 常見訊息
- 打包過程出現 `No module named 'pytest'` 等 **WARNING** 屬正常（測試模組未納入），不影響成品。
- `pyinstaller 不是內部或外部命令`：請改用 `python -m PyInstaller`。

---

## 五、進階：命令列直接產生（免介面）
引擎支援三種輸入來源（擇一）：
```
REM (1) 指定單一檔
python build_dashboard.py 輸入檔.xlsx [輸出檔.xlsx]

REM (2) 以環境變數指定多檔（路徑用 ; 分隔；Windows）
set EVA_INPUT_FILES=C:\a\DataExportAPR2026.xlsx;C:\a\DataExportMAY2026.xlsx
set EVA_OUTPUT_FILE=C:\out\儀表板.xlsx
python build_dashboard.py
```
其他環境變數：`EVA_UPDATE_MODE=FULL|DATA_ONLY`、`--rebuild`（強制重建）。

---

## 六、輸出內容（10 個分頁）
使用說明、互動查詢、年度目標達成、問卷項目總覽、交叉分析、城市排行、原因解析、滿意度明細、原因明細、清單（隱藏）。

品牌樣式：EVA 綠 `#004D33`、Microsoft JhengHei、壹貳參肆 章節編號。
公式採整欄固定範圍（列數增減自動連動），全程一般函式（無動態陣列），各版本 Excel 開啟不會出現修復訊息。
