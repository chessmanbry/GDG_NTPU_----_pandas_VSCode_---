# GDG NTPU｜跨部門報銷抓鬼遊戲：pandas 版（Windows／macOS）

**適用：**Python 初學者、VS Code、最後 40 分鐘實作課程。  
**課程任務：**用 pandas 讀取四份 Excel，合併成一份稽核名單，抓出相同發票號碼，將可疑紀錄標記並輸出 Excel 報告。  
**注意：**所有人物、部門及發票資料都是虛構教學資料；「重複發票」僅代表需人工核對，**不能單憑此結果指控有人重複請款或舞弊**。

## 1. 第一次使用：最短路徑

下載 ZIP 後，**先完整解壓縮**，再打開資料夾，不要直接在 ZIP 壓縮檔內雙擊程式。Windows、macOS 都需要先安裝 **Python 3.10 或更新版本**，並且安裝套件時需要網路。

### Windows 11／10

1. 到 https://www.python.org/downloads/ 安裝 Python 3.10+，安裝時勾選 `Add Python to PATH`（如果尚未安裝）。
2. 開啟已解壓縮資料夾，**雙擊 `setup_windows.bat`**，它會在本資料夾建立 `.venv`，然後安裝 pandas、openpyxl。
3. 安裝完成後，**雙擊 `run_windows.bat`**。它會執行分析及驗收。
4. 開啟 `outputs/跨部門報銷稽核報告.xlsx`，觀看「稽核明細」、「異常清單」、「部門摘要」三張工作表。

若學校的 Windows 安全政策不允許執行 `.bat`，改在 VS Code 的終端機輸入：

```powershell
.\setup_windows.bat
.\run_windows.bat
```

也可以完全不使用 `.bat`，在 VS Code PowerShell 中手動執行：

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe run_audit.py
.\.venv\Scripts\python.exe verify_results.py
```

> 如果 `py` 不存在但 `python` 可執行，將第一行的 `py -3` 改成 `python`。若有多個 Python 版本，請選 3.10 或更新版。

### macOS（Intel／Apple Silicon 都適用）

1. 安裝 Python 3.10+（若電腦原本就有可用的 `python3`，不用重裝）。
2. 將 ZIP 完整解壓縮，雙擊 **`setup_mac.command`** 安裝套件。
3. 雙擊 **`run_mac.command`** 執行分析與驗收。
4. 成品儲存在 `outputs/跨部門報銷稽核報告.xlsx`。

macOS 的檔案執行政策可能阻止剛下載的 `.command` 雙擊；遇到此情況，改在 VS Code 內的 Terminal 輸入下列指令（**在專案根目錄**）：

```bash
bash setup_mac.command
bash run_mac.command
```

或完全使用手動指令（無需雙擊）：

```bash
python3 -m venv .venv
./.venv/bin/python -m pip install -r requirements.txt
./.venv/bin/python run_audit.py
./.venv/bin/python verify_results.py
```

**注意：** `.command` 檔已設定為可執行；若解壓縮工具沒有保留權限，可在終端機輸入 `chmod +x setup_mac.command run_mac.command`。若系統因來源未受信任而阻擋，不必繞過安全政策，改用上面的 `bash` 指令執行檢查過內容的腳本即可。

## 2. VS Code 打開、執行 Python

1. 啟動 VS Code → **File → Open Folder**，選擇解壓縮出來的整個專案資料夾。
2. 建議安裝 Microsoft 提供的 **Python** 擴充功能（初次使用 VS Code 才需要）。
3. 使用 **Terminal → New Terminal** 開啟內建終端機。
4. 執行上述 `setup_*` 腳本，然後執行 `run_*` 腳本。或直接執行 `run_audit.py`。
5. 若要使用 VS Code 的 `Run Python File` 按鈕，先選擇本專案的 Python Interpreter：
   - Windows：`.venv\Scripts\python.exe`
   - macOS：`.venv/bin/python`

每個腳本都使用自己的所在資料夾定位假資料，因此**不依賴終端機的工作目錄**；但簡化操作時仍建議從專案根目錄下指令。

## 3. 專案資料夾與檔案用途

```text
GDG_NTPU_報銷抓鬼_pandas_VSCode_跨平台/
├── README.md                  本文件：操作方式、雙系統指令、完整程式流程
├── requirements.txt           需安裝的 pandas 和 openpyxl
├── setup_windows.bat          Windows：雙擊安裝套件（建立 .venv）
├── run_windows.bat            Windows：雙擊執行案例及驗收
├── setup_mac.command          macOS：雙擊安裝套件（建立 .venv）
├── run_mac.command            macOS：雙擊執行案例及驗收
├── run_audit.py               主要 pandas 實作程式（兩種系統共用）
├── generate_fake_data.py      重建虛構資料（原始四份 Excel 已附）
├── verify_results.py          驗收報表及各階段 CSV
├── data/
│   ├── 報銷_活動部.xlsx
│   ├── 報銷_器材部.xlsx
│   ├── 報銷_公關部.xlsx
│   └── 報銷_總務部.xlsx
└── outputs/                   執行後自動產生，不須事先建立
```

程式檔 `*.py` 本身**跨平台共用**；只有一鍵安裝與執行腳本依 Windows、macOS 分成不同檔名。安裝套件時建立的 `.venv` 不在 ZIP 裡，因為不同電腦需要各自建立相容環境。

## 4. 今天要學的 pandas 操作

| 步驟（終端機輸出代號） | Python 操作 | 預期成果 |
|---|---|---|
| G1 讀取與合併 | `pd.read_excel()`、`pd.concat()` | 4 份 Excel → 12 筆明細 |
| G2 找出重複 | `duplicated(subset=["發票號碼"], keep=False)` | 2 組發票號碼、4 筆相關紀錄 |
| G3 建立稽核摘要 | `.loc[...]`、`.groupby()`、`.agg()` | 保留 12 筆並新增「稽核狀態」、完成四部門統計 |
| G4 匯出資料 | `.to_csv()`、`pd.ExcelWriter()` | 四份中途 CSV＋一份多工作表 Excel |
| G5 重新驗證 | `pd.read_excel()`、資料框比較 | 三個工作表內容一致，印出 PASS |

**兩組重複發票**：`INV-2026-0881`、`INV-2026-0943`；兩個號碼各出現兩次，涉及 **4 筆**。  
**申請總額**：28,520 元，僅是提交的申請額合計，**並非實際損失或核准金額**。

### 讀者可直接觀察的中途成果

每次執行 `run_audit.py` 後，都會在 `outputs/` 出現：

- `01_合併總表.csv`：原始四份合併，12 筆。
- `02_疑似重複發票.csv`：2 組重複發票的 4 筆紀錄。
- `03_稽核明細.csv`：12 筆都保留，新增稽核狀態。
- `04_部門摘要.csv`：四個部門申請筆數、申請金額合計、待覆核筆數。
- `跨部門報銷稽核報告.xlsx`：三個工作表，方便行政人員翻閱。
- `稽核結果摘要.txt`：純文字結果，方便螢幕閱讀器讀取。

輸出會覆寫同名的**舊結果檔**，但**不會修改 `data/` 的四份原始假資料**。從頭重跑會重新建立資料框，不會把前一次的 12 筆往上累加。

## 5. 最後 40 分鐘建議進度

| 時間 | 學員任務 | 完成判定 |
|---|---|---|
| 0–5 分 | 解壓縮、安裝套件（課前先安裝比較穩） | `READY: pandas ... openpyxl ...` |
| 5–13 分 | G1 四份 Excel 合併 | 12 筆 |
| 13–22 分 | G2 找出兩組重複發票 | 4 筆涉及重複 |
| 22–30 分 | G3 標記待覆核、計算部門摘要 | 稽核明細 12 筆、部門 4 筆 |
| 30–37 分 | G4 匯出 CSV 與 Excel、G5 驗證 | PASS、報告 3 工作表 |
| 37–40 分 | 開啟成品、討論風險與資料倫理 | 不把「疑似」說成已證明舞弊 |

**課堂建議**：提前完成 Python／套件安裝；不建議把現場網路下載的等待時間算進 40 分鐘實作。

## 6. 假資料產生器（選做）

四份 Excel 已放在 `data/` 中，**不必先產生**。想確認假資料如何建立，可打開 `generate_fake_data.py` 閱讀 pandas 程式。

Windows：

```powershell
.\.venv\Scripts\python.exe generate_fake_data.py
```

macOS：

```bash
./.venv/bin/python generate_fake_data.py
```

預設發現既有檔案就不改動。確定要重建並**覆寫** `data/` 四份假資料，才加上 `--force`。例如：

```bash
./.venv/bin/python generate_fake_data.py --force
```

## 7. 常見問題

- **雙擊沒有反應／瞬間關閉**：改用 VS Code 終端機輸入 `setup_*` 或 `run_*`，便於閱讀錯誤訊息。
- **`python`、`py`、`python3` 找不到**：先安裝 Python 3.10+，Windows 勾選 PATH，安裝後重開 VS Code 終端機。
- **`No module named pandas`**：先執行 `setup_*`。在 VS Code 也要選 `.venv` 的 Python Interpreter，不要誤用其他 Python。
- **`No module named openpyxl`**：同上，`openpyxl` 是 pandas 讀寫 `.xlsx` 使用的引擎。
- **`FileNotFoundError`**：四份 Excel 必須放在 `data/` 子資料夾，不是主目錄或 `outputs/`。
- **Windows PowerShell 不允許執行 `.ps1`**：本教材提供的是 `.bat`，不需要修改 PowerShell ExecutionPolicy。
- **macOS 權限提示**：改用 `bash setup_mac.command` / `bash run_mac.command`，不必修改系統安全設定。
- **輸出檔案被 Excel 鎖住**：關閉已開啟的輸出報表，再重跑。
- **第一次安裝很慢**：下載套件依網速不同可能需要數分鐘；建議課前完成。

## 8. 最後練習（不需改大量程式）

1. 找出 `run_audit.py` 裡的 `keep=False`。如果改成 `keep="first"`，為何只剩兩筆標記？改完請恢復。
2. 為何不同部門的同號發票需要人工覆核，而不能直接 `drop_duplicates()`？
3. 試著比較 `outputs/01_合併總表.csv` 與 `outputs/03_稽核明細.csv`：哪個欄位是本次新增？
4. 為什麼部門申請總額不能直接稱作「已花費金額」？

完成這個案例的重點不是抓到壞人，而是理解**資料整合 → 定義規則 → 標記疑點 → 保存紀錄 → 產出可追溯報告 → 重新驗證**這條工作流程。
