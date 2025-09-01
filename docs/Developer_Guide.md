# Excel Watchdog 開發者指南與未來規劃

本文檔記錄了 Excel Watchdog 專案的進階開發概念、未來可行的功能擴展方向以及核心的軟體工程理念，旨在為未來的維護和功能開發提供清晰的指引。

## 1. 核心開發理念：自動化測試 (Pytest)

### 1.1 為什麼要用 Pytest？

`pytest` 是我們專案的**品質保證基石**和**安全網**。它的核心價值在於：

1.  **防止「迴歸」(Preventing Regressions)**: 確保新的修改不會意外破壞現有的、正常運作的功能。
2.  **賦予重構的信心 (Confidence to Refactor)**: 讓開發者可以大膽地改善程式碼內部結構，而不用擔心改變其外部行為。
3.  **充當「活文件」(Living Documentation)**: 測試案例用實際的程式碼展示了每個函數的預期行為，是永遠不會過時的技術文件。

### 1.2 如何運作？

- **測試檔案**: 所有測試都存放在根目錄的 `tests/` 資料夾中，以 `test_*.py` 命名。
- **執行測試**: 在專案根目錄 `watchdog_1/` 下，執行以下指令：
  ```bash
  python -m pytest
  ```
- **開發流程**: 每當你修改了任何核心邏輯，或新增了功能後，都應該執行一次測試。如果所有測試都通過 (`passed`)，你就可以很有信心地認為你的修改是安全的。

---

## 2. 未來功能擴展藍圖

### 2.1 處理加密的 Excel 檔案

**問題**: 目前無法讀取受密碼保護的 Excel 檔案。

**解決方案**: 對於已知且數量有限的密碼，採用「逐一嘗試」的策略。

**實施步驟**:
1.  在 `config/settings.py` 中新增一個密碼列表：
    ```python
    # 警告: 純文字密碼存在安全風險，請確保環境安全
    EXCEL_PASSWORDS = ['pass1', 'pass2', ...]
    ```
2.  修改 `core/excel_parser.py` 中的 `safe_load_workbook` 函數。
3.  在該函數中，先嘗試無密碼開啟。如果失敗，則遍歷 `EXCEL_PASSWORDS` 列表，使用 `openpyxl.load_workbook(..., password=pwd)` 逐個嘗試，直到成功為止。如果所有密碼都失敗，則記錄錯誤並放棄處理該檔案。

### 2.2 與版本控制系統整合 (Git Integration)

**問題**: 目前的 baseline 只記錄了「上一個」狀態，無法追溯更早的歷史版本。

**解決方案**: 使用 `GitPython` 函式庫，在偵測到變更後，自動將檔案的最新版本提交到一個 Git 倉庫中，建立永久的變更歷史。

**實施步驟**:
1.  安裝函式庫: `pip install GitPython`。
2.  在 `config/settings.py` 中設定 Git 倉庫的路徑 `GIT_REPO_PATH`。
3.  建立一個新的輔助函數 (例如在 `utils/git_handler.py` 中)，使用 `git.Repo` 來開啟倉庫，並執行 `repo.index.add()` 和 `repo.index.commit()`。
4.  在 `core/comparison.py` 中，當 `AUTO_UPDATE_BASELINE_AFTER_COMPARE` 成功執行後，呼叫這個新的 Git 提交函數，將變更歸檔。

**重要概念**: 此功能**不是用來取代**即時的 `baseline` vs `current` 比較，而是作為其**後續的歸檔步驟**。你的比較邏輯是「顯微鏡」，Git 是「時光機」。

### 2.3 使用 SQLite 提高擴展性

**問題**: 目前的檔案狀態 (`watcher.py`) 和日誌去重簽名 (`comparison.py`) 都儲存在記憶體中。當監控的檔案數量極大時，會消耗大量記憶體，且程式重啟後狀態會遺失。

**解決方案**: 使用 Python 內建的 `sqlite3` 模組，將這些狀態持久化到一個硬碟上的資料庫檔案中。

**實施步驟**:
1.  建立一個 `utils/database.py` 模組。
2.  在其中定義資料庫檔案的路徑 (例如 `LOG_FOLDER/watchdog_state.db`)。
3.  編寫 `init_db()` 函數，使用 `CREATE TABLE IF NOT EXISTS` 建立儲存狀態的表格 (例如 `file_states`)。
4.  編寫 `get_file_state()` 和 `update_file_state()` 等函數，用來讀寫資料庫。
5.  在 `watcher.py` 和 `comparison.py` 中，將對記憶體字典的操作，替換為對上述資料庫函數的呼叫。

**好處**:
- **持久化**: 狀態不會因程式重啟而遺失。
- **擴展性**: 記憶體佔用極低，可以支援監控海量檔案。
- **健壯性**: 資料庫交易比記憶體操作更安全。
