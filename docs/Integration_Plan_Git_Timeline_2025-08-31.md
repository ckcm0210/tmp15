# Git 時間線整合到 Watchdog 工具 — 設計與落地計劃（2025-08-31）

本文檔描述：
- 現況：值/公式的保存形式與命名位置
- 要達成的目標：完整時間線（從最初到現在）、作者與時間、瀏覽器可視化
- 落地計劃：最小侵入、可維護的模組化設計
- 長期擴充：摘要 JSON、SQLite 事件索引、Viewer 時間線

---

## 1) 現況：值與公式如何保存（polars / polars_xml）
- 值引擎輸出（polars / polars_xml / xml）：
  - 結構：`{ sheet_name: { 'A1': 值, 'B2': 值, ... } }`（跳過空白，保留原型別）
- 與 openpyxl 公式合併（core/excel_parser.dump_excel_cells_with_timeout）：
  - 產出：`result = { sheet_name: { address: { "formula": fstr, "value": v, "cached_value": v } } }`
  - 當值引擎為 ('polars','polars_xml','xml') 時，會把值也寫入 `cached_value`，避免再跑 data_only pass。
- 基準線（Baseline）實體檔：
  - 路徑：`<LOG_FOLDER>/<base_key>.baseline.json.<ext>`（ext= .lz4/.zst/.gz）
  - base_key：以原始檔名 + 路徑 hash8；確保同名不同路徑不衝突。
  - 內容：`cells` 即上述結構（每格 formula/value/cached_value）。
- polars 合併 CSV（診斷可選）：
  - 路徑：`<CACHE_FOLDER>/values/<base_key>.values.csv`
  - 欄位：`sheet,address,value`（一個 workbook 一份，便於抽查）。
- 歷史快照（LOG_FOLDER / Git）：
  - LOG_FOLDER 壓縮快照：`<LOG_FOLDER>/history/<base_key>/<timestamp>.cells.json.<ext>` + `index.csv.gz`
  - Git 純 JSON（利於 diff）：`excel_git_repo/history/<base_key>/<timestamp>.cells.json`

---

## 2) 目標：完整時間線 + 作者/時間 + 瀏覽器可視化
- 每次「非靜默且有變更」比較後：
  1) 保存**全量 JSON**快照（LOG_FOLDER），包含每格公式/值/時間/作者與統計；
  2) 同步一份**純 JSON**至 Git repo 並自動 commit（commit author/時間對齊 Excel last_author / 事件時間）。
- Viewer（Flask）：
  - /history：該檔所有版本的清單（由最舊到最新）
  - /diff：選兩個版本做地址級差異，提供摘要與「只顯示有意義變更」
  - /timeline（擴充）：一年內全部事件＋摘要＋快速比對

---

## 3) 落地設計：最小侵入、可維護
- 新增與集中模組：
  - `utils/history.py`：快照/摘要/同步至 Git（已存在）
  - `web/` 或 `tools/viewer/`：放置 `git_viewer.py` 與資源（之後從根目錄搬遷過去）
  - `utils/git_sync.py`（可選）：若未來需要抽離 Git 同步邏輯
- 薄 Hook：
  - `core/comparison.py`（已接好）在非靜默且有變更時呼叫 `utils.history.*`
  - `config.settings`/`ui.settings_ui` 新增開關，不影響舊功能
- 不破壞相容：
  - baseline 保持原狀；CSV 仍然產生；只新增歷史快照/摘要

---

## 4) 欄位與命名（時間與作者）
- 欄位：
  - `event_time`：比較執行時間（ISO 格式）
  - `excel_mtime`：來源檔 `mtime`（秒）
  - `source_size`：來源檔大小（bytes）
  - `last_author`：Excel `docProps/core.xml` 的 lastModifiedBy
  - `summary`：變更數、按類型統計
  - `meaningful_breakdown`：DIRECT_VALUE_CHANGE / FORMULA_CHANGE_INTERNAL / EXTERNAL_REF_LINK_CHANGE / EXTERNAL_REFRESH_UPDATE / CELL_ADDED / CELL_DELETED
- 命名：
  - LOG_FOLDER 快照：`history/<base_key>/<YYYYMMDD_HHMMSS_mmmmmm>.cells.json.<ext>`
  - Git 純 JSON：`excel_git_repo/history/<base_key>/<YYYYMMDD_HHMMSS_mmmmmm>.cells.json`
  - Git 摘要 JSON（可選）：`excel_git_repo/history/<base_key>/diffs/<YYYYMMDD_HHMMSS_mmmmmm>.summary.json`

---

## 5) Viewer 擴充（已做 & 待做）
- 已做：
  - /history：版本選擇器（A/B + 只顯示有意義變更）
  - /diff：摘要（總數/分類）＋只顯示變化地址；xlsx → 自動引導至 .cells.json
- 待做：
  - /diff：分類對齊主工具（完整六大類）＋分類/工作表/關鍵字篩選
  - /history：一鍵「最新 vs 上一版」
  - /timeline：一年內全部事件清單（含摘要/作者/時間）

---

## 6) SQLite 事件索引（輕量級擴充，非必需但建議）
- 目的：快速列出 timeline、支援多條件篩選而不用每次掃 JSON
- Schema（草案）：
  - `events(id INTEGER PK, base_key TEXT, file_path TEXT, event_time TEXT, excel_mtime REAL, source_size INTEGER, last_author TEXT, git_commit_sha TEXT, snapshot_path TEXT, summary_path TEXT, total_changes INTEGER, dvc INTEGER, fci INTEGER, xrlc INTEGER, xru INTEGER, addc INTEGER, delc INTEGER)`
  - 索引：`idx_events_basekey_time (base_key, event_time DESC)`
- 接入點：
  - 在 `utils.history.sync_history_to_git_repo` 成功後寫入一條記錄
  - Viewer `/timeline` 直接查 SQLite 顯示
- 導入影響：
  - 對現有流程幾乎零影響；只是多寫一條索引記錄
  - SQLite 檔可放在 `LOG_FOLDER/events.sqlite`（易備份與攜帶）

---

## 7) 設定（config.settings）建議新增
- `ENABLE_HISTORY_SNAPSHOT = True`
- `HISTORY_GIT_REPO_PATH = r"C:\\rovo\\watchdog_1\\excel_git_repo"`
- `HISTORY_SYNC_FULL = True`
- `HISTORY_SYNC_SUMMARY = True`
- `HISTORY_RETENTION_DAYS = 90`（或 0 表示不清理）
- `HISTORY_MAX_VERSIONS = 500`（或 0 表示不限制）
- `HISTORY_GIT_AUTHOR_FROM_EXCEL = True`

---

## 8) 落地步驟（三階段）
- M1（已完成 + 補強）：
  - 全量 JSON 快照 + Git 純 JSON，同步與自動 commit；/history + /diff（MVP）
  - 補 event_time/excel_mtime/summary/breakdown；設定參數；/diff 完整分類
- M2（擴充）：
  - /timeline；摘要 JSON 同步；Viewer 支援直接讀摘要；首頁/歷史頁「一鍵比較」
- M3（治理）：
  - 保留策略（days / versions）；UI 選項；搬遷 viewer 至 web/；docs/Change_Log…

---

## 9) 啟動 Viewer 的方式
- 預設不自動開瀏覽器（避免干擾監控），提供：
  - 工具列/設定 UI 一鍵打開；
  - 在 Console 印出 Viewer 的 URL；
  - 可選門檻（例如變更數 > N ）時自動提示用戶開啟。

---

（完）
