# 比較邏輯與高效值讀取引擎設計（更新版）

更新時間：2025-08-30
負責：Rovo Dev（AI 助手）

---

## 1. 設計目標（按使用者最新要求）

- 公式字串有改變，一律顯示（若涉及外部參照，特別標記為「外部連結變更」），即使結果巧合相同也要顯示。
- 公式字串沒變，但存在外部參照且結果變了（典型為 Refresh），必須顯示前後值（「外部 refresh 變更」）。
- 內部連鎖（非外部）造成的結果改變，預設不顯示；只保留「來源」格變化，避免噪音。
- 可同時兼顧「直接值變更」與「公式變更」。

關鍵術語：
- formula：公式字串（可 prettify 外部參照）
- display_value：值比對基準。若有 cached_value 則以 cached_value 為準，否則用 value

---

## 2. 變更分類（最新版）

- CELL_ADDED / CELL_DELETED
- DIRECT_VALUE_CHANGE
  - 兩邊都不是公式（old_formula/new_formula 為 None），display_value 改變
- FORMULA_CHANGE_INTERNAL
  - 公式字串變了，且不涉及外部參照（例：=A1+B1 → =A2+B2）
- EXTERNAL_REF_LINK_CHANGE（外部連結變更）
  - 公式字串變了，且涉及外部參照（路徑/檔名/工作表等）
  - 規則：一定顯示，無論結果是否相同（符合使用者：換連結即視為有意義改動）
- EXTERNAL_REFRESH_UPDATE（外部 refresh 變更）
  - 公式字串不變，存在外部參照，display_value 改變（代表被刷新）
- INDIRECT_CHANGE_INTERNAL（內部連鎖變更）
  - 公式字串不變，且不涉及外部參照，但 display_value 改變（因上游內部格改動）
  - 規則：預設不顯示（IGNORE_INDIRECT_CHANGES=True）
- NO_CHANGE

推薦設定（可同時滿足所有情境）：
- FORMULA_ONLY_MODE=False
- TRACK_DIRECT_VALUE_CHANGES=True
- TRACK_FORMULA_CHANGES=True
- TRACK_EXTERNAL_REFERENCES=True
- ENABLE_FORMULA_VALUE_CHECK=True（但採取更快的值讀取引擎，見下）
- IGNORE_INDIRECT_CHANGES=True

---

## 3. 值讀取引擎的兩條高效方案

本設計支援兩種高效率值（結果）讀取引擎，並允許同時保留 openpyxl/（或 XML）讀公式的管道，最後按 address 合併。

### 3.1 方案 2：純 XML 直讀（值與公式）

- 作法：
  - 使用 zipfile + XML iterparse 解析 .xlsx 結構，逐 sheet 讀取 `xl/worksheets/sheetN.xml`
  - 每個 `<c r="A1" t="s">` 表示單元格，`r` 是地址；子節點 `<f>` 是公式字串，`<v>` 是 cached 值
  - t="s" 時 `<v>` 是 sharedStrings 索引；數字/布林直接是最終值；（日期可先當數值處理，視需要增補 styles.xml）
- 優點：
  - 單 pass 即可掃到整張表的 address / formula / cached，速度快、無需第二次開檔
- 是否知道 cell 對應的值？
  - 是的，`<c r="A1">` 就是地址映射，能準確對應每格的值與公式
- 粗略效能估算（視 CPU/I/O）：
  - 10k 公式格：~0.2–0.6 秒
  - 100k 公式格：~1.5–3 秒

### 3.2 方案 3：Polars 管線（值）+ openpyxl 或 XML（公式）

- 值管線：
  - 用 xlsx2csv 將每個 worksheet 轉成 CSV（可使用記憶體中的 BytesIO，避免落地）
  - 用 Polars 讀取 CSV（`pl.read_csv` 或 `scan_csv`），再將寬表轉為長表（address → value）
- 公式管線：
  - 初期沿用 openpyxl 讀公式（read_only=True）；後期可改為 XML 讀公式以提速
- 合併：
  - 以 address 為鍵 join，得到 address → {formula, display_value}
- CSV 會不會產生實體檔？
  - 可選：預設用 BytesIO（不落地）；若要除錯/留存，可落地到 cache_folder。
- 粗略效能估算（單張表，視大小）：
  - 10k–100k cells：xlsx2csv ~0.1–1.0 s；Polars 讀取 ~0.1–1.0 s（總計多為 1 秒級）
  - 100 萬 cells：xlsx2csv ~2–5 s；Polars 讀取 ~2–6 s
  - 多張表可並行處理，縮短總時間

---

## 4. Address 對齊與空白儲存格處理

- Address 產生：
  - XML 路徑：`<c r="A1">` 已自帶地址
  - Polars 路徑：
    - CSV 為寬表（第 1 列、第 1 欄…），需將列/欄索引轉為 Excel 地址（例如 col=1→A、row=1→1 → A1）
    - 然後做寬轉長，得到 address → value（可同時記錄 sheet 名稱）
- 空白儲存格：
  - openpyxl 現行流程：只會記錄「有內容（公式或值）」的 cell；空白 cell 不記錄
  - Polars 路徑：CSV 通常會以逗號分隔代表空白；在寬轉長時可「跳過空字串/Null」，與 openpyxl 行為一致
  - 空白 worksheet：
    - 轉 CSV → 0 行有效資料；Polars 會輸出 0 筆記錄；與 openpyxl 一致（不會記錄任何 cell）
- 總結：兩條管線都會對「空白儲存格」保持一致：不產生記錄；因此 address 數量層面是可對齊的（只要兩邊的「非空定義」一致）

---

## 5. 舉例（摘要）

- 直接值變更（非公式）：
  - B2: 'test' → 'text'、C3: 123 → 124 → DIRECT_VALUE_CHANGE
- 內部公式改動：
  - D4: =A1+B1 → =A2+B2 → FORMULA_CHANGE_INTERNAL
- 外部連結變更（值同也顯示）：
  - E5: ='C:\a\[Book1.xlsx]Sheet1'!A1 → ='\\s\share\[Book2.xlsx]Sheet2'!B2 → EXTERNAL_REF_LINK_CHANGE
- 上游內部值變導致下游結果改：
  - A1: 10 → 12（顯示）
  - C1: =A1+B1（結果跟著變）→ INDIRECT_CHANGE_INTERNAL（預設忽略）
- 外部 refresh（公式不變，值變）：
  - F6: ='\\s\share\[Rates.xlsx]FX'!B2（公式不變）
  - F6 值 7.8000 → 7.8235 → EXTERNAL_REFRESH_UPDATE

---

## 6. UI 與引擎選擇策略

- 新增設定：
  - VALUE_ENGINE: 'polars'（預設）| 'xml'
  - FORMULA_ENGINE: 'openpyxl'（預設）| 'xml'
  - CSV_PERSIST: False（預設；True 時落地到 cache_folder）
  - MAX_SHEET_WORKERS: 併發數（預設=CPU 核心數）
- 自動回退：
  - 當選擇 'polars' 但環境未安裝 polars/xlsx2csv 時，自動回退至 'xml'，並在 console 提示

---

## 7. 實作里程碑（建議）

- M1：完成 XML 值解析（方案 2），沿用 openpyxl 讀公式 → 先根治慢讀 cached 的問題
- M2：加入 Polars 值管線（方案 3），提供 VALUE_ENGINE 切換，預設 'polars'、自動 fallback 'xml'
- M3：視需求將公式讀取也改為 XML（FORMULA_ENGINE='xml'），再做端到端效能優化

---

## 8. 常見問答（針對 Polars/CSV 與空白格）

Q：Polars 會如何處理空白儲存格？會輸出嗎？

A：CSV 中的空格通常以空字串/缺失值呈現。我們在寬轉長時會「跳過空字串/Null」，因此不會為空白 cell 產生記錄，與 openpyxl 現行行為一致。

Q：讀完 value 後，是否每格都有記錄其 cell address？

A：會。XML 直接有 `<c r="A1">`；Polars 路徑則由寬表索引推導出 address（A1、B2…），並在長表中填入 address 欄。

Q：openpyxl 讀公式也會記錄 address 嗎？

A：會。openpyxl 的 cell.coordinate 即地址（例如 'A1'），我們會保留該欄位，用 address 作為 join key 與值合併。

Q：兩邊讀到的「cell 數量」會一樣嗎？

A：一致的前提是「同一個非空定義」。現狀下：
- openpyxl：只記錄「有內容（公式或值）」的 cell
- Polars：在長表化過程中跳過空字串/Null
因此兩者都不為空白 cell 產生記錄，能對齊。

---

## 9. 效能估算（再次彙整）

- XML 解析（值/公式）
  - 10k 公式格：~0.2–0.6 秒；100k：~1.5–3 秒
- Polars（值）+ openpyxl（公式）
  - 10k–100k cells：總體多為 1 秒級
  - 100 萬 cells：數秒（依 I/O 與 CPU），可用併發分攤

---

## 10. 下一步待確認

- 是否同時推進 方案 2 + 方案 3（預設 'polars'，自動 fallback 'xml'）？
- 是否需要 CSV 落地（CSV_PERSIST=True）以利除錯/追查？
- 是否允許併發（MAX_SHEET_WORKERS），建議值？

