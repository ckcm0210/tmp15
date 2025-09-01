# Robocopy Deep Reading Notes and End-to-End Flow (2025-08-29)

This document records a detailed, line-by-line understanding of how the Watchdog copies Excel files to a local cache, where and how Robocopy is invoked, what "silent" vs "non-silent" means in the compare pipeline, and how results propagate. It also includes a human-friendly reference for Robocopy exit codes and practical guidance on the "stability window" during the save tail.

---

## 1) End-to-end flow: from event to copy to compare

- Entry: `main.py`
  - Initializes logging, launches Settings UI (values are saved to `config/runtime_settings.json` and immediately applied via `config.runtime.apply_to_settings`).
  - Schedules both `WATCH_FOLDERS` and `MONITOR_ONLY_FOLDERS`. Creates baselines if startup scan is enabled.

- File event handling: `core/watcher.py::ExcelFileEventHandler.on_modified`
  - Debounce (skip repeated events within `DEBOUNCE_INTERVAL_SEC`).
  - Ignore cache/log folders and temporary lock files (`~$` prefix) when `SKIP_WHEN_TEMP_LOCK_PRESENT` is enabled.
  - Preview compare (silent) to reduce noise:
    - Calls `core.comparison.compare_excel_changes(file_path, silent=True, is_polling=False)` to compute if there are changes without printing tables.
    - If preview detects changes, prints a banner and proceeds.
  - For monitor-only roots: on first change, logs info + creates baseline, then returns (no compare yet). Subsequent changes go through normal compare.
  - Performs an immediate visible compare (`silent=False`) and then starts adaptive polling to wait for stability and catch tail-end changes.

- Compare: `core/comparison.py::compare_excel_changes`
  - Optional fast skip: If baseline stores `source_mtime` + `source_size` and `QUICK_SKIP_BY_STAT` is enabled, and current stat matches baseline (mtime within tolerance + equal size), skip reading entirely.
  - Otherwise, it calls `core.excel_parser.dump_excel_cells_with_timeout(file_path, silent=True)` to read the workbook content.
  - If current content equals baseline content: return False (no visible output if polling; an informational line appears in polling mode).
  - If content differs and this is a non-silent call (visible compare), it renders the aligned diff table and optionally logs CSV records of “meaningful changes.” When enabled, it also updates the baseline with the latest content and current `source_mtime`/`source_size`.

- Reading Excel content: `core/excel_parser.py::dump_excel_cells_with_timeout`
  - Always obtains a local cached copy first by calling `utils.cache.copy_to_cache(path)`.
  - Reads data from the cached file using `openpyxl` in `read_only=True` mode, extracting formula/value pairs for each non-empty cell. It prettifies formulas (restoring human-readable external paths) and can optionally read cached values for formulas (value-check mode).
  - No heavy reads touch the original file; everything is performed on the cached copy to avoid file locks.

- Copy to cache (central piece): `utils/cache.py::copy_to_cache`
  - Ensures cache folder exists and normalizes the cache filename (safe, capped-length).
  - If a fresh-enough cached copy already exists (mtime >= source), it reuses the cache.
  - Before copying, runs a “stability window” check (see Section 4 below): `_wait_for_stable_mtime` compares `mtime` (and size) repeatedly for `COPY_STABILITY_CHECKS` times, spaced by `COPY_STABILITY_INTERVAL_SEC`, with an overall timeout `COPY_STABILITY_MAX_WAIT_SEC`.
  - Copy engines (configurable):
    - If `COPY_ENGINE` is `'robocopy'` or `'powershell'`, it uses a child process via `_run_subprocess_copy`.
    - Else, if `PREFER_SUBPROCESS_FOR_XLSM` and the file ends with `.xlsm`, it uses `SUBPROCESS_ENGINE_FOR_XLSM`.
    - Otherwise, it uses Python’s internal copy (`shutil.copy2` or chunked copy if `COPY_CHUNK_SIZE_MB > 0`).
  - On success, it waits `COPY_POST_SLEEP_SEC` seconds for filesystem stabilization, records a success ops log (engine is recorded), and returns the cache path.
  - On repeated failure, in strict mode (`STRICT_NO_ORIGINAL_READ=True`), it returns `None` (caller skips reading). In non-strict mode, it falls back to the original path (not recommended when avoiding locks).

---

## 2) Where Robocopy is actually invoked

- Function: `utils/cache.py::_run_subprocess_copy(src, dst, engine='robocopy')`
  - It constructs the Robocopy command:
    ```bat
    robocopy "<SRC_DIR>" "<DST_DIR>" "<FILENAME>" /COPY:DAT /R:2 /W:1 /NJH /NJS /NFL /NDL /NP /J
    ```
    - SRC_DIR = `os.path.dirname(src)`
    - DST_DIR = `os.path.dirname(dst)`
    - FILENAME = `os.path.basename(src)`
  - Return codes 0–7 are considered success (per Microsoft semantics); >7 is failure.
  - After a “successful” return code, the function performs a post-copy validation to ensure the destination file exists and size is not smaller than the source (and a tolerant `mtime` check). If validation fails, it raises `OSError` to trigger retry in `copy_to_cache`.
  - When `SHOW_DEBUG_MESSAGES=True`, it prints the exact command and the return code to help diagnosis.

- Important distinction (process model):
  - Robocopy is run in a separate process (a child of the Python process). It does not spin up a new kernel; it uses the same Windows kernel. The benefit of a child process is that when the process exits, the OS guarantees all file handles owned by that process are closed, avoiding handle-leak or lingering-lock issues within your main Python process.

---

## 3) "Silent" vs "Non-silent" in the compare pipeline

- "Silent" (`silent=True`):
  - No visible table output in console.
  - Used for preview checks (to decide whether to show a banner) and retried reads. It still performs the work (copy to cache, read workbook) but suppresses user-facing prints, except for minimal info logs.
  - In polling, the silent mode is used to avoid flooding the console until stability is confirmed.

- "Non-silent" (`silent=False`):
  - Produces visible, aligned diff tables.
  - Triggers CSV logging of “meaningful changes.”
  - Can auto-update baseline (`AUTO_UPDATE_BASELINE_AFTER_COMPARE=True`).

In the watcher: first call is silent (preview), then if changes are present, a visible (non-silent) compare runs immediately.

---

## 4) The "stability window" during the save tail

- What is the "save tail"?
  - When Excel saves (especially on network or OneDrive/SharePoint), metadata like `mtime` can update before the file is fully flushed/synchronized. Also, the presence of temporary lock files (`~$`) indicates Excel is still in the process of saving.
  - The “tail” refers to the last phase of the save, where data is still being written or synchronized, and file content or size may still be changing.

- How the stability window works (pre-copy):
  - `_wait_for_stable_mtime(path, checks, interval, max_wait)` reads the file’s `mtime` (and size) repeatedly.
  - It requires `COPY_STABILITY_CHECKS` consecutive identical observations, spaced by `COPY_STABILITY_INTERVAL_SEC`, all within `COPY_STABILITY_MAX_WAIT_SEC`. If the sequence can’t be achieved in time, this attempt is deferred (with backoff) and retried.

- Related settings:
  - `SKIP_WHEN_TEMP_LOCK_PRESENT=True` causes the watcher to delay work if `~$<filename>` exists.
  - `POLLING_STABLE_CHECKS` defines how many consecutive “no change” observations are required before a visible compare is allowed in the polling loop.

---

## 5) Robocopy exit codes (0–16): human-friendly reference

Robocopy’s exit code is a bitmask where:
- 1 = Files were copied
- 2 = Extra files or directories were detected (on destination not present in source)
- 4 = Mismatched files or directories were detected (file time/size differences)
- 8 = Some files or directories could not be copied (failures occurred)

Codes 0–7 (no failure bit) are treated as success. Codes ≥8 (the 8-bit set) indicate failure. 16 is a special “serious error.”

Common combinations and meanings:
- 0: No files were copied; no failures; no mismatches. Everything is already up to date (often seen when cache already has the same file). Note: if this occurs on a first-time copy and the destination file doesn’t exist, treat as a logic/validation problem.
- 1: One or more files were copied successfully (new or updated files).
- 2: Extra files or directories were detected; no files were copied (clean-up likely needed on destination).
- 3: Files were copied AND extra files were detected.
- 4: Mismatched files or directories were detected; no files were copied.
- 5: Files were copied AND mismatches were detected.
- 6: Extra files were detected AND mismatches were detected.
- 7: Files were copied AND extra files AND mismatches were detected.
- 8: Some files or directories could not be copied (copy errors occurred and retry limit exceeded). Failure.
- 9: Failure (8) + files were copied.
- 10: Failure (8) + extras detected.
- 11: Failure (8) + files copied + extras detected.
- 12: Failure (8) + mismatches detected.
- 13: Failure (8) + files copied + mismatches detected.
- 14: Failure (8) + extras + mismatches.
- 15: Failure (8) + files copied + extras + mismatches.
- 16: Serious error. Robocopy did not copy any files. One or more critical errors occurred (e.g., invalid arguments, insufficient memory/resources, catastrophic I/O, denied access to both source and destination). Check the detailed Robocopy output.

Practical rule used by this project: `0..7 = success`, `>7 = failure`.

---

## 6) Interpreting your errors: "robocopy rc=16"

- Yes, seeing `robocopy rc=16` means Robocopy was in fact launched and returned with exit code 16 (a serious error). It is not a Python-side mock; the child process executed and failed.
- Likely causes:
  - Invalid paths or quoting (source/destination/file name resolution errors).
  - Destination path not accessible or write-protected; source not readable.
  - Path too long or unusual characters (rare with the current cache name sanitation; source side could still be problematic).
  - Network/Share issues (intermittent disconnects) leading to catastrophic failure.
  - Environment/permission problems preventing Robocopy from initializing properly.
- What to do:
  - Check the printed debug command and confirm directories and filename are correct.
  - Verify the destination cache directory exists (the code ensures it does); confirm permissions.
  - If paths are extremely long, consider enabling extended path support or adjusting folder depth.
  - If this recurs on the same path, try PowerShell `Copy-Item` engine for that file as a test.

---

## 7) Recent robustness enhancements added (2025-08-29)

To make Robocopy behavior more observable and reliable, we added the following to `utils/cache.py`:

- Tool availability checks with `shutil.which()` for `robocopy` and `powershell`. If missing, it’s clearer why the copy didn’t run.
- Debug prints (guarded by `SHOW_DEBUG_MESSAGES`) of the exact command and the return code.
- Post-copy validation (destination must exist; size must be ≥ source; tolerant mtime check). If validation fails even with a success return code (e.g., rc=0 but first-time destination missing), it raises an error to trigger retry.
- Added `/J` (unbuffered I/O) flag for Robocopy to improve stability on large files.
- Fixed ops logging to record the actual engine used (was always logging `python` before).

These changes help diagnose issues like “Robocopy returned success but no file is present,” and provide clear evidence of which engine was used.

---

## 8) Quick FAQ

- Q: "Robocopy 係唔係用一個新嘅 kernel 去 copy?"
  - A: 不是。Robocopy 係以「子進程」形式喺同一個 Windows kernel 上執行，並非新 kernel。優點係子進程結束時，OS 一定會收回該進程開住嘅檔案把手，避免主進程有殘留鎖。

- Q: "非靜默/靜默係乜嘢?"
  - A: 靜默（silent=True）即係唔出表格輸出，主要用喺 preview 或輪詢期避免噪音；非靜默（silent=False）會出表格、寫 CSV、同（如啟用）自動更新 baseline。

- Q: "返回碼 0–7 視為成功，分別代表乜?"
  - A: 見上面 Section 5，0 表示已經 up-to-date、冇錯；1 代表有檔案被複製；其餘為 bitmask 組合（2=有額外檔、4=有不匹配），未含 8 即成功。

- Q: "保存尾段要做穩定窗口，保存尾段係乜?"
  - A: 指 Excel 寫檔最後期，檔案可能仲喺同步/flush 狀態，mtime/size 未穩定。穩定窗口即係連續多次觀察 mtime/size 無變先開始 copy，避免「半熟檔」。

- Q: "見到 robocopy rc=16，係咪真係有調用 Robocopy，只係失敗?"
  - A: 係，代表 Robocopy 真係執行並以 16 結束（嚴重錯誤）。請按 Section 6 的做法排查。

- Q: "Robocopy 可能有幾多種 rc? 可唔可以列晒?"
  - A: 實務上可視為 0–16（bitmask 組合），詳見 Section 5 列表；只要 rc ≥ 8 即屬失敗。

---

## 9) Suggested diagnostics

- Enable `SHOW_DEBUG_MESSAGES=True` and set `COPY_ENGINE='robocopy'`. Observe the exact command and `rc` in console.
- Check `LOG_FOLDER/ops_log/copy_success_YYYYMMDD.csv` to confirm Engine column (`robocopy`). If failures occur, inspect `copy_failures_YYYYMMDD.csv`.
- If `rc=0` but destination was missing (first copy), post-copy validation will now trigger a retry—confirm this behavior.
- For persistent failures on specific paths, test with `COPY_ENGINE='powershell'` (or set `PREFER_SUBPROCESS_FOR_XLSM=True` to auto-switch for .xlsm).

---

End of document.
