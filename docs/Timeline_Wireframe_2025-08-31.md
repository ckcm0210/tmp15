# /timeline Page — Wireframe (Draft)

Updated: 2025-08-31
Owner: Rovo Dev
Status: For review/approval before implementation

---

## 1) Goal and scope

Provide a fast, paginated timeline view of Excel change events stored in SQLite (`utils/events_db.py`).
- Read-only UI consuming the events index; writing/insertion handled by processing pipeline.
- Designed to scale to large histories; default filters keep the list snappy.
- Each event links to its snapshot/summary and a quick diff to the previous event.

Out of scope (this phase): multi-file aggregated diffs; editing data; Git repo browsing (handled by git_viewer.py).

---

## 2) Data model alignment (events table)

From `utils/events_db.py`:
- id (PK, auto)
- base_key (TEXT, NOT NULL) — stable key per file (e.g., `Testing.xlsx__abcd1234`)
- file_path (TEXT, NOT NULL) — absolute path at event time (or canonical abs path)
- event_time (TEXT, ISO timestamp)
- excel_mtime (REAL) — source file mtime
- source_size (INTEGER)
- last_author (TEXT)
- git_commit_sha (TEXT)
- snapshot_path (TEXT) — local compressed snapshot path under LOG_FOLDER/history/<base_key>/
- summary_path (TEXT) — summary JSON path (optional)
- total_changes (INTEGER)
- dvc (INTEGER) — Direct Value Changes
- fci (INTEGER) — Formula Change Internal
- xrlc (INTEGER) — External Ref Link Change
- xru (INTEGER) — External Refresh Update
- addc (INTEGER) — Added cells
- delc (INTEGER) — Deleted cells

Implications for UI:
- Filters and chips will refer to these counters and meta fields.
- Quick diff relies on ordering per `event_time` and neighboring events for the same `base_key`.

---

## 3) Routes and URL structure (proposed)

- GET `/ui/timeline`
  - Landing view with filters and results list. Defaults to last 7 days, all files.
- GET `/ui/timeline?base_key=...&q=...&author=...&from=...&to=...&types=dvc,fci,xrlc,xru,addc,delc&min_total=1&has_snapshot=1&has_summary=&sort=desc&limit=50&cursor=`
  - Query string parameters:
    - base_key: string, exact match; if omitted, show across files.
    - q: substring case-insensitive match on file_path (and optionally base_key).
    - author: substring match on last_author.
    - from/to: ISO datetime range for event_time; defaults to from=now-7d, to=now.
    - types: comma-separated list out of [dvc,fci,xrlc,xru,addc,delc]; interpreted as "any of these counters > 0".
    - min_total: minimum total_changes (>=). Default 1.
    - has_snapshot: 1/0 filter by snapshot_path presence.
    - has_summary: 1/0 filter by summary_path presence.
    - sort: desc|asc based on event_time (default desc).
    - limit: page size (10–200). Default 50.
    - cursor: opaque pagination token (or page/offset for v1). For v1, use `after_id` and `before_id` if easier.
- GET `/ui/timeline/event/<id>`
  - Event details drawer/page (raw JSON, links to open snapshot/summary paths if available).
- GET `/ui/timeline/diff/<id>`
  - Quick diff between this event and its previous sibling (same base_key) by event_time.
  - Query options: `compare=prev|next|id:<other_id>`; `meaningful=1`.

Notes:
- For MVP we can use offset-based pagination: `page=1&limit=50`. Later upgrade to cursor.
- All routes are read-only and safe to bookmark/share.

---

## 4) Filters panel (top)

Layout (left to right, wrapping on small screens):
- File (base_key or file search)
  - Text input with suggestions: when user types, show matched `base_key` + filename.
  - Alternatively two inputs: "Search file" (q) and an advanced "Base key" field.
- Author
  - Text input; substring case-insensitive.
- Date range
  - Presets: Last 24h, 7d (default), 14d, 30d, All. Custom from/to via date-time pickers.
- Types (badges toggle, multi-select)
  - DVC, FCI, XRLC, XRU, ADD, DEL; selected types mean "counter > 0".
- Min total changes
  - Slider or numeric input (0–500, step 1). Default 1.
- Has snapshot / has summary
  - Two checkboxes; tri-state allowed (ignored/yes/no). Default: has_snapshot=yes.
- Sort
  - Newest first (desc) or oldest first (asc). Default desc.
- Page size
  - Select: 25, 50 (default), 100, 200.
- Apply / Reset buttons

UX details:
- Submit updates the query string; supports deep-linking. Keyboard Enter applies filters.
- A chip row mirrors active filters for quick removal (e.g., [author: Alex x], [type: DVC x]).

---

## 5) Summary bar (below filters)

Computed over the current filtered result set (for the first page and optionally the total):
- Total events (count)
- Unique files (base_key count)
- Sum of counters: total_changes, and each of dvc, fci, xrlc, xru, addc, delc
- Optional: top authors (first 3) with counts

Display as compact stat cards/chips with tooltips explaining acronyms.

---

## 6) Results list (cards or table)

Density: medium. Each item shows the key metadata and affordances.

Per-event item fields:
- Primary line: [time] file name (from file_path) — base_key
- Secondary meta: last_author • size (source_size, pretty) • mtime (excel_mtime) • commit (git_commit_sha short) • id
- Counters (badges): total_changes, dvc, fci, xrlc, xru, addc, delc (hide zeros unless hovered or in compact mode show zeros greyed)
- Actions (right):
  - View snapshot (if snapshot_path) → opens `/ui/timeline/event/<id>` with link to file path on disk.
  - View summary (if summary_path)
  - Quick diff → `/ui/timeline/diff/<id>?compare=prev&meaningful=1`
  - Compare with… (popover to choose previous/next/by id)

Visual hints:
- Highlight when any of [xrlc, xru] > 0 (external-related) with a badge color.
- Muted styling for low-signal events (e.g., total_changes below threshold).

Empty/missing states:
- If snapshot_path missing: show disabled action with tooltip.
- If previous event not found: quick diff disabled with tooltip.

---

## 7) Pagination

- Default: page=1, limit=50. Controls at bottom and top.
- Next / Prev buttons; also show item range: “Showing 1–50 of 1,234 (filtered)”.
- For large result sets, we can implement an `after_id` cursor next.

---

## 8) Event details page `/ui/timeline/event/<id>`

- Header: time, file name, base_key, author, commit.
- Body: two sections
  - Metadata: show all columns nicely formatted.
  - Payload links: snapshot_path and summary_path
    - Snapshot: open file location (local link) or download
    - Summary: render JSON in-page (pretty-printed), collapsible

---

## 9) Quick diff `/ui/timeline/diff/<id>`

- Compares this event’s snapshot with the previous event’s snapshot for the same base_key.
- If either snapshot is missing, show guidance.
- Query params:
  - compare: prev (default) | next | id:<other_id>
  - meaningful: 1/0 (default 1) — only show changes where formula or display value diff or add/del
- Output sections:
  - Summary: totals and breakdown by type
  - Changes table: [Sheet, Address, Type, Old, New, Old F, New F], cap at 1,000 rows with a note
- Implementation note: reuse logic akin to `git_viewer.py` diff with adapters for reading compressed snapshots.

---

## 10) Keyboard and accessibility

- Keyboard: `/` focuses search; `f` opens filters; `n/p` page next/prev.
- Tooltips have aria-labels; color is not the only indicator (icons + text).

---

## 11) Performance and safety

- Query uses indexed `base_key, event_time DESC` and WHERE clauses on columns; avoid LIKE on large text without bounds.
- Limit page size to 200 max; guardrails on date range (warn on > 180 days unless confirmed).
- All filesystem links are sanitized and read-only; no direct file writes from UI.

---

## 12) Edge cases

- Multiple files with the same name (different base_key): always show base_key; allow grouping by base_key in future.
- Clock skew: prefer `event_time` ordering; tie-breaker by `id`.
- Missing counters: treat NULL as 0.
- Legacy rows without snapshot_path/summary_path: handled gracefully.

---

## 13) Open questions

- Should we include a global toggle to hide indirect internal changes in diffs (align with IGNORE_INDIRECT_CHANGES)?
- Do we want a collapsing group per file (accordion) when base_key is omitted?
- Is `file_path` stored as abs path always? If machines differ, we may need a mapping layer for opening files.

---

## 14) Next steps (after approval)

1) Extend `utils/events_db.py` with query helpers supporting filters + pagination.
2) Add history summary computation at compare time and insert into events (populate counters + snapshot/summary paths).
3) Implement Flask blueprint or extend existing `git_viewer.py` with `/ui/timeline*` routes and templates.
4) Add compressed snapshot reader and reuse diff logic.
5) Wire settings toggles (e.g., default date range, page size) to config/runtime_settings.json.
