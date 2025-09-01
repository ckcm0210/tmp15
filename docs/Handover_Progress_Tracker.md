# Watchdog — Handover Progress Tracker (Living Doc)

Purpose
- Single source of truth for current work, decisions, and next actions.
- Updated continuously; if machines reboot or people change, this doc enables immediate resume.

Owner: Rovo Dev
Status: Active (auto-updated during sessions)

---

## 1. Current scope and vision
- Provide /ui/timeline integrated into the main Watchdog app (no .bat), backed by SQLite events index.
- Record full snapshots to LOG_FOLDER/history and optionally to excel_git_repo; compute change counters; index to SQLite.

## 2. Key decisions (agreed)
- Routes: /ui/timeline, /ui/timeline/event/<id>, /ui/timeline/diff/<id>
- Defaults: recent 7 days, page size 50 (max 200), warn if >180 days; base_key shown; Top authors displayed.
- Tri-state filters (has_snapshot/has_summary); meaningful=1 default; diff row cap 1000.
- Global toggle to hide indirect internal changes; grouping by base_key optional; absolute file paths with mapping layer.
- UI language: Traditional Chinese primary with English acronyms.

## 3. What’s done
- M1 Data layer
  - utils/events_db.py: ensure_db, insert_event, query_events, get_event_by_id, get_neighbor_event
  - utils/history.py: save_history_snapshot, sync_history_to_git_repo, compute_change_counters, insert_event_index
  - core/comparison.py: integrated snapshot + Git JSON + SQLite insert after meaningful changes
- M2 Settings/UI
  - config/settings.py: added UI_TIMELINE_* defaults, PATH_MAPPINGS, ENABLE_TIMELINE_SERVER, HOST/PORT, OPEN_TIMELINE_ON_START
  - ui/settings_ui.py: added Timeline tab with fields and grey help text; runtime persistence works
- M3 Routes (MVP)
  - git_viewer.py: added /ui/timeline list with filters, summary bar, grouping; /ui/timeline/event details; /ui/timeline/diff quick diff
- Conversation logging
  - docs/Conversation_Log_YYYY-MM-DD.md: reverse chronological entries per reply

## 4. In progress
- Enhance /ui/timeline filters UI (chips, presets UX, tri-state inputs polished)
- Styling improvements (consistent layout, badges, spacing)

## 5. Next actions (short-term)
- Add client-side JS to merge type checkboxes into `types` on submit; or server-side parse t_* into types.
- Add pagination controls with full query preservation (Prev/Next keep filters).
- Render summary JSON for event details nicely (collapsible).
- Implement path mapping display for snapshot/summary links.

## 6. Risks/Notes
- Large queries: keep limit<=200; show warning for >180d.
- Missing snapshots for early events: quick diff gracefully disabled.
- Multi-machine paths require PATH_MAPPINGS to be set for best UX.

## 7. How to resume work
- Start main app (main.py). If ENABLE_TIMELINE_SERVER=True, /ui/timeline is served at HOST:PORT.
- If filters/UI need changes, edit git_viewer.py; for data issues, inspect utils/events_db.py and utils/history.py.
- Review docs/Timeline_Wireframe_*.md for design alignment.

## 8. Open items for decision
- Whether to add an export API (CSV/JSON) for filtered timeline results
- Whether to add auth/access control for the embedded server

## 9. Checklist
- [x] Events DB helpers
- [x] History counters + inserts
- [x] Embedded Flask startup in main.py
- [x] /ui/timeline MVP with filters/grouping/summary
- [ ] Refine tri-state controls and presets UX
- [ ] Top authors section finalize styling
- [ ] Summary JSON render on details page
- [ ] Full query-preserving pagination

Last updated: 2025-08-31
