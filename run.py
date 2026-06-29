"""Main entry point for the Tender Expert Experience Aggregator (v2).

Usage
-----
  python run.py sync
      Enumerate all candidate Excel files, extract only new or changed ones,
      merge into Asiantuntijat_Master.xlsx, and update state.json.

  python run.py write <extraction_json>
      Load a pre-built extraction JSON, merge into master, and write.
      Useful for one-off full re-extractions or manual corrections.

  python run.py status
      Show state.json summary and master row count.
"""
import json
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import openpyxl

from config import SYNC_ROOT, MASTER_PATH, STATE_FILE
from enumerate_candidates import enumerate_candidates
from dedupe import deduplicate
from write_master import load_records, dedupe, write_excel, HEADER_LABELS


# ── state helpers ─────────────────────────────────────────────────────────────

def load_state() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            pass
    return {"files": {}, "last_run": None}


def save_state(state: dict) -> None:
    STATE_FILE.write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _nfc(s: str) -> str:
    """NFC-normalize a path string.

    macOS returns NFD paths from the filesystem; openpyxl reads xlsx cells back
    as NFC. Normalize everything to NFC so comparisons are consistent.
    """
    return unicodedata.normalize("NFC", s) if s else s


def file_key(path: Path) -> str:
    return _nfc(str(path.relative_to(SYNC_ROOT)))


# ── master reader (shared by write and sync) ──────────────────────────────────

def _load_master_records() -> list[dict]:
    """Read all rows from the existing master Excel, remapped to field keys."""
    if not MASTER_PATH.exists():
        return []
    try:
        wb = openpyxl.load_workbook(MASTER_PATH)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        inv = {v: k for k, v in HEADER_LABELS.items()}
        records = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            rec = {}
            for h, v in zip(headers, row):
                key = inv.get(h, h)
                rec[key] = str(v) if v is not None else ""
            records.append(rec)
        wb.close()
        print(f"Existing master rows: {len(records)}")
        return records
    except Exception as e:
        print(f"Warning: could not read existing master ({e}) — starting fresh")
        return []


# ── merge + write (shared by write and sync) ──────────────────────────────────

def _merge_and_write(new_records: list[dict], existing_records: list[dict]) -> None:
    """Remove stale rows for re-extracted sources, merge, dedupe, write master."""
    new_sources = {_nfc(r["source_relative_path"]) for r in new_records if r.get("source_relative_path")}
    kept_existing = [r for r in existing_records if _nfc(r.get("source_relative_path", "")) not in new_sources]
    removed = len(existing_records) - len(kept_existing)
    if removed:
        print(f"Removed {removed} stale row(s) for re-extracted source(s)")

    all_records = kept_existing + new_records
    deduped = dedupe(all_records)
    collapsed = len(all_records) - len(deduped)
    print(f"Total after merge:  {len(all_records)}")
    print(f"After row dedupe:   {len(deduped)}  ({collapsed} duplicate(s) collapsed)")

    write_excel(deduped, MASTER_PATH)
    print(f"Master written to:  {MASTER_PATH}")
    print(f"Relative to root:   {MASTER_PATH.relative_to(SYNC_ROOT)}")


# ── sync command ──────────────────────────────────────────────────────────────

def cmd_sync() -> None:
    from extract_requirements import extract_file

    state = load_state()
    tracked = state.get("files", {})

    candidates, dropped = deduplicate(enumerate_candidates())
    print(f"Candidates: {len(candidates)}  ({len(dropped)} dropped by file-level dedupe)")

    new_files, changed_files, unchanged_files = [], [], []
    for path in candidates:
        key = file_key(path)
        try:
            mtime = path.stat().st_mtime
            size  = path.stat().st_size
        except OSError:
            continue
        if key not in tracked:
            new_files.append(path)
        elif tracked[key]["mtime"] != mtime or tracked[key].get("size") != size:
            changed_files.append(path)
        else:
            unchanged_files.append(path)

    print(f"  New:       {len(new_files)}")
    print(f"  Changed:   {len(changed_files)}")
    print(f"  Unchanged: {len(unchanged_files)}")

    to_process = new_files + changed_files
    if not to_process:
        print("\nNothing to sync — all files unchanged since last run.")
        state["last_run"] = datetime.now().isoformat()
        save_state(state)
        return

    print(f"\nExtracting {len(to_process)} file(s)...")
    all_records: list[dict] = []
    failed: list[Path] = []
    for i, path in enumerate(to_process, 1):
        label = file_key(path)
        print(f"  [{i}/{len(to_process)}] {label} ...", end="", flush=True)
        try:
            recs = extract_file(path)
            print(f" {len(recs)} records")
            all_records.extend(recs)
        except Exception as e:
            print(f" ERROR: {e}")
            failed.append(path)

    print(f"\nExtracted {len(all_records)} record(s) from {len(to_process) - len(failed)} file(s).")
    if failed:
        print(f"WARNING: {len(failed)} file(s) failed and will not be marked as synced:")
        for p in failed:
            print(f"  {file_key(p)}")

    if all_records:
        existing = _load_master_records()
        _merge_and_write(all_records, existing)
    else:
        print("No records extracted — master unchanged.")

    # Update state only for files that were successfully processed
    successful = [p for p in to_process if p not in failed]
    for path in successful:
        key = file_key(path)
        tracked[key] = {"mtime": path.stat().st_mtime, "size": path.stat().st_size}

    state["files"] = tracked
    state["last_run"] = datetime.now().isoformat()
    save_state(state)
    print(f"State updated. {len(tracked)} file(s) now tracked.")


# ── write command ─────────────────────────────────────────────────────────────

def cmd_write(extraction_json: str) -> None:
    json_path = Path(extraction_json)
    if not json_path.exists():
        print(f"ERROR: {json_path} not found", file=sys.stderr)
        sys.exit(1)

    new_records = load_records(json_path)
    print(f"New records from extraction: {len(new_records)}")

    existing = _load_master_records()
    _merge_and_write(new_records, existing)


# ── status command ────────────────────────────────────────────────────────────

def cmd_status() -> None:
    state = load_state()
    tracked = state.get("files", {})
    last = state.get("last_run", "never")
    print(f"State file:      {STATE_FILE}")
    print(f"Tracked files:   {len(tracked)}")
    print(f"Last run:        {last}")
    if MASTER_PATH.exists():
        wb = openpyxl.load_workbook(MASTER_PATH)
        ws = wb.active
        print(f"Master rows:     {ws.max_row - 1} (excl. header)")
        wb.close()
    else:
        print("Master:          not yet created")


# ── CLI dispatch ──────────────────────────────────────────────────────────────

def main() -> None:
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(0)

    cmd = args[0]
    if cmd == "sync":
        cmd_sync()
    elif cmd == "write":
        if len(args) < 2:
            print("Usage: python run.py write <extraction_json>", file=sys.stderr)
            sys.exit(1)
        cmd_write(args[1])
    elif cmd == "status":
        cmd_status()
    else:
        print(f"Unknown command: {cmd}", file=sys.stderr)
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
