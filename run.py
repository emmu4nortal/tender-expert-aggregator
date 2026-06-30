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
from write_master import load_records, dedupe, write_excel

# The full, pre-dedup record set is the single source of truth; the master Excel is a generated
# artifact rebuilt as dedup(batch). Same file extract_requirements.py --all writes.
BATCH_FILE = Path(__file__).parent / "extraction_batch.json"


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


# ── batch (source of truth) + master rebuild ──────────────────────────────────

def _load_batch() -> list[dict]:
    """Load the full pre-dedup record set from extraction_batch.json ([] if absent)."""
    if not BATCH_FILE.exists():
        return []
    return load_records(BATCH_FILE)


def _write_batch(records: list[dict]) -> None:
    BATCH_FILE.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")


def _group_by_path(records: list[dict]) -> dict:
    """Group records by NFC source_relative_path (the per-file unit the batch is updated in)."""
    groups: dict[str, list[dict]] = {}
    for r in records:
        groups.setdefault(_nfc(r.get("source_relative_path", "")), []).append(r)
    return groups


def _rebuild_master(records: list[dict]) -> None:
    """Master = dedup(records). The batch is authoritative; the master is a generated artifact,
    so it is rebuilt in full from the current record set rather than patched in place."""
    deduped = dedupe(records)
    collapsed = len(records) - len(deduped)
    print(f"Batch records:      {len(records)}")
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
    failed: list[Path] = []
    processed: dict[str, list[dict]] = {}   # file_key -> fresh records (may be empty)
    for i, path in enumerate(to_process, 1):
        label = file_key(path)
        print(f"  [{i}/{len(to_process)}] {label} ...", end="", flush=True)
        try:
            recs = extract_file(path)
        except Exception as e:
            print(f" ERROR: {e}")
            failed.append(path)
            continue
        print(f" {len(recs)} records")
        processed[label] = recs

    total_new = sum(len(r) for r in processed.values())
    print(f"\nExtracted {total_new} record(s) from {len(processed)} file(s).")
    if failed:
        print(f"WARNING: {len(failed)} file(s) failed and will not be marked as synced:")
        for p in failed:
            print(f"  {file_key(p)}")

    # Update the batch (source of truth): replace each successfully-processed file's records.
    # A file that now yields 0 records gets an empty entry, so its old rows are dropped (fixes
    # orphaned/stale rows). Failed files are left untouched so they retry next run. Then rebuild
    # the master from the full batch, so dedupe always runs over the current record set.
    batch_groups = _group_by_path(_load_batch())
    for key, recs in processed.items():
        batch_groups[key] = recs
    merged = [r for recs in batch_groups.values() for r in recs]
    _write_batch(merged)
    _rebuild_master(merged)

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

    incoming = load_records(json_path)
    print(f"Records from {json_path.name}: {len(incoming)}")

    # Merge the incoming records into the batch by source path (replace those files' records),
    # then rebuild the master. A partial json patches only its files; a full --all batch replaces
    # everything. Writing the canonical batch keeps it and the master in lock-step.
    batch_groups = _group_by_path(_load_batch())
    for path, recs in _group_by_path(incoming).items():
        batch_groups[path] = recs
    merged = [r for recs in batch_groups.values() for r in recs]
    _write_batch(merged)
    _rebuild_master(merged)


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
