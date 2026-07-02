"""enrich.py — Milestone 6 enrichment for the Tender Expert Aggregator.

Populates the `technologies` and `domain_or_industry` fields via a content-keyed side table
(`enrichment.json`) that is joined onto the master at rebuild time (see run.py `_rebuild_master`
+ write_master.apply_enrichment). Storing enrichment outside the batch means it survives
re-extraction of any source file.

HARD CONSTRAINT: no LLM here. `auto` is pure deterministic string matching against curated
dictionaries (enrich_tech.json / enrich_industry.json). Judgment-based tagging is done by Claude
Code interactively via the `todo` -> (tag) -> `apply` round-trip, then `run.py write`.

Commands:
  python3 enrich.py status                 show coverage + unmapped client folders
  python3 enrich.py auto [--force]         deterministic dictionary pass -> enrichment.json
  python3 enrich.py todo [--limit N] [--field tech|domain|both] [--out FILE]
                                           emit un-covered unique rows for interactive tagging
  python3 enrich.py apply <tagged.json>    merge interactively-produced tags into enrichment.json

The master is refreshed by: python3 run.py write extraction_batch.json
"""
import argparse
import json
import re
import sys
from pathlib import Path

from write_master import load_records, dedupe, content_hash

BATCH_FILE = Path(__file__).parent / "extraction_batch.json"
ENRICHMENT_FILE = Path(__file__).parent / "enrichment.json"
TECH_FILE = Path(__file__).parent / "enrich_tech.json"
INDUSTRY_FILE = Path(__file__).parent / "enrich_industry.json"


# ── dictionary loading ──────────────────────────────────────────────────────────

def _load_json(path: Path) -> dict:
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    data.pop("_comment", None)
    return data


def load_enrichment() -> dict:
    if not ENRICHMENT_FILE.exists():
        return {}
    return json.loads(ENRICHMENT_FILE.read_text(encoding="utf-8")) or {}


def save_enrichment(table: dict) -> None:
    # sort keys for a stable, diff-friendly file
    ENRICHMENT_FILE.write_text(
        json.dumps(table, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")


# ── deterministic taggers (no LLM) ───────────────────────────────────────────────

def _compile_tech(tech_dict: dict) -> list:
    """[(canonical_tag, compiled_regex)]. Each pattern is matched token-boundary aware: it must
    not be flanked by another token char (letters, digits, or + # . /), so 'java' does not match
    inside 'javascript' and 'sql' not inside 'mysql', while 'c#', '.net', 'ci/cd' still match."""
    edge = r'[A-Za-z0-9+#./]'
    compiled = []
    for tag, patterns in tech_dict.items():
        alts = "|".join(re.escape(p) for p in patterns)
        rx = re.compile(rf'(?<!{edge})(?:{alts})(?!{edge})', re.IGNORECASE)
        compiled.append((tag, rx))
    return compiled


def match_tech(text: str, compiled: list) -> str:
    """Comma-separated canonical tags whose patterns occur in `text`, in dictionary order."""
    if not text:
        return ""
    found = [tag for tag, rx in compiled if rx.search(text)]
    return ", ".join(found)


def industry_for(rec: dict, industry_map: dict) -> str:
    """Industry tag from the top-level client folder of the record's source path."""
    top = (rec.get("source_relative_path") or "").split("/")[0]
    return industry_map.get(top, "")


def _ctx(rec: dict) -> str:
    req = (rec.get("requirement_text") or "").strip().replace("\n", " ")
    return f"{(rec.get('developer_name') or '').strip()} · {req[:50]}"


# ── unique-row helper ─────────────────────────────────────────────────────────

def unique_rows() -> list:
    """Deduped master rows (same set write_master.dedupe produces), each carrying content_hash."""
    return dedupe(load_records(BATCH_FILE))


# ── commands ──────────────────────────────────────────────────────────────────

def cmd_status(_args) -> None:
    rows = unique_rows()
    table = load_enrichment()
    industry_map = _load_json(INDUSTRY_FILE)

    have_tech = have_dom = have_both = 0
    for rec in rows:
        e = table.get(content_hash(rec)) or {}
        t = bool((e.get("technologies") or "").strip())
        d = bool((e.get("domain_or_industry") or "").strip())
        have_tech += t
        have_dom += d
        have_both += t and d
    n = len(rows)

    def pct(x):
        return f"{100*x/n:.0f}%" if n else "-"

    print(f"Unique master rows:      {n}")
    print(f"Side-table entries:      {len(table)}  ({ENRICHMENT_FILE.name})")
    print(f"  technologies filled:   {have_tech}  ({pct(have_tech)})")
    print(f"  domain filled:         {have_dom}  ({pct(have_dom)})")
    print(f"  both filled:           {have_both}  ({pct(have_both)})")

    # Unmapped client folders (industry gaps), by row count.
    from collections import Counter
    unmapped = Counter()
    for rec in rows:
        top = (rec.get("source_relative_path") or "").split("/")[0]
        if top and top not in industry_map:
            unmapped[top] += 1
    if unmapped:
        print(f"\nUnmapped client folders (add to {INDUSTRY_FILE.name}):")
        for folder, c in unmapped.most_common():
            print(f"  {c:5d}  {folder}")


def cmd_auto(args) -> None:
    rows = unique_rows()
    table = load_enrichment()
    compiled = _compile_tech(_load_json(TECH_FILE))
    industry_map = _load_json(INDUSTRY_FILE)

    tech_added = dom_added = entries_touched = 0
    for rec in rows:
        h = content_hash(rec)
        entry = table.get(h, {})
        text = f"{rec.get('requirement_text', '')}\n{rec.get('evidence', '')}"

        new_tech = match_tech(text, compiled)
        if new_tech and (args.force or not (entry.get("technologies") or "").strip()):
            if new_tech != entry.get("technologies"):
                entry["technologies"] = new_tech
                tech_added += 1

        new_dom = industry_for(rec, industry_map)
        if new_dom and (args.force or not (entry.get("domain_or_industry") or "").strip()):
            if new_dom != entry.get("domain_or_industry"):
                entry["domain_or_industry"] = new_dom
                dom_added += 1

        # Only persist entries that carry at least one tag (keeps the file meaningful/small).
        if (entry.get("technologies") or entry.get("domain_or_industry")):
            entry["_ctx"] = _ctx(rec)
            if h not in table:
                entries_touched += 1
            table[h] = entry

    save_enrichment(table)
    print(f"auto: technologies set/updated on {tech_added} rows, domain on {dom_added} rows.")
    print(f"Side table now has {len(table)} entries -> {ENRICHMENT_FILE.name}")
    print("Next: python3 run.py write extraction_batch.json   (to rebuild the master)")


def cmd_todo(args) -> None:
    rows = unique_rows()
    table = load_enrichment()

    def uncovered(e):
        t = bool((e.get("technologies") or "").strip())
        d = bool((e.get("domain_or_industry") or "").strip())
        if args.field == "tech":
            return not t
        if args.field == "domain":
            return not d
        return not (t and d)  # both

    out = []
    for rec in rows:
        h = content_hash(rec)
        e = table.get(h) or {}
        if not uncovered(e):
            continue
        out.append({
            "content_hash": h,
            "developer_name": rec.get("developer_name", ""),
            "role": rec.get("role", ""),
            "requirement_text": rec.get("requirement_text", ""),
            "evidence": rec.get("evidence", ""),
            "technologies": e.get("technologies", ""),
            "domain_or_industry": e.get("domain_or_industry", ""),
        })
        if args.limit and len(out) >= args.limit:
            break

    payload = json.dumps(out, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(payload, encoding="utf-8")
        print(f"Wrote {len(out)} un-covered row(s) (field={args.field}) to {args.out}")
    else:
        print(payload)


def cmd_apply(args) -> None:
    incoming = json.loads(Path(args.tagged).read_text(encoding="utf-8"))
    if isinstance(incoming, dict):
        incoming = [{"content_hash": k, **v} for k, v in incoming.items()]

    table = load_enrichment()
    applied = 0
    for item in incoming:
        h = item.get("content_hash") or item.get("key")
        if not h:
            continue
        entry = table.get(h, {})
        for field in ("technologies", "domain_or_industry"):
            if field in item and (item[field] or "").strip():
                entry[field] = item[field].strip()
        if "_ctx" in item:
            entry["_ctx"] = item["_ctx"]
        if entry.get("technologies") or entry.get("domain_or_industry"):
            table[h] = entry
            applied += 1

    save_enrichment(table)
    print(f"apply: merged tags for {applied} row(s); side table now {len(table)} entries.")
    print("Next: python3 run.py write extraction_batch.json")


# ── CLI ─────────────────────────────────────────────────────────────────────────

def main() -> None:
    p = argparse.ArgumentParser(description="M6 enrichment for the tender expert aggregator.")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("status", help="show coverage + unmapped client folders")

    a = sub.add_parser("auto", help="deterministic dictionary pass")
    a.add_argument("--force", action="store_true",
                   help="overwrite existing non-empty tags (default: only fill empties)")

    t = sub.add_parser("todo", help="emit un-covered rows for interactive tagging")
    t.add_argument("--limit", type=int, default=0, help="max rows to emit (0 = all)")
    t.add_argument("--field", choices=("tech", "domain", "both"), default="both",
                   help="which coverage gap to select on")
    t.add_argument("--out", help="write JSON to this file instead of stdout")

    ap = sub.add_parser("apply", help="merge interactively-produced tags")
    ap.add_argument("tagged", help="JSON file: {hash: {technologies, domain_or_industry}} or a list")

    args = p.parse_args()
    {"status": cmd_status, "auto": cmd_auto, "todo": cmd_todo, "apply": cmd_apply}[args.cmd](args)


if __name__ == "__main__":
    main()
