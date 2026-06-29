"""extract_requirements.py — v2 deterministic requirement-row extractor.

Reads source Excel files directly. Emits one JSON record per requirement row
per expert. requirement_text and evidence are verbatim source cell values.

Usage:
    python extract_requirements.py <file1> [<file2> ...]
    python extract_requirements.py --all
"""

import json
import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

import openpyxl

from config import SYNC_ROOT
from enumerate_candidates import enumerate_candidates
from dedupe import deduplicate

OUTPUT_FILE = Path(__file__).parent / "extraction_batch.json"

# ── sheet filtering ───────────────────────────────────────────────────────────

_HELPER_EXACT = {"pisteet yhteensä", "data", "pisteet", "ohjeet"}
_HELPER_PREFIXES = ("data-", "data -")

def _is_helper_sheet(name: str) -> bool:
    n = name.lower().strip()
    return n in _HELPER_EXACT or any(n.startswith(p) for p in _HELPER_PREFIXES)


# ── requirement row detection ─────────────────────────────────────────────────

# Matches "A 1", "A1", "A 10", "D 1", "10", "1" etc. — letter prefix is optional.
# The $ anchor prevents matching long instruction texts that start with a letter+digit.
_REQ_ROW_RE = re.compile(r'^[A-Z]{0,2}\s*\d+\s*$', re.IGNORECASE)

def _is_req_row(b_val) -> bool:
    if b_val is None:
        return False
    return bool(_REQ_ROW_RE.match(str(b_val).strip()))


# ── format detection ──────────────────────────────────────────────────────────

_DATE_PREFIX = re.compile(r'^\d{1,2}/\d{4}')

def _detect_format(ws) -> int:
    """
    1 = narrative (G starts with date like "08/2015")
    2 = parallel columns (G and H are both numbered lists in separate columns)
    3 = consolidated prose (fallback — G contains plain prose, not a date or numbered list)
    Inspects only mandatory requirement rows for detection.
    """
    for row in ws.iter_rows(min_row=2, values_only=True):
        b = row[1] if len(row) > 1 else None
        c = row[2] if len(row) > 2 else None
        d = row[3] if len(row) > 3 else None
        g = str(row[6]).strip() if len(row) > 6 and row[6] else ""
        h = str(row[7]).strip() if len(row) > 7 and row[7] else ""

        if not _is_req_row(b) or not c or not g:
            continue
        if d and "pisteytettävä" in str(d).lower():
            continue  # skip scoring rows for format detection

        if _DATE_PREFIX.match(g):
            return 1
        if re.match(r'^\d+\.\s*\S', g) and re.match(r'^\d+\.\s*\S', h):
            return 2
        # keep scanning — a plain-text row (e.g. "Koulutus") must not
        # short-circuit detection before the parallel-column rows are seen

    return 3


# ── name / role finder ────────────────────────────────────────────────────────

# Matches "Asiantuntijan rooli:", "Asiantuntijan 1 rooli:", "Asiantuntijan 1. rooli:"
_ROLE_MARKER = re.compile(r'asiantuntijan[\s\d.]*rooli\s*:\s*', re.IGNORECASE)

def _find_name_role(ws) -> tuple[str, str]:
    for row in ws.iter_rows(min_row=1, max_row=50, values_only=True):
        b = str(row[1] or "").strip() if len(row) > 1 else ""
        d = str(row[3] or "").strip() if len(row) > 3 else ""
        if _ROLE_MARKER.search(b) and d:
            role = _ROLE_MARKER.sub("", b).strip().rstrip(":")
            return d, role
    return "", ""


# ── template noise detection ──────────────────────────────────────────────────

_NOISE = (
    # Column header / sub-header instructions
    "kuvaus", "asiakkaat", "toimeksiantojen", "kuvaile", "kirjoita",
    "vastauksesta tulee ilmetä", "tarkennukset tulee", "esimerkkikuvaus",
    # Unfilled template instruction cells ("Ohje tarjoajalle: Täytä...")
    "ohje tarjoajalle",
    # Unfilled "fill in here" placeholders
    "täytä kuvaus", "täytä tähän", "täytä kokemus",
    # Template example answers left in place
    "esimerkki vastauksesta", "esimerkkiasiakas",
)

def _is_template(val: str) -> bool:
    v = val.lower().strip()
    return any(v.startswith(t) for t in _NOISE)


# ── fake / unfilled developer name detection ──────────────────────────────────

_FAKE_NAME = (
    "asiantuntijan nimi",   # header label left unfilled ("Asiantuntijan nimi",
                            # "Asiantuntijan nimi (nimettävä tarjouksessa)", etc.)
    "ohje tarjoajalle",     # template instruction left in the name cell
)

def _is_fake_name(val: str) -> bool:
    v = val.lower().strip().lstrip("(")
    return not v or any(v.startswith(t) for t in _FAKE_NAME)


# ── cell value helper ─────────────────────────────────────────────────────────

def _s(row, idx: int) -> str:
    if idx >= len(row) or row[idx] is None:
        return ""
    return str(row[idx]).strip()


# ── evidence builders ─────────────────────────────────────────────────────────

def _evidence_fmt1_fmt3(row, is_scoring: bool) -> str:
    """Format 1 and 3: G = mandatory evidence, H = scoring evidence."""
    return _s(row, 7 if is_scoring else 6)


# Column labels for Format 2 evidence, in source order (clients..description).
_FMT2_LABELS = ["Asiakkaat", "Projektit", "Ajankohta", "Rooli", "HTP", "Kuvaus"]

def _evidence_fmt2(row, is_scoring: bool) -> str:
    """
    Format 2: parallel numbered-list columns. Stored verbatim — never split or
    zipped — so unequal column lengths cannot silently drop or mis-pair items.

    Mandatory rows  → clients=G(6) projects=H(7) dates=I(8) roles=J(9) htps=K(10) desc=L(11)
    Scoring rows    → score=G(6)   clients=H(7) projects=I(8) dates=J(9) roles=K(10) htps=L(11) desc=M(12)
    """
    primary = 7 if is_scoring else 6           # first content column (clients)
    content_idx = range(primary, primary + 6)  # clients..description

    # Genuine parallel numbered lists → store each non-empty column verbatim, labelled.
    if re.match(r'^\d+\.\s', _s(row, primary)):
        blocks = []
        for label, idx in zip(_FMT2_LABELS, content_idx):
            val = _s(row, idx)
            if val:
                blocks.append(f"{label}:\n{val}")
        return "\n\n".join(blocks)

    # Free-text row inside a Format-2 sheet (e.g. "Koulutus", a certification answer).
    # Not a column structure → join non-empty cells verbatim, no labels, no numbering.
    # The answer may sit in any column (some rows leave clients empty), so scan all.
    parts = [_s(row, idx) for idx in content_idx]
    val = "\n".join(p for p in parts if p)
    return "" if _is_template(val) else val


# ── per-sheet extraction ──────────────────────────────────────────────────────

def _extract_sheet(ws, rel_path: str, file_name: str,
                   mtime: str, today: str) -> list[dict]:
    developer_name, role = _find_name_role(ws)
    if _is_fake_name(developer_name):
        return []
    fmt = _detect_format(ws)
    records = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 4:
            continue

        b, c, d = row[1], row[2], row[3]

        if not _is_req_row(b):
            continue

        req_text = str(c or "").strip()
        if not req_text or _is_template(req_text):
            continue

        is_scoring = bool(d and "pisteytettävä" in str(d).lower())

        if fmt == 2:
            # Labelled Format-2 evidence starts with "Asiakkaat:" etc., which would
            # false-match the template-noise list; its template filtering happens inside
            # _evidence_fmt2 (free-text branch), so only the empty check applies here.
            evidence = _evidence_fmt2(row, is_scoring)
            if not evidence:
                continue
        else:
            evidence = _evidence_fmt1_fmt3(row, is_scoring)
            if not evidence or _is_template(evidence):
                continue

        records.append({
            "developer_name": developer_name,
            "role": role,
            "requirement_text": req_text,
            "evidence": evidence,
            "technologies": "",
            "domain_or_industry": "",
            "source_file_name": file_name,
            "source_relative_path": rel_path,
            "source_sheet": ws.title,
            "source_last_modified": mtime,
            "extracted_date": today,
        })

    return records


# ── per-file extraction ───────────────────────────────────────────────────────

def extract_file(path: Path) -> list[dict]:
    rel = unicodedata.normalize("NFC", str(path.relative_to(SYNC_ROOT)))
    try:
        mtime = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
    except OSError:
        mtime = ""
    today = date.today().isoformat()

    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        print(f"  ERROR opening {path.name}: {e}", file=sys.stderr)
        return []

    records = []
    for ws in wb.worksheets:
        if _is_helper_sheet(ws.title):
            continue
        records.extend(_extract_sheet(ws, rel, path.name, mtime, today))

    wb.close()
    return records


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    args = sys.argv[1:]
    if not args:
        print(
            "Usage:\n"
            "  python extract_requirements.py <file1> [<file2> ...]\n"
            "  python extract_requirements.py --all",
            file=sys.stderr,
        )
        sys.exit(1)

    if args == ["--all"]:
        candidates, dropped = deduplicate(enumerate_candidates())
        print(f"Candidates: {len(candidates)}  ({len(dropped)} dropped by file-level dedupe)")
        paths = candidates
    else:
        paths = []
        for a in args:
            p = Path(a)
            if not p.is_absolute():
                p = SYNC_ROOT / a
            paths.append(p)

    all_records: list[dict] = []
    for i, path in enumerate(paths, 1):
        try:
            label = str(path.relative_to(SYNC_ROOT))
        except ValueError:
            label = path.name
        print(f"[{i}/{len(paths)}] {label} ... ", end="", flush=True)
        recs = extract_file(path)
        print(f"{len(recs)} records")
        all_records.extend(recs)

    OUTPUT_FILE.write_text(
        json.dumps(all_records, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    missing_text = sum(1 for r in all_records if not r.get("requirement_text"))
    empty_ev    = sum(1 for r in all_records if not r.get("evidence"))
    print(f"\nTotal records:      {len(all_records)}")
    print(f"Missing req_text:   {missing_text}")
    print(f"Empty evidence:     {empty_ev}")
    print(f"Written to:         {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
