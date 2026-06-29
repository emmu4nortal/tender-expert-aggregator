"""Milestone 2: enumerate and filter candidate xlsx files across the sync tree."""
import os
from collections import defaultdict
from pathlib import Path

from config import SYNC_ROOT, EXCLUDED_TOP_LEVEL, CANDIDATE_KEYWORDS, MASTER_PATH


def is_excluded(path: Path) -> bool:
    try:
        rel = path.relative_to(SYNC_ROOT)
    except ValueError:
        return False
    return rel.parts[0] in EXCLUDED_TOP_LEVEL


def keywords_matched(filename: str) -> list[str]:
    lower = filename.lower()
    return [kw for kw in CANDIDATE_KEYWORDS if kw in lower]


def enumerate_candidates() -> list[Path]:
    candidates = []
    for root, dirs, files in os.walk(SYNC_ROOT):
        root_path = Path(root)
        dirs[:] = sorted(
            d for d in dirs
            if not is_excluded(root_path / d) and not d.startswith(".")
        )
        for fname in files:
            fpath = root_path / fname
            if fpath == MASTER_PATH:
                continue
            if not fname.lower().endswith(".xlsx"):
                continue
            if keywords_matched(fname):
                candidates.append(fpath)
    return candidates


def main():
    print("=== Milestone 2: Enumeration + filter ===\n")
    print("Walking tree (this reads only metadata — no file downloads)...")
    candidates = enumerate_candidates()
    print(f"\nTotal candidate files: {len(candidates)}\n")

    # Per-keyword breakdown
    kw_hits: dict[str, list[Path]] = defaultdict(list)
    for path in candidates:
        for kw in keywords_matched(path.name):
            kw_hits[kw].append(path)

    print("Per-keyword hit count (a file may match multiple keywords):")
    for kw in CANDIDATE_KEYWORDS:
        hits = kw_hits[kw]
        print(f"\n  [{kw}]  {len(hits)} file(s)")
        for p in hits[:4]:
            print(f"    {p.relative_to(SYNC_ROOT)}")
        if len(hits) > 4:
            print(f"    … and {len(hits) - 4} more")

    # Files that matched only by a specific keyword (unique matches)
    print("\n\nSample of all candidates (first 20):")
    for p in candidates[:20]:
        matched = keywords_matched(p.name)
        print(f"  [{', '.join(matched)}]  {p.relative_to(SYNC_ROOT)}")
    if len(candidates) > 20:
        print(f"  … and {len(candidates) - 20} more")

    # Flag any keywords that caught zero files
    zero = [kw for kw in CANDIDATE_KEYWORDS if not kw_hits[kw]]
    if zero:
        print(f"\nKeywords with zero matches (candidates for removal): {zero}")
    else:
        print("\nAll keywords matched at least one file.")


if __name__ == "__main__":
    main()
