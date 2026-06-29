"""File-level deduplication — keeps newest mtime per normalised filename within each directory."""
import re
from collections import defaultdict
from pathlib import Path

from config import SYNC_ROOT
from enumerate_candidates import enumerate_candidates

# Noise tokens stripped from filenames before grouping.
# Only well-known version/state markers — subject-area words are intentionally
# left intact so distinct documents (different service areas, etc.) stay separate.
_NOISE_TOKENS = [
    'korjattu', 'korj', 'vanha', 'luonnos', 'draft',
    'työstö', 'päivitetty', 'mvp',
]
_NOISE_PAT = re.compile(
    r'(?<![a-zäöå])(' + '|'.join(re.escape(t) for t in _NOISE_TOKENS) + r')(?![a-zäöå])'
)


def normalize_name(filename: str) -> str:
    """Return a normalised form of filename used only for grouping, never stored."""
    name = filename.lower()
    if name.endswith('.xlsx'):
        name = name[:-5]

    # Parenthesised blocks: "(1)", "(korjattu 25.4.2025)", "(04062023 Nortal korjattu versio)"
    name = re.sub(r'\s*\([^)]{0,60}\)\s*', ' ', name)

    # Dates: d.m.yyyy  dd.mm.yy  d m yyyy  pvm040326
    name = re.sub(r'[\s_\.]*\d{1,2}[\s\._]\d{1,2}[\s\._]\d{2,4}', '', name)
    name = re.sub(r'[\s_]*pvm\d+', '', name)

    # Version suffixes: _ver1  _ver02  _v2  " v2"
    # Note: bare trailing " 2" is intentionally NOT stripped — it would also
    # remove part numbers like "Osa-alue 2", collapsing distinct documents.
    name = re.sub(r'[\s_]+ver\d+\b', '', name)
    name = re.sub(r'[\s_]+v\d+\b', '', name)

    # Lone _s suffix (confidential marker)
    name = re.sub(r'[\s_]+s$', '', name)

    # Noise tokens
    name = _NOISE_PAT.sub('', name)

    # Collapse runs of separators to a single space
    name = re.sub(r'[\s_\-]+', ' ', name)
    return name.strip()


def deduplicate(candidates: list[Path]) -> tuple[list[Path], list[tuple[Path, Path]]]:
    """Group by (parent_dir, normalised_name); keep newest mtime per group.

    Returns (kept_sorted, [(dropped, winner), ...]).
    """
    groups: dict[tuple[Path, str], list[Path]] = defaultdict(list)
    for path in candidates:
        key = (path.parent, normalize_name(path.name))
        groups[key].append(path)

    kept, dropped = [], []
    for paths in groups.values():
        if len(paths) == 1:
            kept.append(paths[0])
        else:
            winner = max(paths, key=lambda p: p.stat().st_mtime)
            kept.append(winner)
            for p in paths:
                if p != winner:
                    dropped.append((p, winner))

    return sorted(kept), dropped


def main():
    print("=== Milestone 3: File-level dedupe ===\n")

    candidates = enumerate_candidates()
    print(f"Candidates before dedupe: {len(candidates)}")

    kept, dropped = deduplicate(candidates)
    print(f"Candidates after dedupe:  {len(kept)}")
    print(f"Collapsed:                {len(dropped)}\n")

    if dropped:
        print("Collapsed groups (– dropped  + kept):")
        for dropped_path, winner in sorted(dropped, key=lambda x: x[0]):
            folder = dropped_path.parent.relative_to(SYNC_ROOT)
            print(f"\n  {folder}/")
            print(f"    – {dropped_path.name}")
            print(f"    + {winner.name}")
    else:
        print("No files collapsed.")

    # Normalisation spot-check: show cases where the name actually changed
    print("\n\nNormalisation spot-check (names that changed):")
    shown = 0
    for path in sorted(candidates):
        norm = normalize_name(path.name)
        orig = path.name.lower().removesuffix('.xlsx')
        if orig != norm and shown < 20:
            print(f"  {path.name!r}")
            print(f"    → {norm!r}")
            shown += 1
    if shown == 0:
        print("  (none — all names already normalised)")


if __name__ == "__main__":
    main()
