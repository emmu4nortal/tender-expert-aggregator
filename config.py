from pathlib import Path

SYNC_ROOT = Path(
    "/Users/panu/Library/CloudStorage/OneDrive-SharedLibraries-Nortal"
    "/Public Sales - Documents"
)

EXCLUDED_TOP_LEVEL = {
    "General",
    "Hinnankorotukset 2024",
    "Asiakkuussuunnitelmat",
}

CANDIDATE_KEYWORDS = [
    "cv",
    "osaaminen",
    "asiantuntij",
    "kokemus",
    "resurssi",
]

MASTER_PATH = (
    SYNC_ROOT
    / "General"
    / "Referenssit"
    / "Asiantuntijat"
    / "Asiantuntijat_Master.xlsx"
)

STATE_FILE = Path(__file__).parent / "state.json"
