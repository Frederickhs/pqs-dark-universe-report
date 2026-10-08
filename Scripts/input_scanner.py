from __future__ import annotations

import json
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SETTINGS_FILE = PROJECT_ROOT / "Config" / "project_settings.json"

QUALIFICATION_FILES = [
    "CUMOD0 Qualification.xlsx",
    "CUMOD1 Qualification.xlsx",
    "CUMOD2 Qualification.xlsx",
    "CUMOD3 Qualification.xlsx",
    "DUMOD0 Qualification.xlsx",
    "DUMOD1 Qualification.xlsx",
    "UCMOD0 Qualification.xlsx",
    "UCMOD1 Qualification.xlsx",
    "UCMOD2 Qualification.xlsx",
    "UCMOD3 Qualification.xlsx",
    "UCMOD4 Qualification.xlsx",
]


def newest_xlsx(folder: Path) -> Path | None:
    if not folder.exists():
        return None

    files = [
        file
        for file in folder.glob("*.xlsx")
        if not file.name.startswith("~$")
    ]

    if not files:
        return None

    return max(files, key=lambda file: file.stat().st_mtime)


def check_file(label: str, path: Path | None, errors: list[str]) -> None:
    if path is not None and path.exists():
        print(f"[OK] {label}: {path}")
    else:
        print(f"[MISSING] {label}")

        if path is not None:
            print(f"          Expected: {path}")

        errors.append(label)


def main() -> int:
    print("=" * 70)
    print("PQS DARK UNIVERSE REPORT - INPUT SCANNER")
    print("=" * 70)

    errors: list[str] = []

    if not SETTINGS_FILE.exists():
        print(f"[MISSING] Project settings: {SETTINGS_FILE}")
        return 1

    try:
        with SETTINGS_FILE.open("r", encoding="utf-8") as file:
            settings = json.load(file)
    except json.JSONDecodeError as error:
        print(f"[ERROR] Invalid project_settings.json: {error}")
        return 1

    controlled_files = {
        "Line Schedule": (
            PROJECT_ROOT
            / "Input"
            / "Line Schedule"
            / "Line Schedule.xlsx"
        ),
        "Animation Roster": (
            PROJECT_ROOT
            / "Input"
            / "Rosters"
            / "Animation Roster.xlsx"
        ),
        "Supervisor Roster": (
            PROJECT_ROOT
            / "Input"
            / "Supervisors"
            / "Supervisor Roster.xlsx"
        ),
        "PQS Exceptions": (
            PROJECT_ROOT
            / "Input"
            / "Exceptions"
            / "pqs_exceptions.xlsx"
        ),
        "Module Mapping": (
            PROJECT_ROOT
            / "Config"
            / "module_mapping.csv"
        ),
    }

    print("\nCONTROLLED FILES")

    for label, path in controlled_files.items():
        check_file(label, path, errors)

    print("\nMAXIMO REPORTS")

    operations_report = newest_xlsx(
        PROJECT_ROOT / "Input" / "Maximo" / "Operations"
    )
    animation_report = newest_xlsx(
        PROJECT_ROOT / "Input" / "Maximo" / "Animation"
    )

    check_file(
        "Operations Maximo Report",
        operations_report,
        errors,
    )
    check_file(
        "Animation Maximo Report",
        animation_report,
        errors,
    )

    print("\nQUALIFICATION REPORTS")

    qualification_folder = Path(settings["qualification_folder"])

    if not qualification_folder.exists():
        print(f"[MISSING] Qualification folder: {qualification_folder}")
        errors.append("Qualification folder")
    else:
        print(f"[OK] Qualification folder: {qualification_folder}")

        for filename in QUALIFICATION_FILES:
            check_file(
                filename,
                qualification_folder / filename,
                errors,
            )

    print("\nPROJECT FOLDERS")

    required_folders = [
        PROJECT_ROOT / "Output",
        PROJECT_ROOT / "Archive" / "Outputs",
        PROJECT_ROOT / "Archive" / "Snapshots",
        PROJECT_ROOT / "Logs",
    ]

    for folder in required_folders:
        if folder.exists():
            print(f"[OK] {folder}")
        else:
            print(f"[MISSING] {folder}")
            errors.append(str(folder))

    print("\n" + "=" * 70)

    if errors:
        print(f"SCAN FAILED: {len(errors)} item(s) missing.")

        for item in errors:
            print(f" - {item}")

        return 1

    print("SCAN PASSED: All required inputs and folders were found.")
    return 0


if __name__ == "__main__":
    sys.exit(main())