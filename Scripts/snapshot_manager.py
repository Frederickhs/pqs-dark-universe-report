from __future__ import annotations

import csv
from dataclasses import replace
from datetime import date, datetime
from pathlib import Path

from module_result_processor import ModuleResult, build_results


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SNAPSHOT_FOLDER = PROJECT_ROOT / "Archive" / "Snapshots"
SNAPSHOT_PREFIX = "pqs_snapshot_"
SNAPSHOT_DATE_FORMAT = "%Y-%m-%d_%H%M%S"

SNAPSHOT_COLUMNS = [
    "Run Timestamp",
    "SSO",
    "Module Code",
    "Result",
    "Completed Tasks",
    "Total Tasks",
    "Percentage",
    "Child Status",
    "Work Order",
    "Source",
]

PRESERVED_RESULTS = {
    "ACK Error - Technician Action Required",
    "Completed",
    "Completed - Awaiting Maximo Status",
    "Review Required",
    "Not Required",
    "Not Started",
    "No Maximo Record",
}


class SnapshotRecord:
    def __init__(
        self,
        run_timestamp: datetime,
        sso: str,
        module_code: str,
        result: str,
        completed_tasks: float | None,
        total_tasks: float | None,
        percentage: float | None,
        child_status: str,
        work_order: str,
        source: str,
    ) -> None:
        self.run_timestamp = run_timestamp
        self.sso = sso
        self.module_code = module_code
        self.result = result
        self.completed_tasks = completed_tasks
        self.total_tasks = total_tasks
        self.percentage = percentage
        self.child_status = child_status
        self.work_order = work_order
        self.source = source


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())


def to_float(value: object) -> float | None:
    text = clean_text(value)
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def format_number(value: int | float | None) -> str:
    if value is None:
        return ""
    number = float(value)
    if number.is_integer():
        return str(int(number))
    return str(number)


def snapshot_timestamp_from_path(path: Path) -> datetime | None:
    stem = path.stem
    if not stem.startswith(SNAPSHOT_PREFIX):
        return None

    value = stem[len(SNAPSHOT_PREFIX):]
    try:
        return datetime.strptime(value, SNAPSHOT_DATE_FORMAT)
    except ValueError:
        return None


def list_snapshots() -> list[tuple[datetime, Path]]:
    SNAPSHOT_FOLDER.mkdir(parents=True, exist_ok=True)
    snapshots: list[tuple[datetime, Path]] = []

    for path in SNAPSHOT_FOLDER.glob(f"{SNAPSHOT_PREFIX}*.csv"):
        timestamp = snapshot_timestamp_from_path(path)
        if timestamp is not None:
            snapshots.append((timestamp, path))

    return sorted(snapshots, key=lambda item: item[0])


def find_prior_date_snapshot(run_timestamp: datetime) -> Path | None:
    eligible = [
        (timestamp, path)
        for timestamp, path in list_snapshots()
        if timestamp.date() < run_timestamp.date()
    ]

    if not eligible:
        return None

    return max(eligible, key=lambda item: item[0])[1]


def read_snapshot(path: Path) -> dict[tuple[str, str], SnapshotRecord]:
    lookup: dict[tuple[str, str], SnapshotRecord] = {}

    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)

        if not reader.fieldnames or not set(SNAPSHOT_COLUMNS).issubset(reader.fieldnames):
            raise ValueError(f"Snapshot is missing required columns: {path}")

        for row in reader:
            timestamp = datetime.fromisoformat(row["Run Timestamp"])
            record = SnapshotRecord(
                run_timestamp=timestamp,
                sso=clean_text(row["SSO"]),
                module_code=clean_text(row["Module Code"]).upper(),
                result=clean_text(row["Result"]),
                completed_tasks=to_float(row["Completed Tasks"]),
                total_tasks=to_float(row["Total Tasks"]),
                percentage=to_float(row["Percentage"]),
                child_status=clean_text(row["Child Status"]),
                work_order=clean_text(row["Work Order"]),
                source=clean_text(row["Source"]),
            )
            lookup[(record.sso, record.module_code)] = record

    return lookup


def numbers_equal(left: float | None, right: float | None) -> bool:
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    return abs(left - right) < 0.000001


def apply_snapshot_comparison(
    results: dict[tuple[str, str], ModuleResult],
    previous: dict[tuple[str, str], SnapshotRecord] | None,
) -> tuple[dict[tuple[str, str], ModuleResult], list[str]]:
    if not previous:
        return dict(results), []

    compared: dict[tuple[str, str], ModuleResult] = {}
    comparison_notes: list[str] = []

    for key, current in results.items():
        prior = previous.get(key)

        if current.result in PRESERVED_RESULTS or prior is None:
            compared[key] = current
            continue

        if current.total_tasks is not None and prior.total_tasks is not None:
            if not numbers_equal(float(current.total_tasks), prior.total_tasks):
                note = (
                    f"Task Total Changed: {current.module_code} changed from "
                    f"{format_number(prior.total_tasks)} to "
                    f"{format_number(current.total_tasks)} tasks"
                )
                compared[key] = replace(
                    current,
                    result="Task Total Changed - Review Data",
                    note=note,
                )
                comparison_notes.append(f"{current.sso} | {note}")
                continue

        current_completed = float(current.completed_tasks or 0)
        prior_completed = float(prior.completed_tasks or 0)

        if current_completed < prior_completed:
            note = (
                f"Progress Decreased: {current.module_code} changed from "
                f"{format_number(prior.completed_tasks)}/"
                f"{format_number(prior.total_tasks)} to "
                f"{format_number(current.completed_tasks)}/"
                f"{format_number(current.total_tasks)}"
            )
            compared[key] = replace(
                current,
                result="Progress Decreased - Review Data",
                note=note,
            )
            comparison_notes.append(f"{current.sso} | {note}")
        elif current_completed > prior_completed:
            compared[key] = replace(
                current,
                result="In Progress",
                note=f"In Progress: {current.module_code}",
            )
        else:
            compared[key] = replace(
                current,
                result="No Progress",
                note=f"No Progress: {current.module_code}",
            )

    return compared, comparison_notes


def write_snapshot(
    results: dict[tuple[str, str], ModuleResult],
    run_timestamp: datetime,
) -> Path:
    SNAPSHOT_FOLDER.mkdir(parents=True, exist_ok=True)
    path = SNAPSHOT_FOLDER / (
        f"{SNAPSHOT_PREFIX}{run_timestamp.strftime(SNAPSHOT_DATE_FORMAT)}.csv"
    )

    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=SNAPSHOT_COLUMNS)
        writer.writeheader()

        for key in sorted(results):
            result = results[key]
            writer.writerow(
                {
                    "Run Timestamp": run_timestamp.isoformat(timespec="seconds"),
                    "SSO": result.sso,
                    "Module Code": result.module_code,
                    "Result": result.result,
                    "Completed Tasks": format_number(result.completed_tasks),
                    "Total Tasks": format_number(result.total_tasks),
                    "Percentage": "" if result.percentage is None else result.percentage,
                    "Child Status": result.child_status,
                    "Work Order": result.work_order,
                    "Source": result.source,
                }
            )

    return path


def build_results_with_snapshot(
    run_timestamp: datetime | None = None,
    save_snapshot: bool = False,
):
    run_timestamp = run_timestamp or datetime.now()
    people, initial_results, warnings, unknown_modules = build_results()

    previous_path = find_prior_date_snapshot(run_timestamp)
    previous = read_snapshot(previous_path) if previous_path else None

    final_results, comparison_notes = apply_snapshot_comparison(
        initial_results,
        previous,
    )

    snapshot_path = None
    if save_snapshot:
        snapshot_path = write_snapshot(final_results, run_timestamp)

    return (
        people,
        final_results,
        warnings,
        unknown_modules,
        previous_path,
        snapshot_path,
        comparison_notes,
    )


def main() -> int:
    run_timestamp = datetime.now()

    (
        people,
        results,
        warnings,
        unknown_modules,
        previous_path,
        snapshot_path,
        comparison_notes,
    ) = build_results_with_snapshot(
        run_timestamp=run_timestamp,
        save_snapshot=True,
    )

    counts: dict[str, int] = {}
    for result in results.values():
        counts[result.result] = counts.get(result.result, 0) + 1

    print("=" * 84)
    print("PQS DARK UNIVERSE REPORT - SNAPSHOT COMPARISON")
    print("=" * 84)
    print(f"Run timestamp: {run_timestamp:%Y-%m-%d %H:%M:%S}")
    print(f"Employees: {len(people)}")
    print(f"Module results: {len(results)}")
    print(
        "Comparison snapshot: "
        + (str(previous_path) if previous_path else "None - first-run baseline")
    )
    print(f"Current snapshot saved: {snapshot_path}")

    print("\nFINAL RESULT COUNTS")
    print("-" * 84)
    for result_name, count in sorted(counts.items()):
        print(f"  {result_name}: {count}")

    print("\nCOMPARISON FLAGS")
    print("-" * 84)
    if comparison_notes:
        for note in comparison_notes:
            print(f"  {note}")
    else:
        print("  No task-total changes or progress decreases found.")

    print("\nPROCESSING WARNINGS")
    print("-" * 84)
    if warnings:
        for warning in warnings:
            print(f"  {warning}")
    else:
        print("  No processing warnings found.")

    print("\nUNKNOWN MODULES")
    print("-" * 84)
    if unknown_modules:
        for record in unknown_modules:
            print(
                f"  {record.source_type} | {record.sso} | "
                f"{record.work_order} | {record.source_module_name}"
            )
    else:
        print("  No unknown module descriptions found.")

    print("\nSNAPSHOT PROCESSING PASSED.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
