from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from module_result_processor import ModuleResult, load_settings, required_modules_for
from roster_reader import RosterPerson
from snapshot_manager import build_results_with_snapshot


PROJECT_ROOT = Path(__file__).resolve().parents[1]

VENUE_MODULES = {
    "Wolf": ["CUMOD0", "CUMOD1", "CUMOD2", "CUMOD3"],
    "Darkmoor": ["DUMOD0", "DUMOD1"],
    "Unchained": ["UCMOD0", "UCMOD1", "UCMOD2", "UCMOD3", "UCMOD4"],
    "Animation": ["UCMOD0", "UCMOD4"],
}

COMPLETE_RESULTS = {"Completed", "Not Required"}

ACTION_PRIORITY = [
    "ACK Error",
    "Task Total Changed",
    "Progress Decreased",
    "No Maximo Record",
    "Completed - Awaiting Maximo Status",
    "Review Required",
    "In Progress",
    "No Progress",
    "Not Started",
    "Completed",
]


@dataclass(frozen=True)
class VenueSummary:
    venue: str
    required_modules: tuple[str, ...]
    status: str
    work_orders: tuple[str, ...]
    incomplete_modules: tuple[str, ...]


@dataclass(frozen=True)
class EmployeeSummary:
    sso: str
    name: str
    position: str
    shift: str
    venue: str
    population: str
    roster_order: int
    pqs_status: str
    missing_pqs: str
    notes: str
    action_category: str
    review_modules: tuple[str, ...]
    venue_summaries: tuple[VenueSummary, ...]


def required_venues(person: RosterPerson) -> list[str]:
    if person.population == "Animation":
        return ["Animation"]

    if person.population in {"1st Shift", "2nd Shift"}:
        return ["Wolf", "Darkmoor", "Unchained"]

    if person.population == "3rd Shift Wolf":
        return ["Wolf", "Darkmoor"]

    if person.population == "3rd Shift Unchained":
        return ["Unchained"]

    if person.population == "Supervisors":
        if person.shift in {"1st", "2nd"} or person.venue == "Dark Universe":
            return ["Wolf", "Darkmoor", "Unchained"]
        if person.venue == "Wolf":
            return ["Wolf", "Darkmoor"]
        if person.venue == "Unchained":
            return ["Unchained"]

    raise ValueError(
        f"Unable to assign venues: {person.sso} | {person.population} | "
        f"{person.shift} | {person.venue}"
    )


def modules_for_venue(person: RosterPerson, venue: str) -> list[str]:
    if venue == "Animation":
        return list(VENUE_MODULES["Animation"])
    return list(VENUE_MODULES[venue])


def unique_in_order(values: list[str]) -> tuple[str, ...]:
    seen: set[str] = set()
    ordered: list[str] = []

    for value in values:
        if value and value not in seen:
            seen.add(value)
            ordered.append(value)

    return tuple(ordered)


def summarize_venue(
    person: RosterPerson,
    venue: str,
    results: dict[tuple[str, str], ModuleResult],
) -> VenueSummary:
    module_codes = modules_for_venue(person, venue)
    module_results = [results[(person.sso, code)] for code in module_codes]

    incomplete = [
        result.module_code
        for result in module_results
        if result.result not in COMPLETE_RESULTS
    ]

    work_orders = unique_in_order(
        [result.work_order for result in module_results if result.work_order]
    )

    status = "COMPLETE" if not incomplete else "INCOMPLETE"

    return VenueSummary(
        venue=venue,
        required_modules=tuple(module_codes),
        status=status,
        work_orders=work_orders,
        incomplete_modules=tuple(incomplete),
    )


def format_missing_pqs(venue_summaries: list[VenueSummary]) -> str:
    missing = [summary.venue for summary in venue_summaries if summary.status != "COMPLETE"]

    if not missing:
        return "Nothing - Fully Qualified"

    if len(missing) == 1:
        return f"Missing {missing[0]} PQS"

    if len(missing) == 2:
        return f"Missing {missing[0]} and {missing[1]} PQS"

    return "Missing " + ", ".join(missing[:-1]) + f", and {missing[-1]} PQS"


def result_action_label(result: ModuleResult) -> str | None:
    if result.result.startswith("ACK Error"):
        return "ACK Error"
    if result.result.startswith("Task Total Changed"):
        return "Task Total Changed"
    if result.result.startswith("Progress Decreased"):
        return "Progress Decreased"
    if result.result == "No Maximo Record":
        return "No Maximo Record"
    if result.result == "Completed - Awaiting Maximo Status":
        return "Completed - Awaiting Maximo Status"
    if result.result == "Review Required":
        return "Review Required"
    if result.result == "In Progress":
        return "In Progress"
    if result.result == "No Progress":
        return "No Progress"
    if result.result == "Not Started":
        return "Not Started"
    return None


def choose_action_category(module_results: list[ModuleResult]) -> str:
    labels = {
        label
        for result in module_results
        if (label := result_action_label(result)) is not None
    }

    for priority in ACTION_PRIORITY:
        if priority in labels:
            return priority

    return "Completed"


def build_notes(module_results: list[ModuleResult]) -> str:
    groups: list[tuple[str, list[str]]] = [
        ("ACK Error", []),
        ("Task Total Changed", []),
        ("Progress Decreased", []),
        ("No Maximo Record", []),
        ("Awaiting Maximo Status", []),
        ("Review Required", []),
        ("In Progress", []),
        ("No Progress", []),
        ("Not Started", []),
        ("Exception", []),
    ]
    group_lookup = {label: modules for label, modules in groups}

    detailed_notes: list[str] = []

    for result in module_results:
        if result.result.startswith("ACK Error"):
            group_lookup["ACK Error"].append(result.module_code)
        elif result.result.startswith("Task Total Changed"):
            group_lookup["Task Total Changed"].append(result.module_code)
            detailed_notes.append(result.note)
        elif result.result.startswith("Progress Decreased"):
            group_lookup["Progress Decreased"].append(result.module_code)
            detailed_notes.append(result.note)
        elif result.result == "No Maximo Record":
            group_lookup["No Maximo Record"].append(result.module_code)
        elif result.result == "Completed - Awaiting Maximo Status":
            group_lookup["Awaiting Maximo Status"].append(result.module_code)
        elif result.result == "Review Required":
            group_lookup["Review Required"].append(result.module_code)
        elif result.result == "In Progress":
            group_lookup["In Progress"].append(result.module_code)
        elif result.result == "No Progress":
            group_lookup["No Progress"].append(result.module_code)
        elif result.result == "Not Started":
            group_lookup["Not Started"].append(result.module_code)
        elif result.result == "Not Required":
            group_lookup["Exception"].append(result.module_code)
            if result.note:
                detailed_notes.append(result.note)

    note_parts: list[str] = []

    for label, modules in groups:
        if modules:
            note_parts.append(f"{label}: {', '.join(modules)}")

    for note in detailed_notes:
        if note and note not in note_parts:
            note_parts.append(note)

    if not note_parts:
        return "Nothing - Fully Qualified"

    return "; ".join(note_parts)


def build_employee_summaries(
    people: list[RosterPerson],
    results: dict[tuple[str, str], ModuleResult],
) -> list[EmployeeSummary]:
    settings = load_settings()
    summaries: list[EmployeeSummary] = []

    for person in people:
        required_modules = required_modules_for(person, settings)
        module_results = [results[(person.sso, code)] for code in required_modules]

        venue_summaries = [
            summarize_venue(person, venue, results)
            for venue in required_venues(person)
        ]

        pqs_status = (
            "Full PQS"
            if all(summary.status == "COMPLETE" for summary in venue_summaries)
            else "Not Full PQS"
        )

        review_modules = tuple(
            result.module_code
            for result in module_results
            if result.result == "Review Required"
        )

        summaries.append(
            EmployeeSummary(
                sso=person.sso,
                name=person.name,
                position=person.position,
                shift=person.shift,
                venue=person.venue,
                population=person.population,
                roster_order=person.roster_order,
                pqs_status=pqs_status,
                missing_pqs=format_missing_pqs(venue_summaries),
                notes=build_notes(module_results),
                action_category=choose_action_category(module_results),
                review_modules=review_modules,
                venue_summaries=tuple(venue_summaries),
            )
        )

    return summaries


def venue_work_orders(summary: EmployeeSummary, venue: str) -> str:
    for item in summary.venue_summaries:
        if item.venue == venue:
            return ", ".join(item.work_orders)
    return ""


def venue_status(summary: EmployeeSummary, venue: str) -> str:
    for item in summary.venue_summaries:
        if item.venue == venue:
            return item.status
    return ""


def main() -> int:
    (
        people,
        results,
        warnings,
        unknown_modules,
        previous_path,
        _snapshot_path,
        comparison_notes,
    ) = build_results_with_snapshot(save_snapshot=False)

    summaries = build_employee_summaries(people, results)

    print("=" * 92)
    print("PQS DARK UNIVERSE REPORT - EMPLOYEE AND VENUE SUMMARY")
    print("=" * 92)
    print(f"Employees: {len(summaries)}")
    print(
        "Comparison snapshot: "
        + (str(previous_path) if previous_path else "None - first-run baseline")
    )

    population_order = [
        "1st Shift",
        "2nd Shift",
        "3rd Shift Wolf",
        "3rd Shift Unchained",
        "Supervisors",
        "Animation",
    ]

    print("\nPOPULATION SUMMARY")
    print("-" * 92)
    for population in population_order:
        rows = [summary for summary in summaries if summary.population == population]
        full_count = sum(summary.pqs_status == "Full PQS" for summary in rows)
        print(
            f"  {population}: {len(rows)} employees | "
            f"Full PQS {full_count} | Not Full PQS {len(rows) - full_count}"
        )

    print("\nACTION CATEGORY COUNTS")
    print("-" * 92)
    action_counts: dict[str, int] = {}
    for summary in summaries:
        action_counts[summary.action_category] = action_counts.get(summary.action_category, 0) + 1

    for category in ACTION_PRIORITY:
        if category in action_counts:
            print(f"  {category}: {action_counts[category]}")

    print("\nVENUE SUMMARY")
    print("-" * 92)
    for venue in ["Wolf", "Darkmoor", "Unchained", "Animation"]:
        applicable = [
            summary
            for summary in summaries
            if any(item.venue == venue for item in summary.venue_summaries)
        ]
        complete = sum(venue_status(summary, venue) == "COMPLETE" for summary in applicable)
        print(
            f"  {venue}: {len(applicable)} employees | "
            f"Complete {complete} | Incomplete {len(applicable) - complete}"
        )

    print("\nNOT FULL PQS EMPLOYEES")
    print("-" * 92)
    not_full = [summary for summary in summaries if summary.pqs_status != "Full PQS"]

    if not_full:
        for summary in not_full:
            print(
                f"  {summary.population} | {summary.sso} | {summary.name} | "
                f"{summary.missing_pqs} | {summary.action_category} | {summary.notes}"
            )
    else:
        print("  Everyone is Full PQS.")

    print("\nPROCESSING CHECKS")
    print("-" * 92)
    print(f"  Module processing warnings: {len(warnings)}")
    print(f"  Unknown module descriptions: {len(unknown_modules)}")
    print(f"  Snapshot comparison flags: {len(comparison_notes)}")

    duplicate_ssos = len(summaries) - len({summary.sso for summary in summaries})
    print(f"  Duplicate employee summaries: {duplicate_ssos}")

    if warnings or unknown_modules or duplicate_ssos:
        print("\nEMPLOYEE SUMMARY COMPLETED WITH REVIEW ITEMS.")
        return 1

    print("\nEMPLOYEE AND VENUE SUMMARY PASSED.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
