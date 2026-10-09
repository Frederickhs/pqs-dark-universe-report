# CLAUDE.md

Instructions for Claude Code when working in this repository.

## Project Summary

This repository generates a local Excel PQS report for Dark Universe Technical Services. It combines controlled rosters, active qualification reports, IBM Maximo work-order progress, individual exceptions, and prior-run snapshots.

The project is production-oriented. Preserve confirmed business rules and do not infer new ones from sample data.

## Non-Negotiable Rules

1. Employee ID, called SSO or Labor ID, is the primary key.
2. Never match employees by name.
3. Never commit real employee data.
4. Never edit, rename, move, or archive input reports.
5. Never write into the shared qualification folder.
6. Do not treat parent work-order status as module completion.
7. Do not count `CLOSE` or `WCLOSE` as completed.
8. Do not guess unknown module descriptions.
9. A successful snapshot is saved only after the Excel report is generated successfully.
10. Preserve the local, no-hosting architecture unless the user explicitly requests a change.

## Privacy Boundary

Never inspect, print, commit, or reproduce private data unless required for a local validation run. Private paths include:

```text
Input/
Output/
Archive/
Logs/
G:\Dev\Qual_Tracker\Input\Qualification
```

When creating tests, use synthetic SSOs and fictional names. State clearly that test data is synthetic.

## Current Business Rules

### Populations

- Line Schedule controls 1st Shift, 2nd Shift, 3rd Shift Wolf, and 3rd Shift Unchained.
- Animation Roster controls Animation.
- Supervisor Roster controls Supervisors.
- Animation and Supervisors are exclusive populations.
- Roster order is preserved.

### Required modules

- Dark Universe: CUMOD0-3, DUMOD0-1, UCMOD0-4
- Wolf: CUMOD0-3 and DUMOD0-1
- Unchained: UCMOD0-4
- Animation: UCMOD0 and UCMOD4

### Module source priority

1. Active `Not Required` exception
2. Active qualification
3. Maximo child record
4. Missing module or missing Maximo state

### Completion

A Maximo child is officially completed only when status is `COMP` and progress is 100 percent.

- `ACK`: visible error
- `CLOSE` and `WCLOSE`: not official completion
- One or two remaining tasks: Review Required
- Zero remaining tasks without COMP: Completed - Awaiting Maximo Status
- Zero completed tasks: No Progress

### Multiple work orders

- Merge Full and Partial work orders.
- Any valid historical completed child permanently satisfies the module for current processing.
- If no completed record exists, use highest percentage.
- If percentages tie, use the newest record.
- If multiple completed records exist, use the newest completed record.

### Snapshots

Compare against the newest successful snapshot from a prior date.

- Increased completed tasks: In Progress
- Unchanged completed tasks: No Progress
- Decreased completed tasks: Progress Decreased - Review Data
- Changed total tasks: Task Total Changed - Review Data

## Script Responsibilities

- `input_scanner.py`: verify required inputs and folders
- `roster_reader.py`: build exclusive report populations
- `qualification_reader.py`: load active qualification matches
- `maximo_reader.py`: parse parent and child rows and map child descriptions
- `module_result_processor.py`: calculate initial module results
- `snapshot_manager.py`: compare and save weekly snapshots
- `employee_summary_processor.py`: calculate venue and employee summaries
- `report_generator.py`: archive previous outputs and generate Excel

Keep responsibilities separated. Prefer adding a focused helper over duplicating logic in multiple scripts.

## Modification Workflow

Before changing code:

1. Read `README.md`.
2. Read `Documentation/BUSINESS_RULES.md`.
3. Identify the script that owns the behavior.
4. Check whether the change affects snapshots, status priority, privacy, or output compatibility.

After changing code:

```powershell
.\.venv\Scripts\python.exe -m py_compile .\Scripts\changed_script.py
.\.venv\Scripts\python.exe .\Scripts\input_scanner.py
```

If the change affects processing, run the relevant reader or processor. If the change affects output, close Excel and run the BAT.

## Excel Rules

- Use openpyxl.
- Preserve the seven visible sheets.
- Use `ws.auto_filter.ref`, not openpyxl Table objects.
- Preserve SSO and work-order numbers as text.
- Keep filters and frozen headers.
- Hide Progress and Percentage pairs only when nobody on the sheet has actual progress.
- Hide a Work Order column only when the entire column is blank.
- Keep Full PQS columns visible.
- Archive existing Output files before creating the current workbook.

## Coding Style

- Use type hints.
- Use `pathlib.Path`.
- Keep normalization functions explicit.
- Fail clearly on missing required structure.
- Continue with visible warnings for unknown module descriptions, unless report generation would be unsafe.
- Do not silently repair business data.
- Avoid introducing new dependencies without a clear need.

## Git

Repository:

```text
https://github.com/Frederickhs/pqs-dark-universe-report.git
```

Branch:

```text
main
```

When offering save commands, provide add, commit, and push on one line:

```powershell
git add <files>; git commit -m "<message>"; git push
```

