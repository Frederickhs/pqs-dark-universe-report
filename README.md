# PQS Dark Universe Report

Local Python reporting solution that converts roster, qualification, exception, and IBM Maximo training data into a redesigned Excel PQS report for Dark Universe Technical Services.

## Purpose

The project replaces a manually maintained Excel report with a repeatable local workflow. The generated workbook shows PQS status, active Maximo progress, venue completion, review readiness, exceptions, and dashboard metrics.

Employee ID, also called SSO or Labor ID, is the primary key throughout the project. Names are display values only.

## Current Output

The generated workbook is:

```text
Output\PQS Dark Universe Report.xlsx
```

Visible sheets:

1. `Dash`
2. `1st Shift`
3. `2nd Shift`
4. `3rd Shift Wolf`
5. `3rd Shift Unchained`
6. `Supervisors`
7. `Animation`

## Quick Start

1. Update the required input files.
2. Close the current output workbook if it is open in Excel.
3. Double-click:

```text
Run_PQS_Dark_Universe_Report.bat
```

4.The runner validates the inputs, generates the workbook, saves a successful snapshot, and opens the `Output` folder.

## Data Sources

### Local project inputs

```text
Input\
â”œâ”€â”€ Maximo\
â”‚   â”œâ”€â”€ Operations\
â”‚   â””â”€â”€ Animation\
â”œâ”€â”€ Line Schedule\
â”‚   â””â”€â”€ Line Schedule.xlsx
â”œâ”€â”€ Rosters\
â”‚   â””â”€â”€ Animation Roster.xlsx
â”œâ”€â”€ Supervisors\
â”‚   â””â”€â”€ Supervisor Roster.xlsx
â””â”€â”€ Exceptions\
    â””â”€â”€ pqs_exceptions.xlsx
```

The Operations and Animation Maximo files keep their original download names. The newest `.xlsx` file in each Maximo folder is selected.

### Shared qualification source

Qualification files are read directly from:

```text
G:\Dev\Qual_Tracker\Input\Qualification
```

The shared qualification files are not copied, renamed, moved, or archived by this project.

## Population Rules

- `1st Shift` and `2nd Shift` come from their respective Line Schedule worksheets.
- `3rd Shift Wolf` and `3rd Shift Unchained` come from their respective Line Schedule worksheets.
- Position is derived from Line Schedule sections such as Lead, Show Controls, and Mechanics.
- Animation comes from `Animation Roster.xlsx`.
- Supervisors come from `Supervisor Roster.xlsx`.
- Animation and Supervisor SSOs are excluded from regular shift populations.
- Roster order is preserved.

## Required Modules

### Dark Universe operational population

Used by 1st Shift, 2nd Shift, and 1st or 2nd Shift Supervisors:

```text
CUMOD0 CUMOD1 CUMOD2 CUMOD3
DUMOD0 DUMOD1
UCMOD0 UCMOD1 UCMOD2 UCMOD3 UCMOD4
```

### 3rd Shift Wolf

```text
CUMOD0 CUMOD1 CUMOD2 CUMOD3
DUMOD0 DUMOD1
```

### 3rd Shift Unchained

```text
UCMOD0 UCMOD1 UCMOD2 UCMOD3 UCMOD4
```

### Animation

```text
UCMOD0 UCMOD4
```

### Supervisors

- Shift `1st` or `2nd`, Venue `Dark Universe`: all modules.
- Shift `3rd`, Venue `Wolf`: all Wolf and Darkmoor modules.
- Shift `3rd`, Venue `Unchained`: all Unchained modules.

## Result Priority

For each employee and required module:

1. Active `Not Required` exception
2. Active qualification report match
3. Selected Maximo child record
4. Missing module or missing Maximo result

An active qualification is the official proof of completion. Completed Maximo work orders may disappear into the Maximo archive, so qualification reports take priority over current work-order data.

## Maximo Completion Rules

A child module is officially completed only when:

- Child status is `COMP`
- Child percentage is 100 percent or completed tasks equal total tasks

Parent work-order status does not control completion.

Other important statuses:

- `ACK`: technician action required
- `CLOSE` and `WCLOSE`: do not count as official completion
- One or two tasks remaining: `Review Required`
- Zero tasks remaining without `COMP`: `Completed - Awaiting Maximo Status`
- Zero completed tasks with a valid child: `No Progress`

Full and Partial work orders are merged by SSO and module code.

## Weekly Snapshot Logic

Successful runs store private CSV snapshots in:

```text
Archive\Snapshots
```

The current run compares against the newest successful snapshot from a prior calendar date.

- Completed tasks increased: `In Progress`
- Completed tasks unchanged: `No Progress`
- Completed tasks decreased: `Progress Decreased - Review Data`
- Total Tasks changed: `Task Total Changed - Review Data`

Same-day reruns do not use another same-day run as the comparison baseline.

## Output Archiving

Before a new workbook is created, all files currently in `Output` are moved to a timestamped folder under:

```text
Archive\Outputs
```

Input files remain untouched.

## Dynamic Excel Columns

- Module Progress and Percentage columns are hidden when nobody on that sheet has Maximo progress for the module.
- Both columns remain visible when at least one employee has progress.
- A venue Work Order column is hidden when every employee on the sheet has a blank work order.
- Module result and venue Full PQS columns remain visible.

## Project Structure

```text
PQS Dark Universe Report\
â”œâ”€â”€ .github\
â”‚   â””â”€â”€ copilot-instructions.md
â”œâ”€â”€ Config\
â”‚   â”œâ”€â”€ module_mapping.csv
â”‚   â””â”€â”€ project_settings.json
â”œâ”€â”€ Input\                 # Private, excluded from Git
â”œâ”€â”€ Output\                # Generated, excluded from Git
â”œâ”€â”€ Archive\               # Private history, excluded from Git
â”œâ”€â”€ Logs\                  # Private logs, excluded from Git
â”œâ”€â”€ Scripts\
â”‚   â”œâ”€â”€ input_scanner.py
â”‚   â”œâ”€â”€ roster_reader.py
â”‚   â”œâ”€â”€ qualification_reader.py
â”‚   â”œâ”€â”€ maximo_reader.py
â”‚   â”œâ”€â”€ module_result_processor.py
â”‚   â”œâ”€â”€ snapshot_manager.py
â”‚   â”œâ”€â”€ employee_summary_processor.py
â”‚   â””â”€â”€ report_generator.py
â”œâ”€â”€ Tests\
â”œâ”€â”€ Documentation\
â”œâ”€â”€ .gitignore
â”œâ”€â”€ AGENTS.md
â”œâ”€â”€ CLAUDE.md
â”œâ”€â”€ README.md
â”œâ”€â”€ requirements.txt
â””â”€â”€ Run_PQS_Dark_Universe_Report.bat
```

## Python Environment

Current validated environment:

```text
Python 3.14.8
pandas 3.0.6
openpyxl 3.1.5
```

Create and install the environment:

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Manual Commands

Validate inputs:

```powershell
.\.venv\Scripts\python.exe .\Scripts\input_scanner.py
```

Generate the report:

```powershell
.\.venv\Scripts\python.exe .\Scripts\report_generator.py
```

## Privacy and Git

Real employee data must never be committed.

The following are excluded from Git:

```text
.venv/
Input/
Output/
Archive/
Logs/
```

Do not add real SSO values, names, Maximo exports, qualification exports, rosters, snapshots, generated workbooks, or private logs to documentation, examples, tests, issues, or commits.

## Repository

```text
https://github.com/Frederickhs/pqs-dark-universe-report.git
```

Primary branch:

```text
main
```

