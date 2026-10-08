# Architecture

## Overview

PQS Dark Universe Report is a local, file-based Python pipeline. It does not connect directly to IBM Maximo and does not use a hosted database.

## Data Flow

```text
Line Schedule ---------------------------+
Animation Roster ------------------------+
Supervisor Roster -----------------------+--> Employee populations
                                         |
Shared Qualification Reports -----------+--> Completed module evidence
                                         |
Operations Maximo Report ----------------+--> Current operational progress
Animation Maximo Report -----------------+--> Current animation progress
                                         |
PQS Exceptions --------------------------+--> Individual requirement overrides
                                         |
Prior-date Snapshot ---------------------+--> Weekly movement comparison
                                         |
                                         v
                               Module result processor
                                         |
                                         v
                               Employee and venue summary
                                         |
                                         v
                               Excel report generator
                                         |
                         +---------------+---------------+
                         |                               |
                         v                               v
              Current Output Workbook          Successful CSV Snapshot
```

## Identity Model

SSO is the master key across all sources. Names are display values and must not be used for matching.

## Processing Layers

### Input validation

`input_scanner.py` verifies paths, files, Maximo reports, shared qualifications, and required folders.

### Roster layer

`roster_reader.py` creates exclusive populations and preserves roster order.

### Qualification layer

`qualification_reader.py` loads qualification evidence only for current report-population SSOs.

### Maximo layer

`maximo_reader.py` reconstructs parent-child relationships and maps child descriptions through `Config/module_mapping.csv`.

Parents without an SSO are ignored because names are not a valid identity key.

### Result layer

`module_result_processor.py` applies exceptions, qualifications, Maximo selection, and initial status rules.

### Snapshot layer

`snapshot_manager.py` compares current module progress with the newest successful snapshot from a prior date.

### Summary layer

`employee_summary_processor.py` calculates venue completion, Full PQS, Missing PQS, Notes, review counts, and dashboard categories.

### Presentation layer

`report_generator.py` archives prior outputs, creates the seven-sheet workbook, applies formatting and dynamic column visibility, validates the workbook structure, and saves the successful snapshot.

## Configuration

### `Config/project_settings.json`

Controls:

- Project name
- Output filename
- Shared qualification folder
- Review task thresholds
- Required module groups
- Visible sheet list

### `Config/module_mapping.csv`

Maps Maximo child descriptions to controlled module codes.

Unknown descriptions are never guessed.

## Output Lifecycle

1. Existing Output files are moved to a timestamped archive folder.
2. The new workbook is generated.
3. The workbook structure is reopened and checked.
4. A successful snapshot is written.
5. The BAT opens the Output folder.

## Extension Points

Safe future additions include:

- Synthetic unit tests
- More module mapping variants
- Additional dashboard metrics
- A run-summary log
- Controlled new populations

Any future population should still use SSO as the master key and should be represented through configuration or a dedicated roster reader.
