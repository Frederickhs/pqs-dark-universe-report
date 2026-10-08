# GitHub Copilot Instructions

## Project Context

This repository is a local Python and Excel reporting application for Dark Universe PQS tracking. It processes private workforce data and produces a multi-sheet Excel report.

## Coding Expectations

- Use employee ID, SSO, or Labor ID as the primary key.
- Never propose name-based matching.
- Use `pathlib.Path` for file paths.
- Use type hints and focused functions.
- Prefer openpyxl for workbook parsing and generation.
- Preserve numeric-looking identifiers as text.
- Do not introduce cloud services, databases, APIs, or hosting unless explicitly requested.
- Do not introduce new packages unless the existing standard library, pandas, or openpyxl cannot reasonably solve the task.

## Privacy

Never suggest committing files from:

```text
Input/
Output/
Archive/
Logs/
```

Never place real employee data in source code, tests, documentation, examples, comments, or Git history. Use clearly synthetic test records.

## Business Rules

- Active qualification files are the official proof of module completion.
- Maximo child records provide current progress when no active qualification exists.
- Child completion requires `COMP` and 100 percent.
- Parent work-order status does not determine completion.
- `CLOSE` and `WCLOSE` are not official completion states.
- `ACK` requires a visible error.
- One or two remaining tasks means Review Required.
- Active `Not Required` exceptions remove the module from employee completion requirements.
- Full and Partial work orders merge by SSO and module.
- Unknown module descriptions must be flagged, not guessed.

## Snapshot Rules

- Compare against the newest successful snapshot from a prior date.
- Same-day reruns do not become the weekly comparison baseline.
- Increased tasks means In Progress.
- Unchanged tasks means No Progress.
- Decreased tasks means Progress Decreased - Review Data.
- Changed total tasks means Task Total Changed - Review Data.
- Save a snapshot only after successful workbook generation.

## Excel Output

Keep the visible sheets:

```text
Dash
1st Shift
2nd Shift
3rd Shift Wolf
3rd Shift Unchained
Supervisors
Animation
```

Use filters, frozen headers, readable widths, and status colors. Hide only fully empty progress, percentage, or work-order columns according to the existing generator rules.
