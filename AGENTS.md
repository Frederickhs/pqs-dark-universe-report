# AGENTS.md

Repository instructions for Codex and other coding agents.

## Goal

Maintain a local Python pipeline that generates `Output\PQS Dark Universe Report.xlsx` from private roster, qualification, exception, Maximo, and snapshot data.

## Required Reading

Before editing:

1. `README.md`
2. `Documentation/BUSINESS_RULES.md`
3. `Documentation/ARCHITECTURE.md`
4. `Documentation/DEVELOPMENT.md`

## Safety and Privacy

- Do not commit anything under `Input`, `Output`, `Archive`, or `Logs`.
- Do not include real names, SSOs, work-order numbers, or report extracts in tests or documentation.
- Do not modify files in `G:\Dev\Qual_Tracker\Input\Qualification`.
- Use employee ID as the only cross-source identity key.
- Never match by name.

## Architecture Contract

Pipeline order:

```text
Input validation
  -> roster population
  -> qualification lookup
  -> Maximo child parsing
  -> exception and module results
  -> prior-date snapshot comparison
  -> employee and venue summaries
  -> Excel generation
  -> successful snapshot save
```

Do not save a snapshot before successful report generation.

## Core Rules

- Qualification completion takes priority over Maximo work orders.
- Maximo completion requires child `Status = COMP` and 100 percent.
- Parent status is informational only.
- `CLOSE` and `WCLOSE` do not count as completed.
- `ACK` is a visible action error.
- One or two remaining tasks means Review Required.
- Active Not Required exceptions remove the module from completion calculations.
- Full and Partial work orders merge by SSO and module.
- Unknown module descriptions must not be guessed.

## Output Contract

Visible sheets:

```text
Dash
1st Shift
2nd Shift
3rd Shift Wolf
3rd Shift Unchained
Supervisors
Animation
```

Dynamic columns:

- Hide module Progress and Percentage only when the entire sheet has no progress for that module.
- Hide a venue Work Order column only when all rows are blank.
- Never hide the module result or venue Full PQS columns.

## Validation Commands

Syntax:

```powershell
.\.venv\Scripts\python.exe -m py_compile .\Scripts\<script>.py
```

Inputs:

```powershell
.\.venv\Scripts\python.exe .\Scripts\input_scanner.py
```

Full workflow:

```text
Run_PQS_Dark_Universe_Report.bat
```

Close the generated workbook in Excel before running the workflow.

## Change Discipline

- Make the smallest focused change.
- Do not alter a confirmed business rule without explicit user direction.
- Do not redesign file locations without explicit approval.
- Keep the current dependency set unless necessary.
- Add synthetic tests for parsing or status logic when practical.
- Explain assumptions and edge cases in the pull request or commit message.

## Git Convention

When presenting a commit command to the user, combine all operations on one line:

```powershell
git add <files>; git commit -m "<message>"; git push
```

