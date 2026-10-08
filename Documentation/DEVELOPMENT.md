# Development Guide

## Local Setup

```powershell
cd "G:\Dev\PQS Dark Universe Report"
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Standard Validation

### Syntax check

```powershell
.\.venv\Scripts\python.exe -m py_compile .\Scripts\script_name.py
```

### Input scanner

```powershell
.\.venv\Scripts\python.exe .\Scripts\input_scanner.py
```

### Component checks

```powershell
.\.venv\Scripts\python.exe .\Scripts\roster_reader.py
.\.venv\Scripts\python.exe .\Scripts\qualification_reader.py
.\.venv\Scripts\python.exe .\Scripts\maximo_reader.py
.\.venv\Scripts\python.exe .\Scripts\module_result_processor.py
.\.venv\Scripts\python.exe .\Scripts\snapshot_manager.py
.\.venv\Scripts\python.exe .\Scripts\employee_summary_processor.py
```

Be aware that running `snapshot_manager.py` directly saves a snapshot. For read-only summary validation, use `employee_summary_processor.py`.

### Full report

```text
Run_PQS_Dark_Universe_Report.bat
```

Close the current output workbook before running.

## Adding a Module Mapping

Edit:

```text
Config\module_mapping.csv
```

Required columns:

```text
Source Module Name
Module Code
Venue
Active
```

Use the exact Maximo description. Do not change an existing mapping to a different module without confirming the business meaning.

## Adding an Exception

Edit the private file:

```text
Input\Exceptions\pqs_exceptions.xlsx
```

Columns:

```text
SSO
Module Code
Exception Type
Reason
Active
Effective Date
End Date
```

Supported exception type:

```text
Not Required
```

## Testing Guidance

Tests must use synthetic data. Recommended synthetic IDs:

```text
900000001
900000002
900000003
```

Do not copy production workbooks into `Tests`.

Useful test areas:

- SSO normalization
- Module mapping
- Maximo parent-child parsing
- Completed record selection
- Review Required threshold
- Snapshot comparisons
- Exception date ranges
- Venue completion
- Dynamic hidden columns

## Output Review Checklist

After Excel-generation changes, verify:

- Workbook opens normally in Microsoft Excel
- Only the seven approved sheets are visible
- Sheet order is correct
- Filters are active
- Headers are frozen
- SSOs and work orders remain text
- Percentages use percentage formatting
- Qualification-only modules have blank progress and work orders
- Progress pairs hide only when fully empty
- Work Order columns hide only when fully empty
- Dashboard counts reconcile with employee sheets
- Previous Output files were archived
- Snapshot was saved only after success

## Git Workflow

Check status:

```powershell
git status
```

The project preference is to combine add, commit, and push:

```powershell
git add <files>; git commit -m "<message>"; git push
```

Never use `git add -f` on ignored private-data folders.
