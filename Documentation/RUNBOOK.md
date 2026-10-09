# Operations Runbook

## Weekly Run

1. Update the Line Schedule if staffing changed.
2. Update the Animation Roster if membership changed.
3. Update the Supervisor Roster if assignment changed.
4. Update `pqs_exceptions.xlsx` if an exception changed.
5. Place the newest Operations Maximo report in:

```text
Input\Maximo\Operations
```

6. Place the newest Animation Maximo report in:

```text
Input\Maximo\Animation
```

7. Confirm the shared qualification reports are current in:

```text
G:\Dev\Qual_Tracker\Input\Qualification
```

8. Close `PQS Dark Universe Report.xlsx` in Excel.
9. Double-click:

```text
Run_PQS_Dark_Universe_Report.bat
```

10. Review the workbook in `Output`.

## What the Runner Does

1. Confirms the virtual environment exists.
2. Runs the input scanner.
3. Stops if required files are missing.
4. Archives existing Output files.
5. Processes roster, qualification, exception, Maximo, and snapshot data.
6. Generates the workbook.
7. Reopens the workbook to verify the expected sheets exist.
8. Saves the successful snapshot.
9. Opens the Output folder.

## Common Problems

### Excel file is open

Close the current output workbook and run again.

### Missing Maximo report

Confirm at least one `.xlsx` file exists in each Maximo folder. Temporary files beginning with `~$` are ignored.

### Unknown child module

Add the exact child description to `Config\module_mapping.csv` after confirming the proper controlled module code.

### No Maximo Record

Confirm the SSO exists in the correct report and that the correct Maximo export was downloaded. A valid qualification can still complete a module without a current work order.

### Task Total Changed

Review the Maximo job plan. The report intentionally flags denominator changes because an active job plan should not normally change after progress begins.

### Progress Decreased

Review the Maximo work order or job-plan data. The report does not silently accept a lower completed-task count.

### Snapshot does not compare

The system compares only to snapshots from a prior calendar date. Same-day reruns intentionally use the earlier prior-date baseline.

## Private Data Handling

Do not email, upload, or commit Input, Output, Archive, Logs, or shared qualification exports unless approved through the appropriate business process.

