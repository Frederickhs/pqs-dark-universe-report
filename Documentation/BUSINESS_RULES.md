# Business Rules

## Primary Key

Employee ID, also called SSO or Labor ID, is the primary key. Names must never be used to join records.

## Population Sources

### 1st and 2nd Shift

Source: Line Schedule.

Requirements: all Wolf, Darkmoor, and Unchained modules.

### 3rd Shift Wolf

Source: Line Schedule.

Requirements: all Wolf and Darkmoor modules.

### 3rd Shift Unchained

Source: Line Schedule.

Requirements: all Unchained modules.

### Animation

Source: Animation Roster.

Defaults:

```text
Venue = Animation
Shift = 3rd
Position = Show Control
```

Requirements: UCMOD0 and UCMOD4.

### Supervisors

Source: Supervisor Roster.

- 1st or 2nd Shift, Dark Universe: all modules
- 3rd Shift Wolf: Wolf and Darkmoor
- 3rd Shift Unchained: Unchained

Supervisors appear only on the Supervisors sheet.

## Qualification Rule

An active qualification match for the employee SSO and module is official completion.

For qualification-completed modules:

- Result: Completed
- Progress: blank
- Percentage: blank
- Work Order: blank
- Effective Date retained internally

Inactive qualifications for a current roster employee are review items and do not count as active completion.

## Exception Rule

Supported exception type:

```text
Not Required
```

An active Not Required exception:

- Displays Not Required
- Removes the module from completion calculations
- Keeps the employee in dashboard population totals
- Displays the reason in Notes

Exceptions use SSO and Module Code.

## Maximo Child Rule

Only child rows determine module progress.

Official Maximo completion requires:

```text
Child Status = COMP
and
Child Percentage = 100 percent
```

Completed Tasks equal to Total Tasks may also confirm 100 percent, but COMP is still required.

Parent status does not override child results.

## Maximo Status Meanings

- `COMP`: eligible for official completion if progress is 100 percent
- `ACK`: technician action required
- `INPRG`: current work
- `CLOSE`: does not count as official completion
- `WCLOSE`: does not count as official completion

## Progress Statuses

- 0 completed with a valid child: No Progress
- More than 2 remaining: In Progress before weekly comparison
- 1 or 2 remaining: Review Required
- 0 remaining without COMP: Completed - Awaiting Maximo Status
- 0 remaining with COMP: Completed

The final two tasks represent Trainer Review and Supervisor Review. The source does not identify which signature is missing, so the report uses the general label Review Required.

## Multiple Work Orders

Full and Partial work orders merge.

Selection order:

1. Any officially completed child satisfies the module.
2. If several completed children exist, use the newest completed record.
3. If none are completed, use the highest percentage.
4. If percentages tie, use the newest record.

## Missing Data

- SSO absent from applicable Maximo but qualifications may exist: qualification results still apply.
- SSO has Maximo records but a required child is absent: Not Started.
- SSO has no applicable Maximo records for unresolved modules: No Maximo Record.
- Unknown child description: visible mapping warning; do not guess.

## Weekly Comparison

Use the newest successful snapshot from a prior calendar date.

- Increase: In Progress
- No increase: No Progress
- Decrease: Progress Decreased - Review Data
- Total task change: Task Total Changed - Review Data

Special statuses override weekly movement:

1. ACK Error
2. Task Total Changed
3. Progress Decreased
4. No Maximo Record
5. Completed
6. Completed - Awaiting Maximo Status
7. Review Required
8. In Progress
9. No Progress

## Venue Completion

A venue is COMPLETE only when every required module is Completed or Not Required.

## Overall Completion

- Full PQS: every required venue is COMPLETE
- Not Full PQS: one or more required venues are INCOMPLETE

## Missing PQS Text

Examples:

```text
Missing Wolf PQS
Missing Wolf and Darkmoor PQS
Missing Wolf, Darkmoor, and Unchained PQS
Nothing - Fully Qualified
```

## Dashboard

Includes:

- Population totals
- Venue totals
- Completion percentages
- Review employee and module counts
- Action-needed categories
- Employee action list
- Overall excluding Animation
- Overall including Animation

The No Maximo Record metric column appears only when at least one current result requires it.
