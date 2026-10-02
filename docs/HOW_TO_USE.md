# How to Use

## A. Prepare your Excel file

Start with `DSB_Input_Template.xlsx` and save a private working copy, for example:

```text
NMRY_DSB_Input.xlsx
```

Do not overwrite the public template with your real alliance data.

### Players

Enter one player per row:

| Name | Team Availability | THP | DA/VS | Status |
|---|---|---:|---:|---|
| Player 001 | A+B | 177.5 | 12345 | |

`Team Availability` must indicate whether the player can join Team A, Team B, or both.

### Blacklist

Use:

| Name | Reason |
|---|---|
| Player 099 | Blacklist DSB next week |

Only entries with the reason `Blacklist DSB next week` are treated as active DSB blacklist entries by the current script.

### DSB Priority

Add eligible names to the priority queue. The script maintains this queue between runs. Selected players are removed; eligible players not selected can be added to the queue.

## B. Configure the run

Edit the `Config` sheet. A typical setup is:

```text
Mode = SAVING
Random Seed = 42
Max Team Size = 30
Main Players per Team = 20
Reserve Players per Team = 10
DSB Priority Count = 0
```

Use `ALL-IN` when the Top 40 DA/VS stage should be active.

## C. Run the script

```bash
python run_dsb_masterdata.py NMRY_DSB_Input.xlsx
```

The script modifies the supplied workbook in place. Make a backup before each production run.

## D. Review the result

Open `DSB Assignment` and check:

- Team A / Team B
- Main / Reserve
- T1–T4
- THP
- Availability
- Selection Reason

The `Players.Status` column is also updated with values such as:

- `Top 10 THP`
- `Top 40 DA/VS`
- `DSB Priority`
- `Spin`
- `Not Selected`
- `Blacklisted`

## E. Reproducibility

The random generator uses the configured seed. With the same input, rules and seed, the selection is intended to be reproducible.

## Recommended operating model

```text
Collect / paste Last War data
        ↓
Update Excel
        ↓
Run Python
        ↓
Review DSB Assignment
        ↓
Make any approved manual adjustment
        ↓
Publish final DSB roster
```
