# Last War DSB Assignment Tool

Python + Excel automation for assigning Last War DSB teams from a structured Excel input file.

## What it does

The script is designed around a simple workflow:

`Excel input → Python script → DSB Assignment → human review → final roster`

It automates the repetitive parts of DSB selection while keeping Excel as the human-friendly input and review layer.

### Selection logic

The current script implements these stages:

1. Exclude active DSB blacklist entries.
2. Select the **Top 10 THP** players.
3. In **ALL-IN** mode, select the next **Top 40 DA/VS** players.
4. Select configurable players from the persistent **DSB Priority** queue.
5. Fill remaining slots by seeded **Spin** selection while respecting team availability.
6. Split selected players between Team A and Team B within configured capacity limits and balance THP.
7. Assign Main/Reserve and T1–T4 positions.
8. Record a selection reason for every selected player and update `Players.Status`.
9. Rewrite the priority queue so selected players leave the queue and unselected eligible players can enter it.

The exact rules are implemented in `run_dsb_masterdata.py`.

## Repository structure

```text
lastwar-dsb-assignment/
├── run_dsb_masterdata.py
├── DSB_Input_Template.xlsx
├── requirements.txt
├── README.md
├── LICENSE
├── .gitignore
├── CONTRIBUTING.md
└── docs/
    ├── HOW_TO_USE.md
    └── GITHUB_SETUP.md
```

## Requirements

- Python 3.10+ recommended
- Microsoft Excel or another application that can edit `.xlsx` files
- `openpyxl`

## Quick start

### 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/lastwar-dsb-assignment.git
cd lastwar-dsb-assignment
```

### 2. Create a virtual environment

Windows:

```bash
python -m venv .venv
.venv\Scripts\activate
```

macOS/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Prepare Excel

Copy `DSB_Input_Template.xlsx` and enter your own alliance data. **Do not publish real member data unless you have permission to do so.**

### 5. Run

```bash
python run_dsb_masterdata.py DSB_Input_Template.xlsx
```

The workbook is updated in place and a `DSB Assignment` sheet is generated.

## Excel input

The workbook contains these important sheets:

| Sheet | Purpose |
|---|---|
| `Players` | Player name, THP, DA/VS and Team Availability |
| `Blacklist` | Players who must be excluded |
| `Config` | Selection mode, seed, team limits and priority count |
| `DSB Priority` | Persistent priority queue |
| `DSB Assignment` | Generated output; created by the script |

### Players columns

- `Name` — player display name
- `Team Availability` — `A`, `B`, or `A+B`
- `THP` — main squad power used for sorting/balancing
- `DA/VS` — DA/VS value used in All-In selection
- `Status` — written by the script

### Config

The template includes:

- `Mode`: `SAVING` or `ALL-IN`
- `Random Seed`: deterministic random selection seed
- `Max Team Size`: maximum selected players per team
- `Main Players per Team`: main roster size
- `Reserve Players per Team`: reserve capacity
- `DSB Priority Count`: number taken from the priority queue per run

## Important: human review

This tool is intended to automate selection and balancing, **not to remove human oversight**. Review the generated `DSB Assignment` before publishing the final roster.

## Open-source repository setup

See [`docs/GITHUB_SETUP.md`](docs/GITHUB_SETUP.md) for GitHub CLI, Git, and GitHub Desktop publishing instructions.

## Privacy

The public repository should contain only the anonymized template. Alliance member names, screenshots, private rankings, identifiers, or other personal data should remain outside the public repository unless the relevant people have authorized publication.

## Disclaimer

This is an independent community tool. It is not affiliated with, endorsed by, or sponsored by the developer or publisher of Last War.
