# Publish this project to GitHub

## Option 1 — GitHub CLI

From the repository directory:

```bash
gh auth login
gh repo create lastwar-dsb-assignment --public --source=. --remote=origin --push
```

GitHub CLI can create a repository from an existing local repository and push it.

## Option 2 — GitHub website + Git

1. Create a new empty repository on GitHub named `lastwar-dsb-assignment`.
2. Do **not** add another README, license, or `.gitignore` when creating the empty repository.
3. From this local project directory:

```bash
git remote add origin https://github.com/YOUR_USERNAME/lastwar-dsb-assignment.git
git push -u origin main
```

4. Replace `YOUR_USERNAME` with your GitHub username.

## Option 3 — GitHub Desktop

Add the existing local repository to GitHub Desktop and choose **Publish repository**. For a public project, clear **Keep this code private**.

## Before publishing

Check:

```bash
git status
git ls-files
```

Make sure your real `NMRY_MasterData.xlsx`, alliance screenshots, private member lists, generated output files, passwords, tokens, or other private information are not tracked. The repository `.gitignore` already excludes common private workbook names.

## Suggested repository description

> Excel + Python automation for Last War DSB team selection, balancing, priority queues, blacklist handling, and reproducible roster generation.

## Suggested topics

`python` `openpyxl` `excel` `automation` `last-war` `dsb` `team-balancing` `gaming-tools`
