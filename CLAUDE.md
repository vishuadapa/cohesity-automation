# Claude Code — Project Rules

## Branch Policy

- **Always work on `dev`**. Never commit or push directly to `main`.
- If the current branch is `main`, switch to `dev` before making any changes:
  ```bash
  git checkout dev
  ```
- When a version is ready to release to `main`, ask the user to confirm before merging.
- Merging dev → main should be a fast-forward or squash merge with a clean commit message summarizing the release.

## Commit Style

- Use conventional commits: `feat`, `fix`, `refactor`, `docs`, `chore`
- Scope should be the subdirectory/script name, e.g. `feat(health_check): ...`
- Include the version bump in the message when applicable, e.g. `(v1.34)`

## Repository Layout

```
cohesity-automation/
├── health_check/
│   ├── health_check_report.py   # Main health check script (currently v1.47)
│   └── health_check_report.md   # Documentation
├── utils/
│   ├── formatters.py            # Shared formatting helpers
│   └── cohesity_auth.py         # Authentication helpers
├── CHANGELOG.md                 # All version history
├── README.md                    # Project overview and usage reference
└── CLAUDE.md                    # This file
```

## General Guidelines

- Run `python3 -c "import ast; ast.parse(open('health_check/health_check_report.py').read()); print('OK')"` after every edit to catch syntax errors before committing.
- With **every version bump**, update all **five** of these in the same commit:
  1. `README.md` — sheet count, filename examples, feature list, requirements
  2. `CHANGELOG.md` — new version entry with date and bullet points
  3. `health_check/health_check_report.md` — full sheet reference and version history
  4. In-file version header — `__version__` constant and the docstring version block
  5. **`_sheet_guide()` in `health_check_report.py`** — keep the Guide tab in sync:
     - Adding a new sheet → add a row to the `sheets_info` list in `_sheet_guide()`
     - Removing a sheet → remove its row from `sheets_info`
     - Changing a scoring model (security, ransomware, risk, health) → update the corresponding section in `_sheet_guide()`
     - Adding/changing recommendation priorities or color usage → update those sections
     - The Guide tab must always accurately reflect the current state of the workbook
- Do not push to any branch other than `dev` without explicit user approval.
