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
│   ├── health_check_report.py   # Main health check script (currently v1.34)
│   └── health_check_report.md   # Documentation
├── utils/
│   ├── formatters.py            # Shared formatting helpers
│   └── cohesity_auth.py         # Authentication helpers
├── CHANGELOG.md                 # All version history
└── CLAUDE.md                    # This file
```

## General Guidelines

- Run `python3 -c "import ast; ast.parse(open('health_check/health_check_report.py').read()); print('OK')"` after every edit to catch syntax errors before committing.
- Update `CHANGELOG.md` and the `.md` doc file alongside every version bump.
- Do not push to any branch other than `dev` without explicit user approval.
