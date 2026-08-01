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
- Scope should be the subdirectory/script name, e.g. `feat(reporting): ...`
- Include the version bump in the message when applicable, e.g. `(v4.23)`

## Repository Layout

```
cohesity-automation/
├── alerts/
│   ├── alert_summary_report.py
│   ├── alert_to_csv.py
│   └── resolve_alerts.py
├── infrastructure/
│   ├── cluster_info_report.py
│   ├── node_status.py
│   └── upgrade_readiness.py
├── mcp/
│   └── cohesity_mcp_v2.py       # MCP server for Claude Desktop
├── policies/
│   ├── clone_policy.py
│   ├── list_policies.py
│   └── policy_compliance_report.py
├── protection/
│   ├── clone_protection_group.py
│   ├── pause_resume_groups.py
│   └── run_protection_group.py
├── recovery/
│   ├── list_snapshots.py
│   ├── recover_files.py
│   └── restore_vm.py
├── reporting/
│   ├── fortknox_vault_report.py
│   └── protection_group_report.py
├── storage/
│   ├── capacity_report.py
│   ├── list_views.py
│   └── storage_domain_report.py
├── utils/
│   ├── formatters.py            # Shared formatting helpers
│   └── cohesity_auth.py         # Authentication helpers
├── CHANGELOG.md                 # All version history
├── README.md                    # Project overview and usage reference
└── CLAUDE.md                    # This file
```

## Excel Formatting Standards

- **All Excel tabs must be set to 130% zoom.** Every sheet function must either call `auto_fit_columns(ws)` (which sets zoom to 130 internally) or explicitly set `ws.sheet_view.zoomScale = 130` inside a `try/except Exception: pass` block. Sheets with fixed column widths (About, Guide) that do not call `auto_fit_columns` must set zoom explicitly.

## General Guidelines

- Run a syntax check after every edit to catch errors before committing:
  ```bash
  python3 -c "import ast; ast.parse(open('<script>.py').read()); print('OK')"
  ```
- With **every version bump**, update these in the same commit:
  1. Script in-file `__version__` constant and docstring version block
  2. `CHANGELOG.md` — new version entry with date and bullet points
  3. Folder `README.md` and script `.md` file — if options, output format, or behaviour change
- Do not push to any branch other than `dev` without explicit user approval.
