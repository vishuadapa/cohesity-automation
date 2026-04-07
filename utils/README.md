# Utils

Shared helper modules imported by scripts in other folders. Keeps auth, formatting, and common API patterns in one place so individual scripts stay focused.

## Modules

| Module | README | Description |
|--------|--------|-------------|
| `cohesity_auth.py` | [docs](cohesity_auth.md) | Unified auth handler for both direct cluster and Helios API key modes |
| `formatters.py` | [docs](formatters.md) | Shared helpers: `usecs_to_datetime`, `bytes_to_gb`, `bytes_to_tb`, `format_duration`, `style_worksheet`, `auto_fit_columns` |

## Usage

These modules can be imported by scripts in other folders:

```python
from utils.cohesity_auth import get_api_key, get_helios_clusters, make_helios_headers
from utils.formatters import usecs_to_datetime, bytes_to_gb, style_worksheet
```

All individual scripts also contain the same auth and formatting logic inline for portability — the utils modules are provided for projects that prefer a shared dependency approach.
