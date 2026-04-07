# formatters.py

Shared formatting helpers module for Cohesity scripts. Provides utility functions for timestamp conversion, unit conversion, duration formatting, and Excel styling (Cohesity green branding).

Designed to be imported by other scripts to eliminate repeated formatting logic across reports.

Required packages:
```bash
pip install openpyxl
```

**Version:** 1.0

---

## Usage

Import in your scripts:

```python
from utils.formatters import (
    usecs_to_datetime,
    bytes_to_gb,
    bytes_to_tb,
    format_duration,
    style_worksheet,
    auto_fit_columns,
    COHESITY_GREEN,
)
```

---

## Functions

### `usecs_to_datetime(usecs: int, tz=None) -> str`

Converts a microsecond epoch timestamp to a human-readable local datetime string.

**Parameters:**
- `usecs` — Microsecond epoch timestamp (int). Returns `""` if `0` or `None`.
- `tz` — Optional `ZoneInfo` timezone object. If `None`, uses UTC.

**Returns:** String in format `"YYYY-MM-DD HH:MM:SS"`

**Example:**
```python
from zoneinfo import ZoneInfo

# UTC (no timezone):
dt = usecs_to_datetime(1712445600000000)
# Output: "2024-04-07 00:00:00"

# America/New_York:
tz = ZoneInfo("America/New_York")
dt = usecs_to_datetime(1712445600000000, tz=tz)
# Output: "2024-04-06 20:00:00"

# Falsy value:
dt = usecs_to_datetime(0)
# Output: ""
```

---

### `bytes_to_gb(b) -> float`

Converts bytes to decimal GB (÷ 1,000,000,000) with 4 decimal places.

Returns `0.0` for `None` or `0`.

**Example:**
```python
gb = bytes_to_gb(1_000_000_000)
# Output: 1.0

gb = bytes_to_gb(1_500_000_000)
# Output: 1.5

gb = bytes_to_gb(None)
# Output: 0.0
```

---

### `bytes_to_tb(b) -> float`

Converts bytes to decimal TB (÷ 1,000,000,000,000) with 4 decimal places.

Returns `0.0` for `None` or `0`.

**Example:**
```python
tb = bytes_to_tb(1_000_000_000_000)
# Output: 1.0

tb = bytes_to_tb(2_500_000_000_000)
# Output: 2.5

tb = bytes_to_tb(0)
# Output: 0.0
```

---

### `format_duration(start_usecs: int, end_usecs: int) -> str`

Formats elapsed time between two microsecond timestamps as a human-readable duration string.

Format: `"Xh Ym Zs"` (omits zero components)

Returns `""` if either value is missing or zero.

**Example:**
```python
# 1 hour 30 minutes 45 seconds:
duration = format_duration(1000000, 6645000000)
# Output: "1h 30m 45s"

# 5 minutes:
duration = format_duration(1000000, 301000000)
# Output: "5m 0s"

# 30 seconds:
duration = format_duration(1000000, 31000000)
# Output: "30s"

# Missing values:
duration = format_duration(0, 1000000)
# Output: ""
```

---

### `style_worksheet(ws) -> None`

Applies Cohesity branding to an openpyxl worksheet:
- **Header row (row 1):** Bold white text on Cohesity green background (`#00B388`)
- **Column widths:** Auto-fitted based on content

Modifies the worksheet in-place. Use after adding all data but before saving.

**Example:**
```python
from openpyxl import Workbook
from utils.formatters import style_worksheet

wb = Workbook()
ws = wb.active
ws.append(["Cluster", "Status", "Capacity (TB)"])
ws.append(["prod-1", "Healthy", 100])
ws.append(["prod-2", "Healthy", 150])

style_worksheet(ws)  # Apply green header + auto-fit columns

wb.save("report.xlsx")
```

---

### `auto_fit_columns(ws) -> None`

Auto-fits all column widths in an openpyxl worksheet based on content.

Modifies the worksheet in-place.

**Example:**
```python
from openpyxl import load_workbook

wb = load_workbook("report.xlsx")
ws = wb.active
auto_fit_columns(ws)  # Adjust all columns to fit content
wb.save("report.xlsx")
```

---

## Constants

- **`COHESITY_GREEN`** = `"00B388"` — Cohesity brand green color (used for header rows in all reports)

---

## Unit Conversion Reference

| Function | Conversion |
|----------|-----------|
| `bytes_to_gb(b)` | ÷ 1,000,000,000 (decimal SI) |
| `bytes_to_tb(b)` | ÷ 1,000,000,000,000 (decimal SI) |

> **Note:** The Cohesity API returns values in bytes (decimal SI units). These functions convert to decimal GB/TB to match the Helios UI. Binary units (GiB, TiB) are not used.

---

## Timezone Support

The `usecs_to_datetime` function supports timezone-aware conversion:

```python
from zoneinfo import ZoneInfo
from utils.formatters import usecs_to_datetime

# Get cluster timezone from API:
tz_str = cluster_info.get("timeZone")  # e.g., "America/New_York"
tz = ZoneInfo(tz_str)

# Convert timestamps in cluster's local time:
start_time = usecs_to_datetime(run["startTimeUsecs"], tz=tz)
end_time = usecs_to_datetime(run["endTimeUsecs"], tz=tz)
```

---

## Styling Reference

**Header row styling applied by `style_worksheet`:**
- Font: Bold, white (`FFFFFF`)
- Fill: Cohesity green (`00B388`)
- Alignment: Centered
- Frozen panes: Header row remains visible when scrolling down

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-04-06 | feat: initial module — extracted repeated formatting helpers from reporting scripts into a shared module |
