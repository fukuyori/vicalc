# vicalc

A Vim-inspired terminal spreadsheet editor.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Latest Release](https://img.shields.io/github/v/release/fukuyori/vicalc)](https://github.com/fukuyori/vicalc/releases/latest)

[日本語](README_ja.md)

## Overview

vicalc is a structure-oriented spreadsheet editor inspired by VisiCalc and Vim. Instead of mouse-driven selection and modal dialogs, vicalc lets you edit tables by explicitly operating on rows, columns, and cells with keyboard commands.

**Key Philosophy:**
- Keyboard-first interaction
- Row/Column mode as a first-class concept
- Vim-like modal editing
- Terminal-native (no GUI required)

## Features

- **Vim-style navigation** - hjkl, gg, G, Ctrl+f/b/d/u
- **Row/Column mode** - Switch editing direction with /r and /c
- **Formula engine** - 35+ functions (SUM, VLOOKUP, IF, etc.)
- **Absolute/Relative references** - $A$1, $A1, A$1, A1
- **Formula adjustment** - Automatic reference adjustment on row/col insert/delete
- **Copy & Paste** - Internal clipboard (y/p) and system clipboard ("*y/"*p)
- **Visual selection** - Select ranges with v and V
- **Undo/Redo** - Unlimited undo with u
- **File formats** - JSON (native), CSV/TSV import/export
- **Unicode support** - Proper handling of CJK characters

## Installation

### From Binary

Download the latest release from [GitHub Releases](https://github.com/fukuyori/vicalc/releases/latest).

### From Source

```bash
git clone https://github.com/fukuyori/vicalc.git
cd vicalc
cargo build --release
```

The binary will be at `target/release/vicalc`.

## Quick Start

```bash
# Start with empty sheet
vicalc

# Open existing file
vicalc data.json

# Open CSV file
vicalc data.csv
```

## Key Bindings

### Mode Switching

| Key | Action |
|-----|--------|
| `/r` | Switch to Row mode |
| `/c` | Switch to Column mode |
| `v` | Visual selection mode |
| `V` | Visual line/column mode |
| `:` | Command mode |
| `Esc` | Return to Normal mode |

### Navigation

| Key | Action |
|-----|--------|
| `h` `j` `k` `l` | Move left/down/up/right |
| `gg` | Go to top-left (A1) |
| `G` | Go to last cell with data |
| `0` | Go to first column |
| `$` | Go to last column with data |
| `Ctrl+f` | Page down |
| `Ctrl+b` | Page up |
| `Ctrl+d` | Half page down |
| `Ctrl+u` | Half page up |

### Editing

| Key | Action |
|-----|--------|
| `r` | Edit cell (single) |
| `R` | Edit cell (continuous) |
| `F2` | Edit cell (preserve content) |
| `=` | Enter formula |
| `x` | Clear cell |
| `dd` | Delete row/column (based on mode) |
| `o` | Insert row/column after |
| `O` | Insert row/column before |

### Copy & Paste

| Key | Action |
|-----|--------|
| `y` | Copy to internal clipboard |
| `p` | Paste from internal clipboard |
| `"*y` | Copy to system clipboard (TSV) |
| `"*p` | Paste from system clipboard |
| `3p` | Paste 3 times (direction based on mode) |

### Column Width

| Key | Action |
|-----|--------|
| `<` | Decrease column width |
| `>` | Increase column width |
| `:autowidth` | Auto-fit column widths |

### Search

| Key | Action |
|-----|--------|
| `:/pattern` | Search forward |
| `:?pattern` | Search backward |
| `n` | Next match |
| `N` | Previous match |

### Commands

| Command | Action |
|---------|--------|
| `:w [file]` | Save |
| `:e file` | Open file |
| `:q` | Quit |
| `:wq` | Save and quit |
| `:export file.csv` | Export as CSV |
| `:import file.csv` | Import CSV |
| `:goto A1` | Go to cell |
| `:autowidth` | Auto-fit all column widths |
| `:autowidth A:C` | Auto-fit columns A to C |
| `:insrow` | Insert row |
| `:inscol` | Insert column |
| `:delrow` | Delete row |
| `:delcol` | Delete column |

## Supported Functions

### Math & Statistics
`SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `ABS`, `ROUND`, `INT`, `MOD`, `POWER`, `SQRT`

### Conditional
`IF`, `SUMIF`, `COUNTIF`, `AVERAGEIF`, `IFERROR`

### Lookup
`VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`

### Text
`LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, `UPPER`, `LOWER`, `CONCATENATE`

### Logical
`AND`, `OR`, `NOT`

### Information
`ISBLANK`, `ISNUMBER`, `ISTEXT`

## File Formats

### Native Format (JSON)

vicalc uses JSON as its native format, storing:
- Cell values and formulas
- Column widths
- Sheet name

```json
{
  "version": "1.0",
  "name": "Sheet1",
  "cells": {
    "A1": "Hello",
    "B1": "=SUM(A2:A10)"
  },
  "col_widths": {
    "A": 15
  }
}
```

### CSV/TSV

- Import: `:import file.csv`
- Export: `:export file.csv`
- System clipboard uses TSV format

## Row/Column Mode

vicalc has a unique concept of "editing axis":

- **Row mode** (`/r`): Operations extend horizontally
  - Continuous paste expands right
  - `dd` deletes a row
  - `o` inserts a row below

- **Column mode** (`/c`): Operations extend vertically
  - Continuous paste expands down
  - `dd` deletes a column
  - `o` inserts a column right

The current mode is shown in the status bar.

## License

MIT License. See [LICENSE](LICENSE) for details.

## Author

[@fukuyori](https://github.com/fukuyori)
