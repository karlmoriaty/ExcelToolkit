# ExcelToolkit

A spreadsheet-like desktop editor with a command console. Edit cells directly in a
Tkinter grid, then apply layout and formatting work with one-line commands instead
of clicking through menus — useful when you format the same kind of report over and
over.

## Features

- **Spreadsheet grid** — 100 × 50 editable cells with an Excel-style layout.
- **Worksheet management** — create, switch between, and delete sheets.
- **Format templates** — reusable styles for headers, data cells, numbers, titles and
  highlights.
- **Command console** — apply templates, borders, fills, alignment and column
  autofit by typing a command.
- **Excel I/O** — open and save `.xlsx` files via openpyxl.

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
python main.py
```

or double-click `run_main.bat`. `card_view.py` and `integrated_app.py` are
alternative entry points with their own `.bat` launchers.

## Layout

- **Left panel** — worksheet list and format templates.
- **Center** — the editable grid.
- **Right panel** — command console.

## Commands

| Command | Description | Example |
|---------|-------------|---------|
| `help` | Show command help | `help` |
| `template` | List available templates | `template` |
| `table <range> <template>` | Apply a table template | `table A1:E10 header` |
| `border <range>` | Add borders | `border A1:E10` |
| `fill <range> <color>` | Fill background color (hex) | `fill A1:A5 FFFF00` |
| `align <range> <left/center/right>` | Set cell alignment | `align A1:E10 center` |
| `autofit` | Autofit column widths | `autofit` |
| `select <sheet>` | Switch worksheet | `select Sheet1` |
| `create <sheet>` | Create a worksheet | `create Data` |
| `set <cell> <value>` | Set a cell value | `set A1 Name` |

### Built-in templates

- `header` — bold, grey background, centered, bordered
- `data` — left aligned, bordered
- `number` — right aligned, bordered
- `title` — large bold centered text
- `highlight` — yellow background, bold

### Shortcuts

`Ctrl+N` new file, `Ctrl+O` open, `Ctrl+S` save.

## Example session

1. Type headers into A1:E1 (Name, Age, Gender, Role, City).
2. Run `table A1:E1 header`.
3. Fill A2:E10 with data and run `table A2:E10 data`.
4. Run `autofit`, then save.

## Stack

Python 3, Tkinter, openpyxl.

## Notes

- The UI is in Chinese.
- The grid is a fixed 100 × 50 area.

## License

MIT
