# Excel → PowerPoint Org Chart Tool

Turns a workbook (one department per sheet) into a fully laid-out PowerPoint
organization chart, using each sheet's `ReportsToName` column to build the
reporting tree — no manual editing of the Excel file required.

## Install

```bash
pip install openpyxl python-pptx
```

## Run

```bash
python org_chart_tool.py YOUR_WORKBOOK.xlsx
```

This creates `YOUR_WORKBOOK_Org_Chart.pptx` in the same folder.

Optional flags:

```bash
python org_chart_tool.py YOUR_WORKBOOK.xlsx -o MyOrgChart.pptx --company "Acme Corp"
python org_chart_tool.py YOUR_WORKBOOK.xlsx --template Sample_Template.pptx --company "Acme Corp"
```

| Flag | What it does | Default |
|---|---|---|
| `-o / --output` | Path for the generated `.pptx` | `<input name>_Org_Chart.pptx` |
| `--company` | Name shown on the title slide | derived from the input filename |
| `--template` | A `.pptx` whose first slide (title) and closing "Thank You" slide should bookend the generated chart content — see below | none (uses a plain built-in title slide, no closing slide) |

### Using `--template`

`Sample_Template.pptx` is a two-slide branded deck: a title slide and a
closing "Thank You" slide, both with real logos/imagery. Passing it via
`--template`:

1. Opens that deck as the base file instead of a blank presentation.
2. Replaces the `xxxx` placeholder text on its first slide with `--company`.
3. Generates the disclaimer + department chart slides in between, using
   that deck's own "Blank" layout — which also means any decorative
   banner/footer/page-number graphics baked into that layout's slide
   master appear automatically on every chart slide, with no extra code.
4. Moves the slide whose text is exactly "Thank You" to the very end.

To use a different title/closing template, swap in any `.pptx` with the
same shape: a first slide containing literal placeholder text `xxxx`, and
a slide somewhere in the file with the exact text `Thank You`.

## Expected workbook format

Each sheet = one department. Required columns (header row 1):

| Column | Meaning |
|---|---|
| `Name` | Employee's full name |
| `ReportsToName` | The **Name** of their manager, as it appears elsewhere **in the same sheet**. Blank, or a name not found in the sheet, means this person is shown as a top-level box for that department. |

Optional columns, shown on each box if present:

| Column | Meaning |
|---|---|
| `Designation` | Job title |
| `Department` | Sub-department / region label, shown after the title |
| `LinkedinUrl` | If present and not blank/`NA`, the person's name is hyperlinked to this URL |

Any other columns are ignored, and the workbook itself is never modified.

## What the tool does automatically

- Builds one reporting tree per sheet from `ReportsToName`.
- Lays out each tree top-down, centering managers over their direct reports.
- Colors each box's border by hierarchy level (4-color palette, repeating for
  levels beyond 4).
- When a manager has too many direct reports to fit legibly on one slide,
  automatically splits them across multiple slides, repeating that manager's
  box at the top of each continuation slide (labelled `Department (2/4)`,
  etc.).
- Adds a title slide and a disclaimer slide (hierarchy-level legend included).

## Customizing appearance

All the visual constants live near the top of `org_chart_tool.py`:

- `LEVEL_COLORS` — the 4-color hex palette used for box borders by level.
- `NODE_W`, `NODE_H` — box width/height in inches.
- `SLIDE_W`, `SLIDE_H` — slide dimensions (default 13.33×7.5in, widescreen).
- `FONT` — typeface used throughout.

Change these and re-run the script to regenerate the deck.
