# DW-SFTIES Product Roadmap Dashboard

Turns `data/roadmap.xlsx` into a filterable web dashboard. Upload a new workbook,
GitHub Actions rebuilds the page and publishes it — no manual steps.

```
data/roadmap.xlsx  ──push──►  GitHub Actions  ──►  docs/index.html  ──►  GitHub Pages
```

---

## Your routine every 2 months

1. Update the roadmap table in `roadmap.xlsx` as usual (tick marks, circles, diamonds, release numbers).
2. On the **Dashboard Info** sheet, edit the yellow cells:
   - **Current release** — the release just shipped, e.g. `0.25`
   - **Current release UAT date** — e.g. `08 August 2026`
   - **Release in development** — e.g. `0.26`
   - **Next release UAT date** — e.g. `October 2026`
3. Save, then upload it to GitHub: open the `data` folder → **Add file** → **Upload files** → drop the workbook in → **Commit changes**.
4. Wait about a minute. The **Actions** tab shows the run; when it's green the published page is live.

That's it. Steps 2 and 3 are the whole update — the release wording travels inside
the workbook, so there's no second file to remember.

---

## One-time setup

1. Create a repository and upload everything in this folder.
2. **Settings → Pages → Build and deployment → Source: GitHub Actions.**
3. **Actions** tab → run **Build and publish roadmap dashboard** once. The run
   summary prints your dashboard URL (`https://<you>.github.io/<repo>/`).

If the repo is private, Pages needs GitHub Enterprise. On a free plan either make
the repo public, or skip Pages and download the HTML from the run's
**roadmap-dashboard-html** artifact instead.

### Running it on your own machine

```bash
pip install -r requirements.txt
python scripts/build_dashboard.py
open docs/index.html
```

---

## The colour rules

Set in `config.yml` under `statuses`, applied to the nine layer columns
(User Story, DB, API, UI, DMS, ODS, SFDW, Production Control, Reporting Services):

| Cell contains | Background | Meaning |
|---|---|---|
| `√` tick | green `#d7ffd7` | complete |
| `●` black circle | yellow `#ffffc3` | planned |
| `♦` diamond | orange `#ffe0b2` | under development |
| `n/a` | white, grey text | not applicable |
| `-` or empty | grey `#e9e9e9` | placeholder |

And in the **Latest Release** column:

| Cell value | Background |
|---|---|
| matches **Current release** | orange |
| matches **Future release label** (`Future Release`) | yellow |
| any other release number | white |

Each status accepts alternative characters via its `match` list, so a `✓` or `✔`
typed instead of `√` still comes out green. Add to that list if your team uses
another symbol.

---

## Where to change things

| Want to change | Where |
|---|---|
| Release numbers, UAT dates | **Dashboard Info** sheet in the workbook (preferred), or `config.yml` → `release:` |
| Any colour | `config.yml` → `statuses:` and `theme:` |
| Title, intro paragraph, release-notes link | `config.yml` → `dashboard:` |
| Which columns merge vertically | `config.yml` → `table.group_columns` |
| Which columns hold symbols | `config.yml` → `table.status_columns` |
| Sheet name or header row | `config.yml` → `source:` |
| Layout, fonts, spacing | `scripts/template.html` |

Precedence, lowest to highest: **`config.yml` → Dashboard Info sheet → workflow inputs.**
The last one lets you correct a typo in a date from the Actions tab
(**Run workflow** → fill a field) without editing any file.

---

## What the dashboard does

- **Column filters** — dropdown per column; free-text box on high-cardinality columns
- **Search** — one box across every column
- **Status chips** — click *Complete* / *Under development* / *Planned* to filter to those rows
- **Expand roadmap** — hides the text panels to give the table the full window
- **Simplified view** — hides Past Releases plus any layer column that is entirely `n/a` for the rows on screen
- **Export data (CSV)** — downloads exactly the rows and columns currently visible, Excel-friendly
- **Sort** — click any header; *Reset all filters and sorting* returns to workbook order
- Merged Component/Module/Submodule/Element cells, recalculated whenever you filter, with tall group labels that follow the scroll
- Sticky header, keyboard focus rings, print stylesheet, works down to mobile width

---

## Files

```
config.yml                        settings, colours, release info
data/roadmap.xlsx                 your workbook (Sheet1 + Dashboard Info)
scripts/build_dashboard.py        reads the workbook, writes the page
scripts/template.html             layout, CSS and browser code
scripts/add_config_sheet.py       one-off: adds the Dashboard Info sheet
docs/index.html                   generated dashboard (one self-contained file)
.github/workflows/build-dashboard.yml   the automation
```

`docs/index.html` is generated — edit `scripts/template.html` instead, or your
changes are overwritten on the next upload.

---

## Troubleshooting

**Action ran but the page looks unchanged.** Hard-refresh (Ctrl/Cmd+Shift+R).
The footer shows the build timestamp and row count.

**"no data rows found".** `source.sheet_name` or `source.header_row` in
`config.yml` no longer matches the workbook. The error lists the sheets it found.

**A symbol shows on a white background.** That character isn't in any status
`match` list. The build log prints per-status counts — compare them against the
workbook to spot it.

**"status columns not found in the sheet, skipped: …"** in the log means a column
in `table.status_columns` was renamed in the workbook. Update the config to match.

**Pages deploy fails with a permissions error.** Settings → Pages → Source must be
**GitHub Actions**, not *Deploy from a branch*.
