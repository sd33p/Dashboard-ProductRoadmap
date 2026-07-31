# DW-SFTIES Release Tracker — Static Dashboard

A self-rebuilding static dashboard that replaces the Qlik app. Python reads
`data/dashboard_data.xlsx` and generates a single static HTML page
(`docs/index.html`) that GitHub Pages serves directly — no server, no live
Python process, just a static file that gets regenerated automatically
whenever you update the spreadsheet.

## How the update cycle works

1. Every ~2 months, replace `data/dashboard_data.xlsx` with the new export
   (keep the exact filename).
2. Commit and push to `main`.
3. GitHub Actions (`.github/workflows/build.yml`) automatically:
   - installs Python + pandas/openpyxl
   - runs `build_dashboard.py`, which reads the Excel file and regenerates
     `docs/index.html`
   - commits that regenerated file back to `main`
4. GitHub Pages (serving from the `main` branch, `/docs` folder) picks up the
   new file within a minute or two. No manual rebuild step needed.

## One-time setup

1. Create a new GitHub repo and push this folder to it.
2. In the repo, go to **Settings → Pages**.
3. Under "Build and deployment", set:
   - Source: `Deploy from a branch`
   - Branch: `main`, folder: `/docs`
4. Save. GitHub will give you a URL like
   `https://<your-username>.github.io/<repo-name>/`.
5. Push once (or run the workflow manually from the **Actions** tab) to
   trigger the first build.

No secrets or extra configuration are required — `permissions: contents:
write` in the workflow is enough for the Action to commit the rebuilt file
back to the repo.

## Updating the data

```bash
cp /path/to/new/export.xlsx data/dashboard_data.xlsx
git add data/dashboard_data.xlsx
git commit -m "Update Q3 release data"
git push
```

That's it — the Action rebuilds the dashboard and republishes it.

## Running locally / previewing before you push

```bash
pip install -r requirements.txt
python build_dashboard.py
open docs/index.html      # or just double-click the file
```

## Expected spreadsheet format

Single sheet, header row containing (order doesn't matter, matching is
case-insensitive):

`Component | Module | Sub Module | Element | Sub element | User Story | DB | API | UI | Latest Release | Past Release`

Status columns (`DB`, `API`, `UI`) use these symbols:

| Symbol | Meaning     |
|--------|-------------|
| ✓      | Complete    |
| ♦      | In Progress |
| ●      | Not Started |
| -      | N/A         |

If the sheet is missing any of the expected columns, `build_dashboard.py`
will fail loudly with a clear error message rather than publishing a broken
dashboard — check the Action's log under the **Actions** tab if a build
fails.

## Customizing

- **Layout / styling / filters**: edit `template.html` — it's a single
  self-contained file (HTML + CSS + JS), no build tooling required.
- **Data shape / column mapping**: edit `COLUMN_MAP` and `STATUS_SYMBOLS` in
  `build_dashboard.py`.
- **New columns**: add them to `COLUMN_MAP`, include them in `build_records`
  in `build_dashboard.py`, and add a matching `<th>`/`<td>` in
  `template.html`.
