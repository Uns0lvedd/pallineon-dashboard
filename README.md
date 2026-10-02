# Apartment dashboard

A static dashboard for the monthly owner statements, published with GitHub Pages.
Upload the month's xlsx to `reports/`, and the site rebuilds itself in about a minute.

## One-time setup

1. Create a new repository on GitHub. On a free account it must be **public** for GitHub Pages to work,
   which means the repository (including the xlsx files) is visible to anyone who finds it.
   With GitHub Pro the repository can be private while the site stays reachable by link.
2. Upload everything in this folder to the repository, keeping the folder structure.
   Check that `.github/workflows/publish.yml` arrived: hidden folders are sometimes skipped by drag and drop.
   If it is missing, use "Add file", "Create new file", type that path as the name and paste the file's contents.
3. In the repository go to **Settings, Pages** and set **Source** to **GitHub Actions**.
4. Go to the **Actions** tab, open "Build and publish dashboard" and press **Run workflow**.
5. When it finishes, the site address appears under Settings, Pages. Share that link with the partners.

## Every month

1. Open the `reports` folder on GitHub (or use the upload button at the bottom of the dashboard).
2. "Add file", "Upload files", drop in the xlsx, "Commit changes".
3. Wait about a minute and refresh the dashboard.

Rules for the file:

- The file name must contain the month and year, for example `PALLINEON_R_1_SEPTEMBER_2026.xlsx`.
  The year printed inside the sheet is ignored because it has been wrong in the past.
- One file per month. To correct a month, delete the old file and upload the new one.
  Two files for the same month stop the update so nothing is counted twice.
- If a file cannot be read (missing column, month in the wrong row), the update stops,
  GitHub emails you the reason, and the dashboard stays as it was.

You can check a file before publishing with the "Preview a file" box at the bottom of the dashboard.
The preview stays in your browser only.

## How it works

| Path | Purpose |
|---|---|
| `index.html` | The dashboard |
| `reports/` | The manager's xlsx statements, one per month |
| `manual/` | Months entered by hand as JSON (April 2026 came as a PDF). A manual month wins over an xlsx for the same month |
| `config.json` | Property name, partners and their shares |
| `js/parser.js` | Reads a statement. Used by both the build and the in-browser preview |
| `scripts/build.mjs` | Builds `data.json` from `reports/` and `manual/` |
| `fx.json` | Month-end euro to shekel rates, fetched once per month and then frozen |
| `data.json` | Generated. Do not edit |
| `.github/workflows/publish.yml` | Runs the build and publishes the site on every change |

## Changing things

- **Partners or shares:** edit `config.json`. `weight` values are relative, so `1, 1, 1` is an even three-way split
  and `2, 1, 1` would be 50% / 25% / 25%.
- **Entering or correcting a month by hand:** copy `manual/2026-04.json`, rename it to the month, and edit the figures.
- **Exchange rate:** the European Central Bank reference rate for the last business day of each month,
  via frankfurter.dev. To override a month, edit its entry in `fx.json`.

## Running it on your own computer

    node scripts/build.mjs
    python3 -m http.server 8000

Then open http://localhost:8000. Node 18 or later is needed; nothing has to be installed.
