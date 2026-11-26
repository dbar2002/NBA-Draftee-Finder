# NBA Draftee Finder

Every active NBA player listed under the team that drafted him, plus undrafted
players under the team they played their first NBA season for. Each team gets a
Top 20 (drafted players first, undrafted starters fill the rest), ranked by career
Win Shares. Players inactive two or more seasons and not on a current roster are
highlighted red.

In the web app, filter each team by **Status** (still with team, played there then
left, or traded before playing there, such as a draft-night trade) and export the
current view to Excel for one team or all teams.

Data comes from [Basketball-Reference](https://www.basketball-reference.com).

- **Web app:** `docs/index.html` (served by GitHub Pages once enabled)
- **Excel:** `docs/NBA_Active_Draftees_by_Team.xlsx`

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\Activate.ps1
pip install -r requirements.txt
python nba_draftees.py
```

About 7 minutes; it waits ~3.5 s between pages to respect the site's rate limit.
Writes `NBA_Active_Draftees_by_Team.xlsx` and `NBA_Draftee_Finder.html` here
(both git-ignored). To update the published copy:

```bash
python nba_draftees.py --out docs/NBA_Active_Draftees_by_Team.xlsx --app docs/index.html
git add docs && git commit -m "Data refresh" && git push
```

Options: `--cache-hours 0` (force fresh downloads), `--delay` (seconds between
requests), `--red-seasons` (default 2), `--app ''` (skip the web app).

## Automatic daily refresh

`.github/workflows/refresh.yml` runs the scraper every morning and commits the
new files in `docs/`. Run it by hand from the **Actions** tab ("Run workflow").

Basketball-Reference sometimes blocks cloud servers with a bot check. If the
workflow fails with "Could not fetch ... (bot check or block)", run it locally
instead (above) or on a schedule from your own machine:

```
0 8 * * * cd /path/to/nba-draftee-finder && .venv/bin/python nba_draftees.py --out docs/NBA_Active_Draftees_by_Team.xlsx --app docs/index.html && git add docs && git commit -qm "Data refresh" && git push -q
```

## Files

| File | What it is |
| --- | --- |
| `nba_draftees.py` | Scraper and workbook/app builder |
| `nba_app_template.html` | Web app template the script fills in |
| `docs/` | Latest published web app and Excel file |
| `.github/workflows/refresh.yml` | Daily GitHub Actions refresh |
