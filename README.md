# NBA Draftee Finder

Active NBA players by the team that drafted them, plus undrafted players by the
team they played their first NBA season for. Ranked by career Win Shares; players
inactive 2+ seasons and not on a current roster are highlighted red.

## Files
- `NBA_Draftee_Finder.html` - the web app. Double-click to open in any browser.
  Data as of November 24, 2026.
- `nba_draftees.py` - refreshes the data from Basketball-Reference.
- `nba_app_template.html` - template the script fills in. Keep it next to the script.
- `requirements.txt` - Python packages the script needs.

## Refresh the data
```
pip install -r requirements.txt
python3 nba_draftees.py
```
Takes about 5 minutes (it waits ~3.5s between pages to respect the site's rate limit).
Writes `NBA_Draftee_Finder.html` and `NBA_Active_Draftees_by_Team.xlsx` in this folder.

Options: `--out` (Excel path), `--app` (web app path, '' to skip), `--delay` (seconds
between requests), `--red-seasons` (default 2).

## Run it on a schedule
Mac/Linux, Mondays at 8 AM (`crontab -e`):
```
0 8 * * 1 cd /path/to/nba-draftee-finder && /usr/bin/python3 nba_draftees.py >> run.log 2>&1
```
Windows: Task Scheduler, action `python nba_draftees.py`, "Start in" set to this folder.

## Host it (optional)
The app is one self-contained HTML file. Rename it `index.html` and upload it to
GitHub Pages, Netlify Drop or any static host.

## Troubleshooting
"Could not fetch ... (bot check or block)": wait a while or raise `--delay`.
