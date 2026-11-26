#!/usr/bin/env python3
"""
Active NBA players by drafting team -> Excel workbook.

Pulls from Basketball-Reference:
  * /players/a/ .. /players/z/          active players (bold names) + last season
  * /draft/NBA_<year>.html              drafting team, pick, college
  * /leagues/NBA_<year>_advanced.html   Win Shares by season (summed to career)

Output: one .xlsx with a Summary tab and one tab per franchise.
A Top 20 tab ranks each team's best 20: drafted players first, then undrafted
starters to fill. Each team tab lists its active draftees, then the active undrafted players who
played their first NBA season with that team, both ranked by career Win Shares. Players who haven't played in the
last two seasons and aren't on a current roster are highlighted red and placed at the bottom.

Usage:
  pip install requests beautifulsoup4 openpyxl
  python3 nba_draftees.py                      # writes NBA_Active_Draftees_by_Team.xlsx
  python3 nba_draftees.py --out ~/Desktop/nba.xlsx --delay 4
"""

import argparse
import datetime as dt
import hashlib
import os
import random
import string
import sys
import time
from collections import defaultdict
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

BASE = "https://www.basketball-reference.com"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# Historical team codes -> current franchise
FRANCHISE = {"SEA": "OKC", "NJN": "BRK", "NOH": "NOP", "NOK": "NOP", "CHA": "CHO", "CHH": "NOP", "VAN": "MEM"}

TEAM_NAMES = {
    "ATL": "Atlanta Hawks", "BOS": "Boston Celtics", "BRK": "Brooklyn Nets", "CHI": "Chicago Bulls",
    "CHO": "Charlotte Hornets", "CLE": "Cleveland Cavaliers", "DAL": "Dallas Mavericks",
    "DEN": "Denver Nuggets", "DET": "Detroit Pistons", "GSW": "Golden State Warriors",
    "HOU": "Houston Rockets", "IND": "Indiana Pacers", "LAC": "LA Clippers",
    "LAL": "Los Angeles Lakers", "MEM": "Memphis Grizzlies", "MIA": "Miami Heat",
    "MIL": "Milwaukee Bucks", "MIN": "Minnesota Timberwolves", "NOP": "New Orleans Pelicans",
    "NYK": "New York Knicks", "OKC": "Oklahoma City Thunder", "ORL": "Orlando Magic",
    "PHI": "Philadelphia 76ers", "PHO": "Phoenix Suns", "POR": "Portland Trail Blazers",
    "SAC": "Sacramento Kings", "SAS": "San Antonio Spurs", "TOR": "Toronto Raptors",
    "UTA": "Utah Jazz", "WAS": "Washington Wizards",
}

ROSTERS = {}  # filled in main(): pid -> current-season team

LINK = Font(color="0563C1", underline="single")


def link_cell(c, pid):
    """Turn a player-name cell into a link to his Basketball-Reference page."""
    if pid:
        c.hyperlink = profile_url(pid)
        c.font = LINK


RED = PatternFill("solid", fgColor="F4B6B6")
HEADER = PatternFill("solid", fgColor="F3F3F3")


# ---------------------------------------------------------------- fetching

class Fetcher:
    """Polite fetcher: delay between requests, retries, on-disk cache."""

    def __init__(self, delay, cache_dir, cache_hours):
        self.delay = delay
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.cache_secs = cache_hours * 3600
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self._last = 0.0

    def get(self, path, must_contain):
        url = BASE + path
        key = self.cache_dir / (hashlib.md5(url.encode()).hexdigest() + ".v2.html")
        if key.exists() and time.time() - key.stat().st_mtime < self.cache_secs:
            return key.read_text(encoding="utf-8")

        for attempt in range(5):
            wait = self.delay - (time.time() - self._last)
            if wait > 0:
                time.sleep(wait + random.uniform(0, 0.5))
            self._last = time.time()
            try:
                r = self.session.get(url, timeout=30)
            except requests.RequestException as e:
                print(f"  network error on {path}: {e}", file=sys.stderr)
                time.sleep(10 * (attempt + 1))
                continue
            # BBR doesn't always send a charset, and requests then assumes Latin-1,
            # which garbles names like Şengün and Jokić. The pages are UTF-8.
            r.encoding = "utf-8"
            if r.status_code == 404:
                return None
            if r.status_code == 429:
                print("  rate-limited (429); waiting 60s", file=sys.stderr)
                time.sleep(60)
                continue
            if r.ok and must_contain in r.text:
                key.write_text(r.text, encoding="utf-8")
                return r.text
            # bot check page or missing table
            print(f"  {path}: HTTP {r.status_code}, table not found; retrying", file=sys.stderr)
            time.sleep(15 * (attempt + 1))
        raise RuntimeError(f"Could not fetch {url} (bot check or block). Try again later or raise --delay.")


def player_id(tag):
    """Basketball-Reference player id from a cell (data-append-csv) or its link."""
    if tag is None:
        return None
    pid = tag.get("data-append-csv")
    if pid:
        return pid
    a = tag.find("a")
    if a and a.get("href", "").startswith("/players/"):
        return a["href"].rsplit("/", 1)[-1].replace(".html", "")
    return None


def cell(row, stat):
    c = row.find(attrs={"data-stat": stat})
    return c.get_text(strip=True) if c else ""


# ---------------------------------------------------------------- scraping

def get_active_players(f):
    """{pid: {name, last_end_year}} for every bold (active) player in the index."""
    active = {}
    for letter in string.ascii_lowercase:
        html = f.get(f"/players/{letter}/", 'id="players"')
        if html is None:
            continue
        soup = BeautifulSoup(html, "html.parser")
        for tr in soup.select("table#players tbody tr"):
            th = tr.find("th")
            if th is None or th.find("strong") is None:
                continue
            a = th.find("a")
            pid = a["href"].rsplit("/", 1)[-1].replace(".html", "")
            to, frm = cell(tr, "year_max"), cell(tr, "year_min")
            active[pid] = {"name": a.get_text(strip=True), "last": int(to) if to.isdigit() else 0,
                           "first": int(frm) if frm.isdigit() else 0,
                           "pos": cell(tr, "pos"), "height": cell(tr, "height"),
                           "college": cell(tr, "colleges") or None}
        print(f"  players/{letter}: {len(active)} active so far")
    return active


def get_drafts(f, first_year, last_year):
    """{pid: {year, pick, team, college}}"""
    drafts = {}
    for year in range(last_year, first_year - 1, -1):
        html = f.get(f"/draft/NBA_{year}.html", 'id="stats"')
        if html is None:
            continue
        soup = BeautifulSoup(html, "html.parser")
        for tr in soup.select("table#stats tbody tr"):
            pid = player_id(tr.find(attrs={"data-stat": "player"}))
            pick = cell(tr, "pick_overall")
            if not pid or not pick.isdigit():
                continue
            drafts[pid] = {
                "year": year,
                "pick": int(pick),
                "team": cell(tr, "team_id"),
                "college": cell(tr, "college_name") or None,
            }
        print(f"  draft {year}: {len(drafts)} picks so far")
    return drafts


def get_win_shares(f, first_season_end, current_end):
    """Returns (career_ws, last_season_seen, current_team, first_team).

    BBR lists a traded player's combined row (TOT / 2TM / 3TM) first, then one row
    per team in the order he played for them. So per season: the first row is the
    season total, the first single-team row is where he started, and the last
    single-team row is where he is now.
    """
    career = defaultdict(float)
    last_seen = {}
    current_team = {}
    first_team = {}
    for end in range(current_end, first_season_end - 1, -1):
        html = f.get(f"/leagues/NBA_{end}_advanced.html", 'id="advanced"')
        if html is None:
            continue
        soup = BeautifulSoup(html, "html.parser")
        seen = set()
        started = {}
        for tr in soup.select("table#advanced tbody tr"):
            if "thead" in (tr.get("class") or []):
                continue
            pid = player_id(tr.find(attrs={"data-stat": "name_display"}) or tr.find(attrs={"data-stat": "player"}))
            if not pid:
                continue
            team = cell(tr, "team_name_abbr") or cell(tr, "team_id")
            single = not team.endswith("TM") and team != "TOT"
            if single:
                started.setdefault(pid, team)
                if end == current_end:
                    current_team[pid] = team  # last single-team row wins
            if pid in seen:
                continue
            seen.add(pid)
            try:
                career[pid] += float(cell(tr, "ws") or 0)
            except ValueError:
                pass
            last_seen.setdefault(pid, end)
        first_team.update(started)  # going back in time, so the earliest season wins
        print(f"  advanced {end}: {len(seen)} players")
    return career, last_seen, current_team, first_team


def get_rosters(f, current_end):
    """{pid: team} for every player on a current-season roster, including players
    who signed but haven't played yet (injured, just returned, two-way)."""
    rosters = {}
    for t in sorted(TEAM_NAMES):
        html = f.get(f"/teams/{t}/{current_end}.html", 'id="roster"')
        if html is None:
            continue
        soup = BeautifulSoup(html, "html.parser")
        for tr in soup.select("table#roster tbody tr"):
            pid = player_id(tr.find(attrs={"data-stat": "player"}))
            if pid:
                rosters[pid] = t
        print(f"  roster {t}: {len(rosters)} players so far")
    return rosters


# ---------------------------------------------------------------- workbook

def season_label(end):
    return f"{end - 1}-{str(end)[2:]}"


def profile_url(pid):
    return f"{BASE}/players/{pid[0]}/{pid}.html"


def player_row(p, pid, franchise, career, last_seen, current_team, current_end, red_cutoff):
    last = max(p["last"], last_seen.get(pid, 0))  # index can lag a few days
    played_now = last >= current_end
    rostered = ROSTERS.get(pid)
    cur = (current_team.get(pid, "") if played_now else "") or rostered or ""
    return {
        "name": p["name"],
        "pid": pid,
        "pos": p.get("pos", ""),
        "height": p.get("height", ""),
        "college": p.get("college"),
        "first": p["first"],
        "last": last,
        "played": "Yes" if played_now else "No",
        "cur": cur,
        "still": "Yes" if cur == franchise else "No",
        "ws": round(career.get(pid, 0.0), 1),
        "red": last <= red_cutoff and not rostered,  # on a roster now = not inactive
    }


def rank(rows):
    rows.sort(key=lambda r: -r["ws"])
    return [r for r in rows if not r["red"]] + [r for r in rows if r["red"]]


def build_undrafted(active, drafts, career, last_seen, current_team, first_team, current_end, red_cutoff):
    """Active players never drafted, grouped by the team they played their first NBA season for."""
    by_team = defaultdict(list)
    for pid, p in active.items():
        if pid in drafts:
            continue
        t = first_team.get(pid)
        if not t:
            continue  # signed but hasn't played a game yet
        franchise = FRANCHISE.get(t, t)
        if franchise in TEAM_NAMES:
            by_team[franchise].append(player_row(p, pid, franchise, career, last_seen, current_team, current_end, red_cutoff))
    return {t: rank(v) for t, v in by_team.items()}


STATUS = {"w": "Still with team", "l": "Played here, then left", "t": "Traded before playing here"}


def build_rows(active, drafts, career, last_seen, current_team, current_end, red_cutoff, first_team=None):
    by_team = defaultdict(list)
    for pid, p in active.items():
        d = drafts.get(pid)
        if not d:
            continue  # undrafted
        franchise = FRANCHISE.get(d["team"], d["team"])
        if franchise not in TEAM_NAMES:
            continue
        r = player_row(p, pid, franchise, career, last_seen, current_team, current_end, red_cutoff)
        r.update(year=d["year"], pick=d["pick"], round="1st" if d["pick"] <= 30 else "2nd",
                 college=d["college"] or r["college"])
        ft = (first_team or {}).get(pid)
        ft = FRANCHISE.get(ft, ft) if ft else None
        r["debut_team"] = ft
        # t = never played for the drafting team (draft-night trade, or traded before his debut)
        r["status"] = "t" if ft and ft != franchise else ("w" if r["still"] == "Yes" else "l")
        by_team[franchise].append(r)
    return {t: rank(v) for t, v in by_team.items()}


def top20(drafted, undrafted, n=20):
    """Drafted first, then undrafted starters, each by career WS; inactive (red) last."""
    key = lambda r: -r["ws"]
    d = [("Drafted", r) for r in drafted]
    u = [("Undrafted", r) for r in undrafted]
    pick = lambda rows, red: sorted([x for x in rows if x[1]["red"] == red], key=lambda x: key(x[1]))
    return (pick(d, False) + pick(u, False) + pick(d, True) + pick(u, True))[:n]


def write_workbook(by_team, undrafted, current_end, out):
    wb = Workbook()
    s = wb.active
    s.title = "Summary"
    cur = season_label(current_end)
    s.append(["Franchise", "Team", "Active Draftees", f"Played {cur}", "On Drafting Team Now", "Undrafted Starters"])
    for c in s[1]:
        c.fill = HEADER
    teams = sorted(TEAM_NAMES)
    for i, t in enumerate(teams, start=2):
        n = len(by_team.get(t, [])) + 1
        if n > 1:
            counts = [f"=COUNTA('{t}'!A2:A{n})", f"=COUNTIF('{t}'!G2:G{n},\"Yes\")",
                      f"=COUNTIF('{t}'!I2:I{n},\"Yes\")"]
        else:
            counts = [0, 0, 0]
        s.append([t, TEAM_NAMES[t]] + counts + [len(undrafted.get(t, []))])
    total = len(teams) + 2
    s.append(["Total", None] + [f"=SUM({c}2:{c}{total-1})" for c in "CDEF"])
    for c in s[total]:
        c.font = Font(bold=True)
    s.auto_filter.ref = f"A1:F{total-1}"
    s.column_dimensions["B"].width = 24
    for col in "CDEF":
        s.column_dimensions[col].width = 18

    # Top 20 per team
    tw = wb.create_sheet("Top 20")
    tw.append(["Team", "Rank", "Player", "Type", "Draft", "Last Season",
               "Current Team", "Still With Team", "Career WS"])
    for c in tw[1]:
        c.fill = HEADER
    for t in teams:
        for i, (kind, r) in enumerate(top20(by_team.get(t, []), undrafted.get(t, [])), start=1):
            how = f"{r['year']} #{r['pick']}" if kind == "Drafted" else \
                  (f"Undrafted, debut {season_label(r['first'])}" if r["first"] else "Undrafted")
            tw.append([t, i, r["name"], kind, how, season_label(r["last"]),
                       r["cur"] or None, r["still"], r["ws"]])
            tw.cell(tw.max_row, 9).number_format = "0.0"
            link_cell(tw.cell(tw.max_row, 3), r.get("pid"))
            if r["red"]:
                for c in tw[tw.max_row]:
                    c.fill = RED
    tw.freeze_panes = "A2"
    tw.auto_filter.ref = f"A1:I{tw.max_row}"
    for col, w in zip("ABCDEFGHI", [7, 6, 24, 10, 22, 11, 12, 13, 10]):
        tw.column_dimensions[col].width = w

    headers = ["Player", "Draft Year", "Pick", "Round", "College", "Last Season",
               f"Played {cur}", "Current Team", "Still With Team", "Career WS", "Status"]
    widths = [24, 10, 6, 7, 22, 11, 13, 12, 13, 11, 30]
    for t in teams:
        ws = wb.create_sheet(t)
        ws.append(headers)
        for c in ws[1]:
            c.fill = HEADER
        for r in by_team.get(t, []):
            ws.append([r["name"], r["year"], r["pick"], r["round"], r["college"],
                       season_label(r["last"]), r["played"], r["cur"] or None, r["still"], r["ws"],
                       STATUS[r["status"]] + (f" (debuted with {r['debut_team']})" if r["status"] == "t" else "")])
            row = ws.max_row
            ws.cell(row, 10).number_format = "0.0"
            link_cell(ws.cell(row, 1), r.get("pid"))
            if r["red"]:
                for c in ws[row]:
                    c.fill = RED
        # second table: undrafted players who started with this team
        ws.append([])
        ws.append(["Undrafted, started here"])
        ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
        ws.append(["Player", "First Season", "Pos", "Height", "College", "Last Season",
                   f"Played {cur}", "Current Team", "Still With Team", "Career WS"])
        for c in ws[ws.max_row]:
            c.fill = HEADER
        for r in undrafted.get(t, []):
            ws.append([r["name"], season_label(r["first"]) if r["first"] else None, r["pos"] or None,
                       r["height"] or None, r["college"],
                       season_label(r["last"]), r["played"], r["cur"] or None, r["still"], r["ws"]])
            row = ws.max_row
            ws.cell(row, 10).number_format = "0.0"
            link_cell(ws.cell(row, 1), r.get("pid"))
            if r["red"]:
                for c in ws[row]:
                    c.fill = RED
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = f"A1:K{max(len(by_team.get(t, [])) + 1, 2)}"
        for col, w in zip("ABCDEFGHIJK", widths):
            ws.column_dimensions[col].width = w
    wb.save(out)


def write_app(by_team, undrafted, template, out, as_of, current_end, red_cutoff):
    """Fill the web-app template with this run's data."""
    import json
    lab = season_label
    D = {t: [[r["name"], r["year"], r["pick"], r["college"], lab(r["last"]), r["played"],
              r["cur"] or None, r["still"], r["ws"], r.get("status", "l"),
              r.get("debut_team") if r.get("status") == "t" else None, r["pos"], r["height"], None, r["pid"]]
             for r in rows] for t, rows in by_team.items()}
    U = {t: [[r["name"], lab(r["first"]) if r["first"] else None, None, r["college"], lab(r["last"]), r["played"],
              r["cur"] or None, r["still"], r["ws"], "w" if r["still"] == "Yes" else "l", None, r["pos"], r["height"],
              None, r["pid"]]
             for r in rows] for t, rows in undrafted.items()}
    html = Path(template).read_text(encoding="utf-8")
    html = html.replace("__DATA__", json.dumps(D, ensure_ascii=False, separators=(",", ":")))
    html = html.replace("__UND__", json.dumps(U, ensure_ascii=False, separators=(",", ":")))
    html = html.replace("as of November 24, 2026", "as of " + as_of)
    # template was written for the 2026-27 season; shift its labels to this run's season
    html = html.replace("+ 1 <= 2025", f"+ 1 <= {red_cutoff}")
    html = html.replace("2024-25 or earlier", f"{season_label(red_cutoff)} or earlier")
    html = html.replace("2026-27", season_label(current_end))
    Path(out).write_text(html, encoding="utf-8")


# ---------------------------------------------------------------- main

def current_season_end(today):
    # NBA season starts in October: Oct 2026 - Jun 2027 is the "2027" season.
    return today.year + 1 if today.month >= 10 else today.year


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default="NBA_Active_Draftees_by_Team.xlsx")
    ap.add_argument("--delay", type=float, default=3.5, help="seconds between requests (keep >= 3)")
    ap.add_argument("--cache", default=os.path.expanduser("~/.cache/nba_draftees"))
    ap.add_argument("--cache-hours", type=float, default=12, help="reuse downloaded pages this long")
    ap.add_argument("--first-draft", type=int, default=2001, help="oldest draft year to check")
    ap.add_argument("--app", default="NBA_Draftee_Finder.html",
                    help="also write the web app here (needs nba_app_template.html next to this script); '' to skip")
    ap.add_argument("--red-seasons", type=int, default=2,
                    help="highlight players whose last season is this many or more seasons ago (default 2)")
    args = ap.parse_args()

    current_end = current_season_end(dt.date.today())
    red_cutoff = current_end - args.red_seasons
    f = Fetcher(args.delay, args.cache, args.cache_hours)

    print(f"Season {season_label(current_end)}; red = last season {season_label(red_cutoff)} or earlier")
    print("1/4 Active players")
    active = get_active_players(f)
    print("2/4 Draft pages")
    drafts = get_drafts(f, args.first_draft, current_end - 1)
    print("3/4 Advanced stats")
    career, last_seen, current_team, first_team = get_win_shares(f, args.first_draft + 1, current_end)

    print("4/4 Current rosters")
    ROSTERS.update(get_rosters(f, current_end))

    by_team = build_rows(active, drafts, career, last_seen, current_team, current_end, red_cutoff, first_team)
    undrafted = build_undrafted(active, drafts, career, last_seen, current_team, first_team, current_end, red_cutoff)
    out = os.path.expanduser(args.out)
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    write_workbook(by_team, undrafted, current_end, out)
    template = Path(__file__).with_name("nba_app_template.html")
    if args.app and template.exists():
        app = os.path.expanduser(args.app)
        Path(app).parent.mkdir(parents=True, exist_ok=True)
        write_app(by_team, undrafted, template, app, dt.date.today().strftime("%B %d, %Y").replace(" 0", " "),
                  current_end, red_cutoff)
        print(f"Web app -> {app}")
    n = sum(len(v) for v in by_team.values())
    u = sum(len(v) for v in undrafted.values())
    reds = sum(r["red"] for v in list(by_team.values()) + list(undrafted.values()) for r in v)
    print(f"Done: {n} active draftees + {u} undrafted ({reds} highlighted red) -> {out}")


if __name__ == "__main__":
    main()
