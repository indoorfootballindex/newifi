#!/usr/bin/env python3
"""
Indoor Football Index — team data consolidator.

Builds the shared teams.json that team.html reads from at runtime, keyed
by a slugified team name. Combines two sources from the master workbook:

  - the 'All Teams' tab: bio data (league, years, records, name/location/
    stadium/coach history, championships, logos, info)
  - the 'Schedule' tab: game-by-game results, matched to each team via
    'Current Home Name' / 'Current Away Name' (so a franchise's full
    history follows it across any past rebrands) and pre-grouped into
    seasons here at build time, so team.html never has to fetch or filter
    the full 14,000+ row schedule itself.

Usage:
    python3 build_teams_json.py path/to/2026_IFI.xlsx path/to/teams.json
"""

import sys
import re
import json
import datetime
import openpyxl
from collections import defaultdict


def slug(name):
    return re.sub(r"[^a-z0-9]+", "-", str(name).lower()).strip("-")


def esc_none(v):
    return v if v not in (None, "") else None


def parse_record_cell(v):
    """Recovers a W-L(-T) value Excel/Sheets silently turned into a date."""
    if v is None:
        return None
    if hasattr(v, "year") and hasattr(v, "month") and hasattr(v, "day"):
        return f"{v.month}-{v.day}"
    return v


def parse_championships(v):
    if v is None or str(v).strip() == "":
        return 0, []
    if hasattr(v, "year") and hasattr(v, "month"):
        return 1, [str(v.year)]
    parts = [p.strip() for p in re.split(r"[,;]", str(v)) if p.strip()]
    if not parts:
        return 0, []
    count = int(parts[0]) if parts[0].isdigit() else len(parts)
    years = parts[1:] if parts[0].isdigit() else parts
    return count, years


def build_coach_list(coach_history):
    if not coach_history:
        return []
    out = []
    for part in str(coach_history).split(","):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(.*?)\s*\(([^)]+)\)$", part)
        if m:
            out.append({"name": m.group(1).strip(), "tenure": m.group(2).strip()})
        else:
            out.append({"name": part, "tenure": ""})
    return out


def load_all_teams_bio(wb):
    ws = wb["All Teams"]
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    idx = {h: i for i, h in enumerate(headers) if h is not None}

    def get(r, col):
        return r[idx[col]] if col in idx and idx[col] < len(r) else None

    bio = {}
    for r in rows[1:]:
        name = get(r, "Team Name")
        if not name:
            continue
        champ_count, champ_years = parse_championships(get(r, "Championships"))

        logos = []
        current_logo = get(r, "Logo")
        if current_logo:
            logos.append({
                "url": current_logo,
                "caption": esc_none(get(r, "Logo Details")),
                "current": True,
            })
        for n in range(1, 6):
            url = get(r, f"Previous Logos {n}")
            if url:
                logos.append({
                    "url": url,
                    "caption": esc_none(get(r, f"Previous Logos {n} Details")),
                    "current": False,
                })

        bio[name] = {
            "name": name,
            "slug": slug(name),
            "league": esc_none(get(r, "League")),
            "years": esc_none(get(r, "Years")),
            "firstSeason": get(r, "First Season"),
            "regRecord": parse_record_cell(get(r, "Reg W-L-T")),
            "playoffRecord": parse_record_cell(get(r, "Playoff W-L")),
            "totalRecord": parse_record_cell(get(r, "Total W-L-T")),
            "winPct": fmt_pct(get(r, "Win %")),
            "nameHistory": esc_none(get(r, "Name History")),
            "locationHistory": esc_none(get(r, "Location History")),
            "stadiumHistory": esc_none(get(r, "Stadium History")),
            "coaches": build_coach_list(get(r, "Coach History")),
            "championships": champ_count,
            "championshipYears": champ_years,
            "logos": logos,
            "info": esc_none(get(r, "Info")),
            "seasons": {},
        }
    return bio


def load_schedule_by_team(wb, team_names):
    ws = wb["Schedule"]
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    idx = {h: i for i, h in enumerate(headers) if h is not None}
    CUR_HOME_COL = idx.get("Current Home Name", 17)
    CUR_AWAY_COL = CUR_HOME_COL + 1

    games_by_team = defaultdict(list)
    for r in rows[1:]:
        date = r[idx["Date"]]
        if date is None:
            continue
        cur_home = r[CUR_HOME_COL] if CUR_HOME_COL < len(r) else None
        cur_away = r[CUR_AWAY_COL] if CUR_AWAY_COL < len(r) else None
        home_team = r[idx["Home Team"]]
        away_team = r[idx["Away Team"]]
        home_score = r[idx["Home Score"]]
        away_score = r[idx["Away Score"]]
        if home_score is None or away_score is None:
            continue
        week = r[idx["Week"]] if "Week" in idx else None
        week_low = str(week).lower()
        is_playoff = any(k in week_low for k in ("playoff", "championship", "semifinal", "quarterfinal"))
        year = str(date.year)

        for team_name, is_home in ((cur_home, True), (cur_away, False)):
            if team_name not in team_names:
                continue
            opp = away_team if is_home else home_team
            own_score = home_score if is_home else away_score
            opp_score = away_score if is_home else home_score
            tie_col = idx.get("Tie")
            tie = (tie_col is not None and r[tie_col] == "Y")
            wl = "T" if tie else ("W" if own_score > opp_score else "L")
            games_by_team[team_name].append({
                "date": date,
                "dateLabel": date.strftime("%b %d"),
                "opp": opp,
                "home": is_home,
                "score": f"{fmt_score(own_score)}-{fmt_score(opp_score)}",
                "wl": wl,
                "playoff": is_playoff,
                "year": year,
            })
    for t in games_by_team:
        games_by_team[t].sort(key=lambda g: g["date"])
    return games_by_team


def fmt_score(v):
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return str(v)


def fmt_pct(v):
    """Excel stores percentages as raw decimals (0.475); the '%' is just a
    display format, which openpyxl doesn't expose via data_only. Convert it
    back to a percent string here so the page doesn't have to guess."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if v <= 1:
            v = v * 100
        return f"{v:.1f}%"
    return v


def season_summary(games):
    w = l = t = 0
    for g in games:
        if g["wl"] == "W":
            w += 1
        elif g["wl"] == "L":
            l += 1
        else:
            t += 1
    return f"{w}-{l}" + (f"-{t}" if t else "")


def build_seasons(games):
    by_year = defaultdict(list)
    for g in games:
        by_year[g["year"]].append(g)
    seasons = {}
    for year, yr_games in by_year.items():
        reg_games = [g for g in yr_games if not g["playoff"]]
        post_games = [g for g in yr_games if g["playoff"]]
        seasons[year] = {
            "reg": season_summary(reg_games),
            "post": season_summary(post_games) if post_games else None,
            "games": [
                {
                    "date": g["dateLabel"] + (" \u00b7 PO" if g["playoff"] else ""),
                    "opp": ("vs " if g["home"] else "at ") + g["opp"],
                    "score": g["score"].replace("-", "\u2013"),
                    "wl": g["wl"],
                }
                for g in yr_games
            ],
        }
    return seasons


def main():
    if len(sys.argv) != 3:
        print("Usage: python3 build_teams_json.py <2026_IFI.xlsx> <teams.json>")
        sys.exit(1)
    xlsx_path, out_path = sys.argv[1], sys.argv[2]

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    bio = load_all_teams_bio(wb)
    print(f"Loaded {len(bio)} teams from 'All Teams' tab")

    games_by_team = load_schedule_by_team(wb, set(bio.keys()))
    matched = 0
    for name, team in bio.items():
        games = games_by_team.get(name, [])
        team["seasons"] = build_seasons(games)
        if games:
            matched += 1
    print(f"Matched schedule data for {matched}/{len(bio)} teams")

    out = list(bio.values())
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, default=str)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
