#!/usr/bin/env python3
"""
Indoor Football Index — awards data builder.

Reads Awards.xlsx (League, Year, Award, Name, Team, Player Position,
Week (if applicable)) and builds awards.json for awards.html and for the
Awards section on player.html. Coach-specific awards (Coach of the Year,
Assistant Coach of the Year) are tagged so they're ready to show on coach
profile pages once those exist.

Usage:
    python3 build_awards_json.py path/to/Awards.xlsx path/to/awards.json
"""

import sys
import os
import re
import json
import openpyxl

COACH_AWARD_KEYWORDS = ("coach",)


def slug(name):
    return re.sub(r"[^a-z0-9]+", "-", str(name).lower()).strip("-")


def fmt_num(v):
    if v is None:
        return None
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return v


def build_team_resolver(teams_json_path):
    """Awards.xlsx is inconsistent about team names — weekly awards use the
    full name ('Green Bay Blizzard'), season-end awards often use just the
    city ('Green Bay'). Resolves either form to the real team's slug by
    matching against teams.json, so links work regardless of which form a
    given row used. Returns a function name -> slug_or_None."""
    if not teams_json_path or not os.path.isfile(teams_json_path):
        return lambda name: slug(name)

    with open(teams_json_path, encoding="utf-8") as f:
        teams = json.load(f)

    exact = {t["name"].strip().lower(): t["slug"] for t in teams}
    # for prefix matching, prefer the longest full name so "San Diego" doesn't
    # accidentally match a shorter unrelated team that also starts that way
    by_length = sorted(teams, key=lambda t: -len(t["name"]))

    def resolve(name):
        if not name:
            return None
        key = str(name).strip().lower()
        if key in exact:
            return exact[key]
        for t in by_length:
            if t["name"].strip().lower().startswith(key):
                return t["slug"]
        return slug(name)  # no match found — fall back to naive slug

    return resolve


def main():
    if len(sys.argv) not in (3, 4):
        print("Usage: python3 build_awards_json.py <Awards.xlsx> <awards.json> [teams.json]")
        sys.exit(1)
    xlsx_path, out_path = sys.argv[1], sys.argv[2]
    teams_json_path = sys.argv[3] if len(sys.argv) == 4 else None
    resolve_team_slug = build_team_resolver(teams_json_path)

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    idx = {h: i for i, h in enumerate(headers) if h is not None}

    required = ["League", "Year", "Award", "Name", "Team"]
    missing = [c for c in required if c not in idx]
    if missing:
        print(f"ERROR: missing required column(s): {missing}")
        print(f"Headers found: {list(idx.keys())}")
        sys.exit(1)

    def get(r, col):
        return r[idx[col]] if col in idx and idx[col] < len(r) else None

    awards = []
    for r in rows[1:]:
        name = get(r, "Name")
        award = get(r, "Award")
        year = get(r, "Year")
        if not name or not award or year is None:
            continue
        year = str(int(year)) if isinstance(year, (int, float)) else str(year)
        award_low = str(award).lower()
        is_coach = any(k in award_low for k in COACH_AWARD_KEYWORDS)
        awards.append({
            "league": get(r, "League"),
            "year": year,
            "award": award,
            "name": name,
            "slug": slug(name),
            "team": get(r, "Team"),
            "teamSlug": resolve_team_slug(get(r, "Team")),
            "position": get(r, "Player Position"),
            "week": fmt_num(get(r, "Week (if applicable)")),
            "isCoachAward": is_coach,
        })

    awards.sort(key=lambda a: (a["year"], a["week"] or 0), reverse=True)

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(awards, f, ensure_ascii=False)

    weekly = sum(1 for a in awards if a["week"])
    season_end = len(awards) - weekly
    print(f"Wrote {out_path}: {len(awards)} awards ({weekly} weekly, {season_end} season-end/All-Pro)")


if __name__ == "__main__":
    main()
