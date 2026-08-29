#!/usr/bin/env python3
"""
Indoor Football Index — full schedule data builder.

Builds schedule.json for schedule.html: every game in the Schedule tab,
past and future, across every league and season. Uses short property names
to keep the file lean at 14,000+ rows — schedule.html expands them into
readable fields. Team logos aren't included here; schedule.html cross-
references team names against teams.json for those.

Usage:
    python3 build_schedule_json.py path/to/2026_IFI.xlsx path/to/schedule.json
"""

import sys
import re
import json
import openpyxl


def fmt_num(v):
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if re.fullmatch(r"-?\d+", s):
            return int(s)
        if re.fullmatch(r"-?\d+\.\d+", s):
            v = float(s)
        else:
            return None  # not actually numeric — treat like a missing score
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return v


def norm_league(v):
    """League codes should be uppercase acronyms, but source sheets
    occasionally have stray lowercase entries (e.g. 'af2' next to 'AF2')."""
    return v.strip().upper() if isinstance(v, str) and v.strip() else v


def main():
    if len(sys.argv) != 3:
        print("Usage: python3 build_schedule_json.py <2026_IFI.xlsx> <schedule.json>")
        sys.exit(1)
    xlsx_path, out_path = sys.argv[1], sys.argv[2]

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Schedule"]
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    idx = {h: i for i, h in enumerate(headers) if h is not None}

    games = []
    for r in rows[1:]:
        date = r[idx["Date"]]
        home = r[idx["Home Team"]]
        away = r[idx["Away Team"]]
        if not date or not home or not away:
            continue

        home_score_raw = r[idx["Home Score"]] if "Home Score" in idx else None
        away_score_raw = r[idx["Away Score"]] if "Away Score" in idx else None
        # A game is "played" if the source has ANY result recorded, even if
        # it's just text (W/L/T) instead of a real number — some older
        # games (mostly LFL-era) only ever had the outcome recorded, not
        # the final score. Those still happened; they just can't be shown
        # with a numeric score or counted in the Scorigami grid.
        played = home_score_raw is not None and away_score_raw is not None
        home_score = fmt_num(home_score_raw)
        away_score = fmt_num(away_score_raw)

        week = r[idx["Week"]] if "Week" in idx else None
        week_low = str(week).lower()
        # Different leagues have branded their title game differently over
        # the years (ArenaBowl for the AFL, ArenaCup for af2, IFL Cup
        # currently) rather than just calling it "Championship", so treat
        # any of those patterns as a championship game.
        is_championship = any(k in week_low for k in ("championship", "bowl", "cup"))
        is_playoff = is_championship or any(k in week_low for k in ("playoff", "semifinal", "quarterfinal"))

        games.append({
            "d": date.strftime("%Y-%m-%d"),
            "dow": r[idx["Day of the Week"]] if "Day of the Week" in idx else None,
            "h": home,
            "a": away,
            "hs": home_score,
            "as": away_score,
            "lg": norm_league(r[idx["League"]]) if "League" in idx else None,
            "wk": week,
            "tm": r[idx["Time (CST)"]] if "Time (CST)" in idx else None,
            "w": r[idx["Watch"]] if "Watch" in idx else None,
            "p": played,
            "po": is_playoff,
            "ch": is_championship,
        })

    games.sort(key=lambda g: g["d"])

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(games, f, ensure_ascii=False, default=str)

    played_count = sum(1 for g in games if g["p"])
    print(f"Wrote {out_path}: {len(games)} total games ({played_count} played, {len(games) - played_count} upcoming)")


if __name__ == "__main__":
    main()
