#!/usr/bin/env python3
"""
Indoor Football Index — site search index builder.

Combines players.json, teams.json, leagues.json, and coaches.json into one
small search-index.json for the site-wide search bar. Deliberately lean
(name, type, link, one line of context per entry) rather than reusing the
full data files directly, so every page can afford to load it — a page
like schedule.html or news.html shouldn't need to pull in the full player
stats file just so search works.

Any of the four source files can be missing — that category is just
skipped rather than causing an error, so this works fine even on a fresh
site that doesn't have everything built yet.

Usage:
    python3 build_search_index.py path/to/site_folder path/to/search-index.json
"""

import sys
import os
import json


def load_json(path):
    if not os.path.isfile(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def build_players(players):
    if players is None:
        return []
    by_slug = {}
    for r in players:
        existing = by_slug.get(r["slug"])
        if not existing or (r.get("Season") or 0) > (existing.get("Season") or 0):
            by_slug[r["slug"]] = r
    out = []
    for slug, r in by_slug.items():
        pos = r.get("Pos")
        meta = str(pos) if pos else ""
        if r.get("team"):
            meta = (meta + " \u00b7 " + r["team"]) if meta else r["team"]
        out.append({"name": r["name"], "type": "Player", "href": f"player.html?player={slug}", "meta": meta})
    return out


def build_teams(teams):
    if teams is None:
        return []
    out = []
    for t in teams:
        meta = t.get("league") or ""
        aliases = []
        if t.get("nameHistory"):
            for entry in t["nameHistory"].split(","):
                # strip the trailing "(year-year)" part, keep just the name
                historical_name = entry.split("(")[0].strip()
                if historical_name and historical_name != t["name"] and historical_name not in aliases:
                    aliases.append(historical_name)
        entry = {"name": t["name"], "type": "Team", "href": f"team.html?team={t['slug']}", "meta": meta}
        if aliases:
            entry["aliases"] = aliases
        out.append(entry)
    return out


def build_leagues(leagues):
    if leagues is None:
        return []
    out = []
    for lg in leagues:
        meta = lg.get("acronym") or ""
        out.append({"name": lg["name"], "type": "League", "href": f"league.html?league={lg['slug']}", "meta": meta})
    return out


def build_coaches(coaches):
    if coaches is None:
        return []
    out = []
    for c in coaches:
        meta = c.get("activeTeam") or (c["teamHistory"][-1]["name"] if c.get("teamHistory") else "")
        out.append({"name": c["name"], "type": "Coach", "href": f"player.html?player={c['slug']}", "meta": meta})
    return out


def main():
    if len(sys.argv) != 3:
        print("Usage: python3 build_search_index.py <site_folder> <search-index.json>")
        sys.exit(1)
    site_dir, out_path = sys.argv[1], sys.argv[2]

    players = load_json(os.path.join(site_dir, "players.json"))
    teams = load_json(os.path.join(site_dir, "teams.json"))
    leagues = load_json(os.path.join(site_dir, "leagues.json"))
    coaches = load_json(os.path.join(site_dir, "coaches.json"))

    index = (
        build_teams(teams) +
        build_players(players) +
        build_leagues(leagues) +
        build_coaches(coaches)
    )
    index.sort(key=lambda e: e["name"])

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False)

    counts = {}
    for e in index:
        key = e["type"] + ("es" if e["type"] == "Coach" else "s")
        counts[key] = counts.get(key, 0) + 1
    print(f"Wrote {out_path}: {len(index)} entries ({', '.join(f'{v} {k}' for k,v in counts.items())})")


if __name__ == "__main__":
    main()
