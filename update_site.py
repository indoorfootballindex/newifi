#!/usr/bin/env python3
"""
Indoor Football Index — one-command site updater.

Regenerates players.json, teams.json, schedule.json, leagues.json,
awards.json, coaches.json, news.json, and index.html from your source
spreadsheets and articles. Doesn't touch git at all — commit and push
through GitHub Desktop like normal once you're happy with the changes.

Must live in the same folder as build_players_json.py, build_teams_json.py,
build_schedule_json.py, build_leagues_json.py, build_awards_json.py,
build_coaches_json.py, build_news_json.py, and generate_home_page.py — it
calls them exactly as you would by hand, in order, and stops immediately
if any required step fails.

Usage:
    python3 update_site.py
    python3 update_site.py --master Master.xlsx --ifi 2026_IFI.xlsx --leagues Leagues.xlsx --awards Awards.xlsx --coaches Coaches.xlsx --news articles --site site
"""

import argparse
import os
import subprocess
import sys


def run(cmd, description):
    print(f"\n=== {description} ===")
    print("  $ " + " ".join(cmd))
    result = subprocess.run(cmd)
    if result.returncode != 0:
        print(f"\nFAILED: {description} (exit code {result.returncode})")
        print("Stopping here — nothing after this step ran.")
        sys.exit(result.returncode)


def main():
    parser = argparse.ArgumentParser(description="Regenerate the Indoor Football Index site's data files.")
    parser.add_argument("--master", default="Master.xlsx", help="Path to the player-stats workbook (default: Master.xlsx)")
    parser.add_argument("--ifi", default="2026_IFI.xlsx", help="Path to the schedule/teams workbook (default: 2026_IFI.xlsx)")
    parser.add_argument("--leagues", default="Leagues.xlsx", help="Path to the leagues workbook (default: Leagues.xlsx)")
    parser.add_argument("--awards", default="Awards.xlsx", help="Path to the awards workbook (default: Awards.xlsx)")
    parser.add_argument("--coaches", default="Coaches.xlsx", help="Path to the coaches workbook (default: Coaches.xlsx)")
    parser.add_argument("--news", default="articles", help="Path to the folder of article .docx files (default: articles)")
    parser.add_argument("--site", default=".", help="Output folder (default: current folder)")
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    py = sys.executable

    for required, label in [(args.master, "master player-stats workbook"), (args.ifi, "schedule/teams workbook")]:
        if not os.path.isfile(required):
            print(f"ERROR: can't find the {label} at '{required}'.")
            print("Pass the right path with --master or --ifi, or run this from the folder that has them.")
            sys.exit(1)

    has_leagues = os.path.isfile(args.leagues)
    if not has_leagues:
        print(f"NOTE: no leagues workbook found at '{args.leagues}' — skipping leagues.json. "
              f"Pass --leagues path/to/file.xlsx if you have one.")

    has_awards = os.path.isfile(args.awards)
    if not has_awards:
        print(f"NOTE: no awards workbook found at '{args.awards}' — skipping awards.json. "
              f"Pass --awards path/to/file.xlsx if you have one.")

    has_coaches = os.path.isfile(args.coaches)
    if not has_coaches:
        print(f"NOTE: no coaches workbook found at '{args.coaches}' — skipping coaches.json. "
              f"Pass --coaches path/to/file.xlsx if you have one.")

    has_news = os.path.isdir(args.news) and any(fn.endswith(".docx") for fn in os.listdir(args.news))
    if not has_news:
        print(f"NOTE: no .docx articles found in '{args.news}' — writing an empty news.json. "
              f"Pass --news path/to/articles_folder if you have some.")

    os.makedirs(args.site, exist_ok=True)

    run(
        [py, os.path.join(script_dir, "build_players_json.py"), args.master, os.path.join(args.site, "players.json")],
        "Rebuilding players.json",
    )
    run(
        [py, os.path.join(script_dir, "build_teams_json.py"), args.ifi, os.path.join(args.site, "teams.json")],
        "Rebuilding teams.json",
    )
    run(
        [py, os.path.join(script_dir, "build_schedule_json.py"), args.ifi, os.path.join(args.site, "schedule.json")],
        "Rebuilding schedule.json",
    )
    if has_leagues:
        run(
            [py, os.path.join(script_dir, "build_leagues_json.py"), args.leagues, os.path.join(args.site, "leagues.json")],
            "Rebuilding leagues.json",
        )
    if has_awards:
        run(
            [py, os.path.join(script_dir, "build_awards_json.py"), args.awards, os.path.join(args.site, "awards.json"),
             os.path.join(args.site, "teams.json")],
            "Rebuilding awards.json",
        )
    if has_coaches:
        run(
            [py, os.path.join(script_dir, "build_coaches_json.py"), args.coaches, os.path.join(args.site, "coaches.json"),
             os.path.join(args.site, "teams.json")],
            "Rebuilding coaches.json",
        )
    run(
        [py, os.path.join(script_dir, "build_news_json.py"), args.news, os.path.join(args.site, "news.json")],
        "Rebuilding news.json",
    )
    run(
        [py, os.path.join(script_dir, "generate_home_page.py"), args.ifi, os.path.join(args.site, "index.html"),
         os.path.join(args.site, "news.json")],
        "Rebuilding index.html",
    )

    print("\n=== Done ===")
    print(f"Updated files are in '{args.site}'. Review and commit/push through GitHub Desktop when ready.")


if __name__ == "__main__":
    main()

