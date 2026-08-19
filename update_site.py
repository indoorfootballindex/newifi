#!/usr/bin/env python3
"""
Indoor Football Index — one-command site updater.

Regenerates players.json, teams.json, and index.html from your source
spreadsheets. Doesn't touch git at all — commit and push through GitHub
Desktop like normal once you're happy with the changes.

Must live in the same folder as build_players_json.py, build_teams_json.py,
and generate_home_page.py — it calls them exactly as you would by hand, in
order, and stops immediately if any step fails.

Usage:
    python3 update_site.py
    python3 update_site.py --master Master.xlsx --ifi 2026_IFI.xlsx --site site
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
    parser.add_argument("--site", default="site", help="Output folder (default: site)")
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    py = sys.executable

    for required, label in [(args.master, "master player-stats workbook"), (args.ifi, "schedule/teams workbook")]:
        if not os.path.isfile(required):
            print(f"ERROR: can't find the {label} at '{required}'.")
            print("Pass the right path with --master or --ifi, or run this from the folder that has them.")
            sys.exit(1)

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
        [py, os.path.join(script_dir, "generate_home_page.py"), args.ifi, os.path.join(args.site, "index.html")],
        "Rebuilding index.html",
    )

    print("\n=== Done ===")
    print(f"Updated files are in '{args.site}'. Review and commit/push through GitHub Desktop when ready.")


if __name__ == "__main__":
    main()
