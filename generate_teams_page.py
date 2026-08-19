#!/usr/bin/env python3
"""
Indoor Football Index — teams directory generator.

Reads an 'All Teams' style sheet (columns: Team Name, League, Years,
Reg W-L-T, Playoff W-L, Total W-L-T, Win %, Name History, Location History,
Stadium History, Coach History, Championships, Logo, Previous Logos 1-5,
Logo Details, Previous Logos 1-5 Details, Info, First Season) and builds
teams.html: a searchable, league-filterable directory covering every
team in franchise history, not just currently-active ones.

Columns are matched by header name, so order doesn't matter. Team data is
embedded directly in the page (no separate fetch), same as teams.html, so
it opens fine locally without a server.

Every card links to <TeamName-with-underscores>.html, matching the naming
convention the rest of the site already uses.

Usage:
    python3 generate_teams_page.py path/to/all_teams.xlsx path/to/teams.html
    python3 generate_teams_page.py path/to/2026_IFI.xlsx "All Teams" path/to/teams.html
"""

import sys
import re
import json
import openpyxl


def slug_filename(team_name):
    safe = re.sub(r"[^A-Za-z0-9]+", "_", str(team_name)).strip("_")
    return safe + ".html"


def esc(s):
    if s is None:
        return ""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def parse_record_cell(v):
    """Recovers a W-L or W-L-T value that Excel/Sheets silently turned into
    a date (e.g. '44-50-0' -> a date it read as month/day, year defaulted
    to the current year). Returns the value unchanged if it isn't a date."""
    if v is None:
        return None
    if hasattr(v, "year") and hasattr(v, "month") and hasattr(v, "day"):
        return f"{v.month}-{v.day}"
    return v


def parse_championships_count(v):
    if v is None or str(v).strip() == "":
        return 0
    if hasattr(v, "year") and hasattr(v, "month"):
        return 1  # same 'silently turned into a date' case seen on team files
    parts = [p.strip() for p in re.split(r"[,;]", str(v)) if p.strip()]
    if not parts:
        return 0
    return int(parts[0]) if parts[0].isdigit() else len(parts)


def load_all_teams(xlsx_path, sheet_name=None):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    idx = {h: i for i, h in enumerate(headers) if h is not None}

    def get(r, col):
        return r[idx[col]] if col in idx and idx[col] < len(r) else None

    out = []
    for r in rows[1:]:
        name = get(r, "Team Name")
        if not name:
            continue
        out.append({
            "name": name,
            "league": get(r, "League"),
            "years": get(r, "Years"),
            "firstSeason": get(r, "First Season"),
            "regRecord": parse_record_cell(get(r, "Reg W-L-T")),
            "playoffRecord": parse_record_cell(get(r, "Playoff W-L")),
            "totalRecord": parse_record_cell(get(r, "Total W-L-T")),
            "winPct": get(r, "Win %"),
            "nameHistory": get(r, "Name History"),
            "championships": parse_championships_count(get(r, "Championships")),
            "logo": get(r, "Logo"),
            "info": get(r, "Info"),
            "href": slug_filename(name),
        })
    out.sort(key=lambda t: t["name"])
    return out


TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Teams — Indoor Football Index</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Anton&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>

  :root{{
    --arena-black:#111316;
    --arena-black-2:#1a1d21;
    --crimson:#D4263A;
    --amber:#E8A33D;
    --turf:#3C6E47;
    --chalk:#F2EFE6;
    --chalk-dim:#c9c6bc;
    --steel:#82868c;
    --steel-line:#2a2d31;
    --card:#1c1f23;
    --card-line:#2e3136;
  }}

  *{{box-sizing:border-box;}}
  body{{
    margin:0; background:var(--arena-black); color:var(--chalk);
    font-family:'Inter',system-ui,sans-serif; -webkit-font-smoothing:antialiased;
  }}
  .display{{font-family:'Anton',sans-serif; text-transform:uppercase; letter-spacing:0.01em;}}
  .mono{{font-family:'IBM Plex Mono',monospace;}}
  a{{color:inherit; text-decoration:none;}}

  .topbar{{display:flex; align-items:center; justify-content:space-between; max-width:1160px; margin:0 auto; padding:22px 24px 0;}}
  .topbar .site{{font-size:13px; letter-spacing:0.14em; text-transform:uppercase; color:var(--steel);}}
  .topbar .site strong{{color:var(--chalk); font-weight:600;}}
  .topbar nav{{display:flex; gap:28px; font-size:13px; letter-spacing:0.06em; text-transform:uppercase; color:var(--steel);}}
  .topbar nav a:hover{{color:var(--chalk);}}

  main{{max-width:1160px; margin:0 auto; padding:0 24px 60px;}}

  .hero{{text-align:center; padding:52px 0 30px;}}
  .hero h1{{font-size:clamp(38px,6vw,56px); margin:0 0 10px; color:var(--chalk);}}
  .hero p{{font-size:15px; color:var(--steel); margin:0;}}

  .controls{{display:flex; gap:12px; margin-bottom:28px; flex-wrap:wrap;}}
  .search-box{{flex:1 1 320px; position:relative;}}
  .search-box input{{
    width:100%; background:var(--card); border:1px solid var(--card-line); border-radius:10px;
    color:var(--chalk); padding:13px 16px 13px 42px; font-size:14px; font-family:'Inter',sans-serif;
  }}
  .search-box input:focus{{outline:1px solid var(--crimson); border-color:var(--crimson);}}
  .search-box input::placeholder{{color:var(--steel);}}
  .search-box .icon{{position:absolute; left:15px; top:50%; transform:translateY(-50%); color:var(--steel); font-size:14px;}}

  select.unit-select{{
    appearance:none; -webkit-appearance:none; background:var(--card); color:var(--chalk);
    border:1px solid var(--card-line); border-radius:10px; padding:13px 38px 13px 16px;
    font-family:'Inter',sans-serif; font-size:14px; cursor:pointer;
    background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='10' height='6'><path d='M0 0L5 6L10 0Z' fill='%2382868c'/></svg>");
    background-repeat:no-repeat; background-position:right 16px center;
  }}

  .result-count{{font-size:13px; color:var(--steel); margin-bottom:16px;}}

  .team-grid{{display:grid; grid-template-columns:repeat(4,1fr); gap:16px;}}
  @media (max-width:1000px){{.team-grid{{grid-template-columns:repeat(2,1fr);}}}}
  @media (max-width:560px){{.team-grid{{grid-template-columns:1fr;}}}}

  .team-card{{
    background:var(--card); border:1px solid var(--card-line); border-radius:12px; padding:18px;
    transition:border-color .15s; display:flex; gap:14px;
  }}
  .team-card:hover{{border-color:var(--crimson);}}
  .team-logo{{
    width:64px; height:64px; flex-shrink:0; background:var(--arena-black-2); border-radius:8px;
    display:flex; align-items:center; justify-content:center; overflow:hidden;
  }}
  .team-logo img{{width:100%; height:100%; object-fit:contain; padding:8px;}}
  .team-logo .no-logo{{font-size:9px; color:var(--steel); text-align:center; letter-spacing:0.03em;}}
  .team-info{{min-width:0; flex:1;}}
  .team-name{{font-size:15px; color:var(--chalk); font-weight:600; line-height:1.3; margin-bottom:3px;}}
  .team-meta{{font-size:12px; color:var(--steel); margin-bottom:8px;}}
  .team-meta .league{{color:var(--amber); font-family:'IBM Plex Mono',monospace;}}
  .team-stats{{display:flex; gap:12px; font-size:12px; flex-wrap:wrap;}}
  .team-stats .stat{{color:var(--chalk-dim); font-family:'IBM Plex Mono',monospace;}}
  .team-stats .stat b{{color:var(--chalk); font-weight:600;}}
  .champ-badge{{
    display:inline-flex; align-items:center; gap:3px; font-size:11px; color:var(--amber);
    background:rgba(232,163,61,0.12); padding:2px 7px; border-radius:4px; margin-top:6px; font-weight:600;
  }}

  .empty{{padding:60px 24px; text-align:center; color:var(--steel);}}

  footer{{max-width:1160px; margin:0 auto; padding:40px 24px 60px; display:flex; justify-content:space-between; flex-wrap:wrap; gap:12px; font-size:12px; color:var(--steel);}}

  ::selection{{background:var(--crimson); color:var(--chalk);}}
</style>
</head>
<body>

<div class="topbar">
  <div class="site"><strong>Indoor Football Index</strong> / Teams</div>
  <nav>
    <a href="leagues.html">Leagues</a>
    <a href="teams.html">Teams</a>
    <a href="index.html#schedule">Schedule</a>
    <a href="index.html#news">News</a>
  </nav>
</div>

<main>
  <div class="hero">
    <h1 class="display">Teams</h1>
    <p>Every indoor and arena football franchise on record &mdash; active and defunct.</p>
  </div>

  <div class="controls">
    <div class="search-box">
      <span class="icon">&#128269;</span>
      <input type="text" id="search" placeholder="Search by team, league, or former name&hellip;">
    </div>
    <select class="unit-select" id="league-filter">
      <option value="">All leagues</option>
{league_options}
    </select>
  </div>

  <div class="result-count" id="result-count"></div>
  <div class="team-grid" id="team-grid"></div>
  <div class="empty" id="empty-state" style="display:none;">No teams match that search.</div>
</main>

<footer>
  <span>Indoor Football Index &middot; Franchise archive compiled from league records.</span>
  <span>{team_count} teams</span>
</footer>

<script>
  var TEAMS = {teams_json};

  var grid = document.getElementById('team-grid');
  var countEl = document.getElementById('result-count');
  var emptyEl = document.getElementById('empty-state');
  var searchInput = document.getElementById('search');
  var leagueFilter = document.getElementById('league-filter');

  function esc(s){{
    var d = document.createElement('div');
    d.textContent = s == null ? '' : s;
    return d.innerHTML;
  }}

  function cardHtml(t){{
    var logo = t.logo
      ? '<img src="' + esc(t.logo) + '" alt="' + esc(t.name) + ' logo" onerror="this.parentElement.innerHTML=&quot;<span class=&amp;quot;no-logo&amp;quot;>No logo</span>&quot;">'
      : '<span class="no-logo">No logo</span>';
    var stats = '';
    if (t.totalRecord) stats += '<span class="stat"><b>' + esc(t.totalRecord) + '</b></span>';
    if (t.winPct) stats += '<span class="stat">' + esc(t.winPct) + '</span>';
    var champBadge = t.championships > 0
      ? '<div class="champ-badge">&#127942; ' + t.championships + (t.championships === 1 ? ' Title' : ' Titles') + '</div>'
      : '';
    return '<a class="team-card" href="' + esc(t.href) + '">' +
      '<div class="team-logo">' + logo + '</div>' +
      '<div class="team-info">' +
        '<div class="team-name">' + esc(t.name) + '</div>' +
        '<div class="team-meta"><span class="league">' + esc(t.league||'') + '</span>' + (t.years ? ' &middot; ' + esc(t.years) : '') + '</div>' +
        '<div class="team-stats">' + stats + '</div>' +
        champBadge +
      '</div>' +
      '</a>';
  }}

  function render(){{
    var q = searchInput.value.trim().toLowerCase();
    var league = leagueFilter.value;
    var filtered = TEAMS.filter(function(t){{
      var haystack = (t.name + ' ' + (t.league||'') + ' ' + (t.nameHistory||'')).toLowerCase();
      var matchesQ = !q || haystack.indexOf(q) !== -1;
      var matchesLeague = !league || t.league === league;
      return matchesQ && matchesLeague;
    }});

    countEl.textContent = filtered.length + (filtered.length === 1 ? ' team' : ' teams');
    grid.style.display = filtered.length ? 'grid' : 'none';
    emptyEl.style.display = filtered.length ? 'none' : 'block';
    grid.innerHTML = filtered.map(cardHtml).join('');
  }}

  searchInput.addEventListener('input', render);
  leagueFilter.addEventListener('change', render);
  render();
</script>

</body>
</html>
"""


def generate(xlsx_path, out_path, sheet_name=None):
    teams = load_all_teams(xlsx_path, sheet_name)
    all_leagues = sorted({t["league"] for t in teams if t["league"]})
    league_options = "\n".join(f'      <option value="{esc(l)}">{esc(l)}</option>' for l in all_leagues)

    html = TEMPLATE.format(
        league_options=league_options,
        team_count=len(teams),
        teams_json=json.dumps(teams, ensure_ascii=False, default=str),
    )
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Wrote {out_path}: {len(teams)} teams across {len(all_leagues)} leagues")


if __name__ == "__main__":
    if len(sys.argv) == 3:
        generate(sys.argv[1], sys.argv[2])
    elif len(sys.argv) == 4:
        generate(sys.argv[1], sys.argv[3], sheet_name=sys.argv[2])
    else:
        print("Usage: python3 generate_teams_page.py <workbook.xlsx> [\"Sheet Name\"] <teams.html>")
        sys.exit(1)
