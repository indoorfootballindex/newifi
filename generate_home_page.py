#!/usr/bin/env python3
"""
Indoor Football Index — home page generator.

Reads the 'Schedule' tab of a workbook like 2026_IFI.xlsx and builds index.html:
  - Recent games: the 10 most recently completed games before today
  - Upcoming schedule: the 10 next unplayed games from today forward

News stays hand-edited (it's editorial content, not schedule data) — see the
NEWS_ITEMS list below to update it.

Usage:
    python3 generate_home_page.py path/to/2026_IFI.xlsx path/to/index.html
"""

import sys
import datetime
import openpyxl

# Edit these by hand whenever you want to change the News section.
NEWS_ITEMS = [
    {
        "kicker": "Season recap", "feature": True,
        "title": "Strike Force's historic season ends in the West Championship",
        "body": "San Diego closed the 2026 regular season 13\u20133 \u2014 the best record in franchise history \u2014 and claimed the #1 seed in the West. After a first-round win over Tucson, the run ended one game short, falling 37\u201343 to eventual champion Arizona in the conference title game.",
        "date": "Aug 10, 2026",
    },
    {
        "kicker": "Franchise records", "feature": False,
        "title": "Marques Rodgers still owns the Strike Force record book",
        "body": "Five years after his final season in San Diego, Rodgers remains the franchise leader in receiving yards, receiving touchdowns, receptions, all-purpose yards, and kickoff return yards.",
        "date": "Aug 6, 2026",
    },
    {
        "kicker": "Looking back", "feature": False,
        "title": "Remembering the Panthers' 2023 title run",
        "body": "Bay Area went 10\u20135 in the regular season before running the postseason table at 3\u20130, capturing the only championship in franchise history to date.",
        "date": "Aug 3, 2026",
    },
    {
        "kicker": "Site update", "feature": False,
        "title": "Welcome to the new Indoor Football Index",
        "body": "Franchise histories, record books, and full box scores are going up team by team. More franchises, more seasons, and league-wide stats are on the way.",
        "date": "Aug 19, 2026",
    },
]


def esc(s):
    if s is None:
        return ""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def fmt_score(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return esc(v)


def load_schedule(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Schedule"]
    rows = list(ws.iter_rows(values_only=True))
    headers = rows[0]
    idx = {h: i for i, h in enumerate(headers) if h is not None}
    games = []
    for r in rows[1:]:
        date = r[idx["Date"]]
        if date is None:
            continue
        home = r[idx["Home Team"]]
        away = r[idx["Away Team"]]
        if not home or not away:
            continue
        home_score = r[idx["Home Score"]] if "Home Score" in idx else None
        away_score = r[idx["Away Score"]] if "Away Score" in idx else None
        games.append({
            "date": date,
            "home": home,
            "away": away,
            "time": r[idx["Time (CST)"]] if "Time (CST)" in idx else None,
            "week": r[idx["Week"]] if "Week" in idx else None,
            "league": r[idx["League"]] if "League" in idx else None,
            "home_score": home_score,
            "away_score": away_score,
            "played": home_score is not None and away_score is not None,
        })
    return games


def recent_card(g):
    date_label = g["date"].strftime("%b %-d" if sys.platform != "win32" else "%b %#d")
    week = esc(g["week"]) if g["week"] else ""
    league = esc(g["league"]) if g["league"] else ""
    return f'''    <div class="game-card">
      <div class="meta"><span>{date_label}{" &middot; " + week if week else ""}</span><span class="tag">{league}</span></div>
      <div class="matchup">{esc(g["away"])}<br>at <b>{esc(g["home"])}</b></div>
      <div class="score-row"><span class="score">{fmt_score(g["away_score"])}&ndash;{fmt_score(g["home_score"])}</span></div>
    </div>'''


def upcoming_card(g):
    date_label = g["date"].strftime("%b %-d" if sys.platform != "win32" else "%b %#d")
    week = esc(g["week"]) if g["week"] else ""
    league = esc(g["league"]) if g["league"] else ""
    time_label = esc(g["time"]) if g["time"] else "TBD"
    return f'''    <div class="game-card">
      <div class="meta"><span>{date_label}{" &middot; " + week if week else ""}</span><span class="tag">{league}</span></div>
      <div class="matchup">{esc(g["away"])}<br>at <b>{esc(g["home"])}</b></div>
      <div class="score-row"><span class="score mono" style="font-size:14px; color:var(--steel);">{time_label}</span></div>
    </div>'''


def build_news_html():
    cards = []
    for item in NEWS_ITEMS:
        cls = "news-card feature" if item["feature"] else "news-card"
        cards.append(
            f'    <div class="{cls}">\n'
            f'      <div class="kicker">{esc(item["kicker"])}</div>\n'
            f'      <h3>{esc(item["title"])}</h3>\n'
            f'      <p>{esc(item["body"])}</p>\n'
            f'      <div class="date">{esc(item["date"])}</div>\n'
            f'    </div>'
        )
    return "\n".join(cards)


TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Indoor Football Index</title>
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
  html{{scroll-behavior:smooth;}}
  body{{
    margin:0;
    background:var(--arena-black);
    color:var(--chalk);
    font-family:'Inter',system-ui,sans-serif;
    -webkit-font-smoothing:antialiased;
  }}

  .display{{font-family:'Anton',sans-serif; text-transform:uppercase; letter-spacing:0.01em;}}
  .mono{{font-family:'IBM Plex Mono',monospace;}}
  a{{color:inherit; text-decoration:none;}}

  .hashes{{width:100%; height:14px; display:flex; align-items:center;}}
  .hashes svg{{width:100%; height:100%; display:block;}}

  .topbar{{
    display:flex; align-items:center; justify-content:space-between;
    max-width:1160px; margin:0 auto; padding:22px 24px 0;
  }}
  .topbar .site{{font-size:13px; letter-spacing:0.14em; text-transform:uppercase; color:var(--steel);}}
  .topbar .site strong{{color:var(--chalk); font-weight:600;}}
  .topbar nav{{display:flex; gap:28px; font-size:13px; letter-spacing:0.06em; text-transform:uppercase; color:var(--steel);}}
  .topbar nav a:hover{{color:var(--chalk);}}

  .hero{{max-width:1160px; margin:0 auto; padding:64px 24px 44px;}}
  .hero .eyebrow{{
    font-size:12px; letter-spacing:0.2em; text-transform:uppercase; color:var(--crimson);
    font-weight:600; margin:0 0 14px;
  }}
  .hero h1{{font-size:clamp(44px,7vw,84px); line-height:0.92; margin:0 0 18px; color:var(--chalk);}}
  .hero p{{font-size:16px; color:var(--steel); max-width:560px; line-height:1.7; margin:0 0 26px;}}
  .hero .cta{{
    display:inline-flex; align-items:center; gap:8px; background:var(--crimson); color:var(--chalk);
    font-size:14px; font-weight:600; padding:13px 22px; border-radius:8px;
  }}
  .hero .cta:hover{{background:#b81f31;}}

  section{{max-width:1160px; margin:0 auto; padding:48px 24px;}}
  .section-head{{display:flex; align-items:baseline; justify-content:space-between; margin-bottom:22px; gap:16px; flex-wrap:wrap;}}
  .section-head h2{{font-size:24px; margin:0; color:var(--chalk);}}
  .section-head .note{{font-size:13px; color:var(--steel);}}

  .hrow{{
    display:flex; gap:14px; overflow-x:auto; padding-bottom:6px; scroll-snap-type:x proximity;
  }}
  .hrow::-webkit-scrollbar{{height:6px;}}
  .hrow::-webkit-scrollbar-thumb{{background:var(--card-line); border-radius:6px;}}

  .game-card{{
    flex:0 0 220px; background:var(--card); border:1px solid var(--card-line); border-radius:10px;
    padding:16px 18px; scroll-snap-align:start;
  }}
  .game-card .meta{{font-size:11px; letter-spacing:0.08em; text-transform:uppercase; color:var(--steel); margin-bottom:12px; display:flex; justify-content:space-between;}}
  .game-card .meta .tag{{color:var(--amber);}}
  .game-card .matchup{{font-size:14px; color:var(--chalk-dim); line-height:1.6;}}
  .game-card .matchup b{{color:var(--chalk); font-weight:600;}}
  .game-card .score-row{{display:flex; align-items:center; justify-content:space-between; margin-top:10px;}}
  .game-card .score{{font-family:'IBM Plex Mono',monospace; font-size:20px; color:var(--chalk);}}
  .game-card .wl{{
    font-family:'IBM Plex Mono',monospace; font-size:11px; font-weight:600; padding:3px 8px; border-radius:4px;
  }}
  .wl.w{{background:rgba(60,110,71,0.25); color:#7fc492;}}
  .wl.l{{background:rgba(212,38,58,0.2); color:#f0919c;}}

  .tba-strip{{
    border:1px dashed var(--card-line); border-radius:10px; padding:28px 24px; text-align:center;
  }}
  .tba-strip .big{{font-size:15px; color:var(--chalk-dim); margin-bottom:6px;}}
  .tba-strip .small{{font-size:13px; color:var(--steel);}}

  .news-grid{{display:grid; grid-template-columns:1.3fr 1fr 1fr; gap:16px;}}
  .news-card{{
    background:var(--card); border:1px solid var(--card-line); border-radius:10px;
    padding:22px; display:flex; flex-direction:column; gap:10px;
  }}
  .news-card.feature{{grid-row:span 2; padding:28px;}}
  .news-card .kicker{{font-size:11px; letter-spacing:0.1em; text-transform:uppercase; color:var(--crimson); font-weight:600;}}
  .news-card h3{{font-size:18px; margin:0; color:var(--chalk); line-height:1.35;}}
  .news-card.feature h3{{font-size:26px;}}
  .news-card p{{font-size:14px; color:var(--chalk-dim); line-height:1.65; margin:0; flex:1;}}
  .news-card .date{{font-size:12px; color:var(--steel); font-family:'IBM Plex Mono',monospace; margin-top:auto;}}
  @media (max-width:900px){{
    .news-grid{{grid-template-columns:1fr 1fr;}}
    .news-card.feature{{grid-row:span 1; grid-column:span 2;}}
  }}
  @media (max-width:600px){{
    .news-grid{{grid-template-columns:1fr;}}
    .news-card.feature{{grid-column:span 1;}}
  }}

  footer{{
    max-width:1160px; margin:0 auto; padding:40px 24px 60px; display:flex; justify-content:space-between;
    flex-wrap:wrap; gap:12px; font-size:12px; color:var(--steel);
  }}

  ::selection{{background:var(--crimson); color:var(--chalk);}}
</style>
</head>
<body>

<div class="topbar">
  <div class="site"><strong>Indoor Football Index</strong></div>
  <nav>
    <a href="leagues.html">Leagues</a>
    <a href="teams.html">Teams</a>
    <a href="#schedule">Schedule</a>
    <a href="#news">News</a>
  </nav>
</div>

<header class="hero">
  <p class="eyebrow">Indoor &amp; arena football, archived</p>
  <h1 class="display">Every team.<br>Every season.<br>One index.</h1>
  <p>Franchise histories, record books, and box scores for indoor and arena football &mdash; built from the ground up, one team at a time.</p>
  <a class="cta" href="teams.html">Browse teams &rarr;</a>
</header>

<div class="hashes"><svg preserveAspectRatio="none" viewBox="0 0 1160 14"><line x1="0" y1="7" x2="1160" y2="7" stroke="#2a2d31" stroke-width="1"/></svg></div>

<section id="schedule">
  <div class="section-head">
    <h2 class="display">Upcoming schedule</h2>
    <p class="note">The next {upcoming_count} games across every league in the archive.</p>
  </div>
{upcoming_html}
</section>

<div class="hashes"><svg preserveAspectRatio="none" viewBox="0 0 1160 14"><line x1="0" y1="7" x2="1160" y2="7" stroke="#2a2d31" stroke-width="1"/></svg></div>

<section id="recent">
  <div class="section-head">
    <h2 class="display">Recent games</h2>
    <p class="note">The last {recent_count} completed games across every league in the archive.</p>
  </div>
{recent_html}
</section>

<div class="hashes"><svg preserveAspectRatio="none" viewBox="0 0 1160 14"><line x1="0" y1="7" x2="1160" y2="7" stroke="#2a2d31" stroke-width="1"/></svg></div>

<section id="news">
  <div class="section-head">
    <h2 class="display">News</h2>
  </div>
  <div class="news-grid">
{news_html}
  </div>
</section>

<footer>
  <span>Indoor Football Index &middot; Team histories, records, and box scores.</span>
  <span>Data compiled from IFL box scores and franchise archives.</span>
</footer>

</body>
</html>
"""


def generate(xlsx_path, out_path):
    games = load_schedule(xlsx_path)
    today = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    completed = [g for g in games if g["played"] and g["date"] <= today]
    completed.sort(key=lambda g: g["date"], reverse=True)
    recent = completed[:10]

    upcoming = [g for g in games if not g["played"] and g["date"] >= today]
    upcoming.sort(key=lambda g: g["date"])
    upcoming = upcoming[:10]

    if recent:
        recent_html = '  <div class="hrow">\n' + "\n".join(recent_card(g) for g in recent) + "\n  </div>"
    else:
        recent_html = ('  <div class="tba-strip"><div class="big">No completed games found before today.'
                        '</div><div class="small">Check that the Schedule tab has scores filled in.</div></div>')

    if upcoming:
        upcoming_html = '  <div class="hrow">\n' + "\n".join(upcoming_card(g) for g in upcoming) + "\n  </div>"
    else:
        upcoming_html = ('  <div class="tba-strip"><div class="big">No upcoming games found.</div>'
                          '<div class="small">Check that the Schedule tab has future dates without scores.</div></div>')

    html = TEMPLATE.format(
        upcoming_count=len(upcoming),
        recent_count=len(recent),
        upcoming_html=upcoming_html,
        recent_html=recent_html,
        news_html=build_news_html(),
    )

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Wrote {out_path} ({len(recent)} recent, {len(upcoming)} upcoming)")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 generate_home_page.py <schedule.xlsx> <output.html>")
        sys.exit(1)
    generate(sys.argv[1], sys.argv[2])
