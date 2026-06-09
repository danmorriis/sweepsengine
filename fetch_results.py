"""
fetch_results.py
----------------
Fetches completed World Cup 2026 match results from ESPN's unofficial API
and writes home/away scores into columns D and E of the FixturesResults sheet
in World Cup 2026 Sweeps.xlsx.

Run manually:  python3 fetch_results.py
Scheduled:     GitHub Actions (.github/workflows/daily_update.yml)
"""

import json
import os
import urllib.request
from datetime import date, timedelta
import openpyxl

# ── Config ────────────────────────────────────────────────────────────────────
XLSX_PATH = os.environ.get(
    "XLSX_PATH",
    os.path.join(os.path.dirname(__file__), "World Cup 2026 Sweeps.xlsx"),
)
TOURNAMENT_START = date(2026, 6, 11)
TOURNAMENT_END   = date(2026, 7, 19)

# ── Team name mapping (ESPN displayName → sheet name) ─────────────────────────
ESPN_TO_SHEET = {
    "Bosnia-Herzegovina": "Bosnia and Herzegovina",
    "Congo DR":           "DR Congo",
    "Türkiye":            "Turkey",
    "United States":      "USA",
}

# ── ESPN API helpers ──────────────────────────────────────────────────────────
BASE_URL = "https://site.api.espn.com/apis/site/v2/sports/soccer/fifa.world/scoreboard"

def fetch_scores_for_date(d: date) -> list[dict]:
    """
    Returns a list of completed matches for the given date:
    [{"home": "Mexico", "away": "South Africa", "home_score": 2, "away_score": 0}, ...]
    Only includes matches whose status is 'Final'.
    """
    url = f"{BASE_URL}?dates={d.strftime('%Y%m%d')}"
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            data = json.load(resp)
    except Exception as e:
        print(f"  Warning: could not fetch {d} — {e}")
        return []

    results = []
    for event in data.get("events", []):
        comp = event["competitions"][0]
        status = comp["status"]["type"]["name"]  # e.g. "STATUS_FINAL", "STATUS_SCHEDULED"
        if status != "STATUS_FINAL":
            continue

        home_comp = next((t for t in comp["competitors"] if t["homeAway"] == "home"), None)
        away_comp = next((t for t in comp["competitors"] if t["homeAway"] == "away"), None)
        if not home_comp or not away_comp:
            continue

        home_name = ESPN_TO_SHEET.get(home_comp["team"]["displayName"], home_comp["team"]["displayName"])
        away_name = ESPN_TO_SHEET.get(away_comp["team"]["displayName"], away_comp["team"]["displayName"])

        try:
            home_score = int(home_comp["score"])
            away_score = int(away_comp["score"])
        except (KeyError, ValueError, TypeError):
            continue

        results.append({
            "home":       home_name,
            "away":       away_name,
            "home_score": home_score,
            "away_score": away_score,
        })

    return results


# ── Main update logic ─────────────────────────────────────────────────────────
def main():
    print(f"Loading: {XLSX_PATH}")
    wb = openpyxl.load_workbook(XLSX_PATH)
    ws = wb["FixturesResults"]

    # Build index: (home_team, away_team) → row number
    fixture_index: dict[tuple[str, str], int] = {}
    for row in range(3, 200):
        home = ws.cell(row, 3).value  # col C
        away = ws.cell(row, 6).value  # col F
        if home and away:
            fixture_index[(home, away)] = row

    print(f"Indexed {len(fixture_index)} fixtures in the sheet.")

    # Fetch results for every day from tournament start to today (or end)
    today = min(date.today(), TOURNAMENT_END)
    fetch_date = TOURNAMENT_START
    updated = 0
    skipped = 0

    while fetch_date <= today:
        scores = fetch_scores_for_date(fetch_date)
        if scores:
            print(f"  {fetch_date}: {len(scores)} completed match(es)")
        for s in scores:
            key = (s["home"], s["away"])
            if key not in fixture_index:
                print(f"    !! No row found for: {s['home']} vs {s['away']}")
                skipped += 1
                continue
            row = fixture_index[key]
            existing_d = ws.cell(row, 4).value
            existing_e = ws.cell(row, 5).value
            if existing_d == s["home_score"] and existing_e == s["away_score"]:
                continue  # already up to date
            ws.cell(row, 4).value = s["home_score"]
            ws.cell(row, 5).value = s["away_score"]
            print(f"    Updated row {row}: {s['home']} {s['home_score']}-{s['away_score']} {s['away']}")
            updated += 1

        fetch_date += timedelta(days=1)

    if updated > 0:
        wb.save(XLSX_PATH)
        print(f"\nSaved. {updated} score(s) written.")
    else:
        print(f"\nNo changes needed. ({skipped} unmatched)")


if __name__ == "__main__":
    main()
