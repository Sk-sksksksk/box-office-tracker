#!/usr/bin/env python3
"""
Automated Daily Box Office Updater
----------------------------------
Usage:
    python auto_update.py

What it does:
1. Connects to trade sources (Sacnilk latest box office tracker).
2. Extracts daily Net India figures and active films.
3. Automatically adds new films to the 'Films' tab in box_office_tracker.xlsx if not present.
4. Appends new daily entries to the 'Daily Log' tab (avoiding duplicates).
5. Automatically runs build_site.py to refresh site/index.html and site/data.json.
"""
import datetime
import json
import re
import sys
import urllib.request
from pathlib import Path

import openpyxl

import build_site

USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
SACNILK_URL = "https://www.sacnilk.com/latest-boxoffice-collection"


def fetch_html(url):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.read().decode("utf-8", errors="ignore")
    except Exception as e:
        print(f"Warning: Could not fetch from {url}: {e}")
        return None


def parse_trade_figures(html):
    """
    Parses table rows from trade tracker page.
    Returns list of dicts: {'film': str, 'lang': str, 'rel': str, 'date': str, 'gross': float, 'source': str}
    """
    if not html:
        return []
    results = []
    today_str = datetime.date.today().isoformat()
    # Simple regex parser for table rows containing film links and numbers
    row_pattern = re.compile(r"<tr[^>]*>(.*?)</tr>", re.DOTALL | re.IGNORECASE)
    cell_pattern = re.compile(r"<td[^>]*>(.*?)</td>", re.DOTALL | re.IGNORECASE)

    for row_match in row_pattern.finditer(html):
        cells = [re.sub(r"<[^>]+>", "", c).strip() for c in cell_pattern.findall(row_match.group(1))]
        if len(cells) >= 4:
            # Example heuristic: film title in cell 0/1, language in cell 1/2, collection in cr
            title = cells[0] if not cells[0].isdigit() else cells[1]
            # Match number in ₹ cr
            cr_match = re.search(r"(\d+(?:\.\d+)?)\s*(?:cr|crore)?", " ".join(cells), re.IGNORECASE)
            if title and cr_match and len(title) > 2 and "total" not in title.lower():
                try:
                    val = float(cr_match.group(1))
                    if 0.05 <= val <= 250:
                        results.append({
                            "film": title,
                            "lang": "Hindi" if "hindi" in title.lower() else "Other",
                            "rel": today_str,
                            "date": today_str,
                            "gross": val,
                            "source": "Sacnilk"
                        })
                except ValueError:
                    continue
    return results


def update_workbook(wb_path, new_entries):
    wb = openpyxl.load_workbook(wb_path)
    ws_films = wb["Films"]
    ws_log = wb["Daily Log"]

    # 1. Existing films
    existing_films = set()
    first_empty_film_row = None
    for r in range(2, 101):
        val = ws_films.cell(r, 1).value
        if val:
            existing_films.add(str(val).strip())
        elif first_empty_film_row is None:
            first_empty_film_row = r

    # 2. Existing log keys (film, date, source, basis)
    existing_log = set()
    first_empty_log_row = None
    for r in range(2, 501):
        f = ws_log.cell(r, 2).value
        d = ws_log.cell(r, 3).value
        src = ws_log.cell(r, 7).value
        basis = ws_log.cell(r, 6).value
        if f and d and src:
            d_str = d.strftime("%Y-%m-%d") if isinstance(d, (datetime.datetime, datetime.date)) else str(d)[:10]
            existing_log.add((str(f).strip(), d_str, str(src).strip(), str(basis or "").strip()))
        elif first_empty_log_row is None:
            first_empty_log_row = r

    # 3. Add new films
    added_films = 0
    for e in new_entries:
        film = e["film"].strip()
        if film not in existing_films and first_empty_film_row and first_empty_film_row <= 100:
            ws_films.cell(first_empty_film_row, 1, film)
            ws_films.cell(first_empty_film_row, 2, e.get("lang", "Other"))
            ws_films.cell(first_empty_film_row, 3, datetime.datetime.strptime(e["rel"], "%Y-%m-%d"))
            ws_films.cell(first_empty_film_row, 4, "Auto-discovered")
            ws_films.cell(first_empty_film_row, 5, "Added by auto_update.py")
            existing_films.add(film)
            first_empty_film_row += 1
            added_films += 1

    # 4. Add new daily logs
    added_logs = 0
    now = datetime.datetime.now()
    for e in new_entries:
        film = e["film"].strip()
        d_str = e["date"]
        src = e["source"].strip()
        key = (film, d_str, src, "Net India")
        if key not in existing_log and first_empty_log_row and first_empty_log_row <= 500:
            ws_log.cell(first_empty_log_row, 1, now)
            ws_log.cell(first_empty_log_row, 2, film)
            ws_log.cell(first_empty_log_row, 3, datetime.datetime.strptime(d_str, "%Y-%m-%d"))
            ws_log.cell(first_empty_log_row, 5, float(e["gross"]))
            ws_log.cell(first_empty_log_row, 6, "Net India")
            ws_log.cell(first_empty_log_row, 7, src)
            ws_log.cell(first_empty_log_row, 8, "Trade estimate")
            ws_log.cell(first_empty_log_row, 9, "Auto-updated")
            existing_log.add(key)
            first_empty_log_row += 1
            added_logs += 1

    wb.save(wb_path)
    print(f"Workbook updated: {added_films} new film(s) added, {added_logs} new daily figure(s) logged.")
    return added_films, added_logs


def main():
    wb_file = Path("box_office_tracker.xlsx")
    if not wb_file.exists():
        sys.exit("Error: box_office_tracker.xlsx not found in current folder.")

    print("Fetching latest Indian box office updates...")
    html = fetch_html(SACNILK_URL)
    if html:
    print(f"DEBUG: fetched {len(html)} chars")
    m = re.search(r"<title>(.*?)</title>", html, re.I | re.S)
    print("DEBUG: title =", m.group(1).strip() if m else None)
    print("DEBUG: <tr> count =", len(re.findall(r"<tr", html, re.I)))
    print("DEBUG: first 600 chars:\n", html[:600])
    entries = parse_trade_figures(html)
    if not entries:
        print("Note: Direct web scraper received 0 new rows or could not reach server. Checking existing data.")
    else:
        print(f"Discovered {len(entries)} candidate figures online.")
        update_workbook(wb_file, entries)

    print("Rebuilding website leaderboard...")
    build_site.main()
    print("Done! Open site/index.html to view your updated Weekly Top 10.")


if __name__ == "__main__":
    main()
