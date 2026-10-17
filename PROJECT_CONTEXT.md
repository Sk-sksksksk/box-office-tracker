# Project Context & AI Handover Document: Indian Box Office Tracker

> **Instructions for Any New AI**:
> Read this document first before generating code or modifying files. This project is a working **Indian Box Office Tracker** that combines an Excel workbook database, an automated Python site generator, and an interactive single-page web app. Follow all file boundaries, Excel formula protections, and naming conventions documented below.

---

## 1. Project Overview & Core Goals
* **Purpose**: Track box office collections of Indian theatrical films across multiple independent trade websites alongside official producer claims, calculating an automated **Consistency / Variance Rating** to quantify trade consensus and identify inflated claims.
* **Current Focus**: Tracking the **Top 10 highest-collecting Indian films in the current week** (rolling 7-day window from the latest logged date). The leaderboard automatically recalculates weekly collections and updates weekly.
* **Tracked Sources**:
  * **Trade Websites**: `Sacnilk`, `Box Office India`, `Bollywood Hungama`, `Ormax Media` (Figure Type: `Trade estimate`).
  * **Producers**: `Producer / Official poster` (Figure Type: `Producer-reported`).
* **Reporting Standard**: All tracking compares **Net India** collections in **₹ Crore** (1 Cr = 10,000,000 INR).

---

## 2. Project Directory Structure
```text
BOX_OFFICE_TRACKER/
├── PROJECT_CONTEXT.md        # This AI handover and system state document
├── box_office_tracker.xlsx   # Main Excel database & automated formulas
├── box_office_log.py         # Interactive CLI logging tool (appends to Daily Log)
├── build_site.py             # Build script (reads Excel -> site/index.html + data.json)
├── site_template.html        # Interactive HTML5/JS app template (placeholder: const INJECTED=null;)
└── site/
    ├── index.html            # Generated standalone website (open directly in any browser)
    └── data.json             # Exported JSON data of all films and figures
```

---

## 3. Architecture & Working Principles

### A. Workbook Structure (`box_office_tracker.xlsx`)
The workbook has 6 sheets:
1. **`Films`**: Film metadata catalog (Rows 2–100).
   * Columns: `A` (Film), `B` (Language), `C` (Release Date), `D` (Director / Lead), `E` (Notes).
2. **`Daily Log`**: Daily box office entries (Rows 2–500).
   * **Input Columns (written by CLI/user)**:
     * `A` (`Logged At` timestamp)
     * `B` (`Film` title, matching `Films!A:A`)
     * `C` (`Business Date`)
     * `E` (`Gross (Rs cr)`)
     * `F` (`Gross Basis`, must be `"Net India"`)
     * `G` (`Source`, e.g. `Sacnilk`, `Box Office India`, `Bollywood Hungama`, `Producer / Official poster`)
     * `H` (`Figure Type`: `"Trade estimate"` or `"Producer-reported"`)
     * `I` (`Notes`)
   * **Formula Columns (DO NOT OVERWRITE OR SHIFT)**:
     * `D` (`Day No`): `=IF(OR(B2="",C2=""),"",IFERROR(C2-INDEX(Films!$C$2:$C$100,MATCH(B2,Films!$A$2:$A$100,0))+1,""))`
     * `J` (`Trade Sources count`): `=IF(B2="","",COUNTIFS($B$2:$B$500,B2,$C$2:$C$500,C2,$F$2:$F$500,F2,$H$2:$H$500,"Trade estimate"))`
     * `K` (`Trade Max`): `MAXIFS` over Trade estimates for same Film, Date, and Net India basis.
     * `L` (`Trade Min`): `MINIFS` over Trade estimates for same Film, Date, and Net India basis.
     * `M` (`Spread %`): `=IF(OR(J2<2,L2=0),"",(K2-L2)/L2)`
     * `N` (`Flag`): `=IF(B2="","",IF(J2<2,"Needs 2+ trade sources",IF(M2>Lists!$E$2,"CHECK","OK")))`
     * `O` (`Consistency Rating`): `=IF(B2="","",IF(J2<2,"Needs 2+ trade sources",IF(M2<=Lists!$E$6,"Good",IF(M2<=Lists!$E$2,"Moderate","Bad"))))`
3. **`Lists`**:
   * `E2`: `0.10` (10% Bad/Disagreement threshold).
   * `E6`: `0.05` (5% Good consistency threshold).
4. **`Summary`**: Aggregates film metrics across all logged days.
   * `A`: Film name (`=IF(Films!A2="","",Films!A2)`)
   * `B`: Language (`=IF(A2="","",Films!B2)`)
   * `C`: Release Date (`=IF(A2="","",Films!C2)`)
   * `D`: Days Logged (`MAXIFS` on Day No)
   * `E`: Day 1 / Latest Trade Avg (`AVERAGEIFS`)
   * `F`: Day 1 / Latest Producer (`AVERAGEIFS`)
   * `G`: Producer Premium % (`F/E - 1`)
   * `H`: Latest Day Trade Avg
   * `I`: Avg Trade Spread % (`AVERAGEIFS` on Col M where Trade count $\ge 2$)
   * `J`: Max Trade Spread % (`MAXIFS` on Col M)
   * `K`: Consistency Rating (`Good` $\le 5\%$, `Moderate` $5\%\text{–}10\%$, `Bad` $> 10\%$)
   * `L`: Producer Alignment (`Good (Aligned)`, `Moderate Gap`, `Bad (Inflated)`)
5. **`Sources`**: List of valid tracking websites and entities.
6. **`README`**: Quick instructions.

### B. Automated Build Pipeline (`build_site.py`)
* Reads `Films` and `Daily Log` from `box_office_tracker.xlsx`.
* Filters for `Gross Basis == "Net India"` and excludes rows with `Confirmed actual`.
* Formats figures into JSON and injects them into `site_template.html` at `const INJECTED=null;`.
* Writes the output to `site/index.html` and `site/data.json`.
* **Golden Rule**: Never edit `site/index.html` directly—always edit `site_template.html` and re-run `python build_site.py`.

### C. Web Frontend (`site_template.html` / `site/index.html`)
* Standalone single-file HTML/CSS/JS (zero external CDN or internet dependencies).
* Computes a rolling 7-day window from the latest logged date (e.g. Oct 10–16, 2026).
* Automatically ranks the **Top 10 highest-collecting movies** of the week by Trade Average.
* Features:
  * Dual threshold sliders (`Good` default 5%, `Bad/Flag` default 10%).
  * 0–100 Consistency Score pill.
  * Rating mode toggle: Website Spread Only vs Combined (Trade + Producer).
  * Scope toggle: Current Week Top 10 vs All Logged Films.
  * Interactive SVG chart (Daily collections with error bars vs Cumulative run).
  * In-page `+ Log Weekly / Daily Figure` form + CSV export.
  * Dark / light mode toggle.

---

## 4. Rating Rubric & Standards
| Rating | Trade Spread % | Meaning |
| :--- | :--- | :--- |
| **Good** | $\le 5\%$ | Strong trade consensus. All websites report near-identical figures. |
| **Moderate** | $5\% < \text{Spread} \le 10\%$ | Minor differences between website tracking panels. Normal variance. |
| **Bad** | $> 10\%$ | Substantial trade disagreement. Needs manual inspection or primary source verification. |
| **Needs 2+ sources** | $< 2$ trade sites | Only 1 trade website has logged for this date. Cannot compute variance. |

---

## 5. Current Data Snapshot (Week of Oct 10–16, 2026)
1. **Drishyam 3: The Conclusion** (Hindi, Rel: Oct 2, 2026) — Week 2: ~₹87.0 cr Trade Avg, Consistency: `Good` (Avg Spread: 3.4%).
2. **Jailer 2** (Tamil, Rel: Oct 15, 2026) — 2-Day opening: ₹57.2 cr Trade Avg, Consistency: `Moderate` (Avg Spread: 7.0%). Producer claimed ₹70.5 cr (+22.2%).
3. **Comrade Kalyan** (Telugu, Rel: Oct 9, 2026) — Week 1: ₹11.0 cr Trade Avg, Consistency: `Good` (Avg Spread: 4.9%).
4. **Baththa** (Tamil, Rel: Oct 1, 2026) — Week 2: ₹8.5 cr Trade Avg, Consistency: `Moderate` (Avg Spread: 6.0%).
5. **Ranabaali** (Telugu, Rel: Oct 16, 2026) — Day 1: ₹6.6 cr Trade Avg, Consistency: `Bad` (Spread: 14.2%). Producer: ₹9.1 cr (+37.9%).
6. **Vvan - Force of the Forrest** (Hindi, Rel: Sep 25, 2026) — Week 3: ₹5.7 cr Trade Avg, Consistency: `Moderate` (Avg Spread: 6.4%).
7. **Other Mommy** (English/Other, Rel: Oct 9, 2026) — Week 1: ₹3.3 cr Trade Avg, Consistency: `Good` (Avg Spread: 4.6%).
8. **Prahaar - The Ujjwal Nikam Story** (Hindi, Rel: Oct 16, 2026) — Day 1: ₹2.27 cr Trade Avg, Consistency: `Bad` (Spread: 16.7%).
9. **Street Fighter (Tamil)** (Tamil, Rel: Oct 16, 2026) — Day 1: ₹0.91 cr Trade Avg, Consistency: `Moderate` (Spread: 6.8%).
10. **The Eken: Kerala-e-Kurukshetra** (Bengali, Rel: Oct 15, 2026) — 2-Day: ₹0.75 cr Trade Avg, Consistency: `Moderate` (Avg Spread: 6.9%).

---

## 6. Common Developer Workflows
* **Adding figures manually**: Open `box_office_tracker.xlsx` -> Go to `Daily Log` -> Append rows in columns A, B, C, E, F, G, H, I. Do not touch grey formula columns.
* **Adding figures via CLI**: Run `python box_office_log.py` and follow prompts.
* **Updating the website**: Run `python build_site.py`.
* **Modifying the UI or Logic**: Edit `site_template.html` -> Run `python build_site.py` -> Open `site/index.html`.
