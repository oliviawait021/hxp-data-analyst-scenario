# HXP Data Analyst Scenario — Parent Builder Survey

Analysis of HXP's "Parent at Home Survey" to understand why parents do or don't sign up as Parent Builders on trips with their kids, and to build a dashboard managers can use to spot trends at a glance.

## Background

Every summer before trips start, HXP surveys builders and their parents (both those going on trips and those not) about their pre-trip experience. Finding enough quality Parent Builders is a persistent growth bottleneck, so this project mines the survey responses to understand:

- What motivates parents who do sign up
- What holds back parents who don't
- What (if anything) would change a "no" into a "yes"

## Contents

- `Data Analyst Scenario.pdf` — the original assignment prompt
- `Data Team Scenario.xlsx` — working copy of the raw survey response data
- `data/` — cleaned/transformed data used for analysis
- `dashboard/` — the assembled multi-tab dashboard workbook
- `SUMMARY.md` — process write-up and findings/recommendations (added once analysis is complete)

## Source data

Raw responses: [Google Sheet](https://docs.google.com/spreadsheets/d/1PhClVzrInmpbLjNGCx37O11qmA0evbSgcuimX1XQeTQ/edit?usp=sharing) (partial export, 1,315 rows including 126 empty rows, 4 columns: planning-to-go status, reason to go, reason not to go, what would change the decision).

## Data transformation

`scripts/transform_data.py` reads `Data Team Scenario.xlsx`, drops empty rows, normalizes text, and tags each free-text response into a category (e.g. Cost / Finances, Work Schedule / Time Off, Wants Child to Have Independent Experience) using keyword matching. Output goes to `data/processed/parent_survey_cleaned.csv`, which is what feeds the dashboard's pivot tables/charts.

```
pip install -r requirements.txt
python3 scripts/transform_data.py
```

Categories are a first pass via regex, not a substitute for reading the responses — the script prints an "Other / Uncategorized" count on each run so you can spot-check how much is still unclassified and refine the keyword lists as needed.

`scripts/qa_check.py` verifies `parent_survey_cleaned.csv` is a faithful transformation of the source (no dropped/corrupted rows, no text without a matching category or vice versa). Run it after any change to `transform_data.py`.

## Insights

`scripts/derive_insights.py` reads the cleaned CSV and adds four independent, separately-labeled insights on top of the categories:

1. **Controllability** — is this "why not" reason something HXP can directly fix (cost, awareness, capacity), something HXP can influence (scheduling, personal hesitation), or a family's own choice? Answers "where should the Parent Builder team actually spend effort."
2. **Soft No signal** — flags responses hinting the "no" isn't final (*"would love to go next year"*) — a previously-invisible list of parents worth re-engaging.
3. **Reason stack count** — how many distinct themes are packed into one response; multi-reason responses are plausibly harder to move with a single fix.
4. **Would-help themes** — categorizes what the 42 "Maybe" respondents said would change their mind (the survey only asks this question of Maybes, never of firm Nos).

Outputs:
- `data/processed/parent_survey_insights.csv` — the cleaned data plus one column per insight, row-level.
- `data/processed/insights_summary.csv` — headline counts/percentages, grouped by an `Insight` column so each of the four stays easy to isolate when pulled into the dashboard.

```
python3 scripts/derive_insights.py
```

## Dashboard

`scripts/build_dashboard_workbook.py` assembles everything into one workbook: `dashboard/HXP Parent Builder Dashboard.xlsx`. Six tabs:

1. **Dashboard** — KPI cards, 4 charts, and a handful of real "warm lead" quotes. Opens first.
2. **Chart Data** — small COUNTIF-driven tables the charts read from (visible but not meant for managers to read directly).
3. **Raw Data** — untouched copy of the original survey export.
4. **Cleaned + Categorized** — output of `transform_data.py`.
5. **Insights** — output of `derive_insights.py`.
6. **Insights Summary** — headline metrics from `derive_insights.py`.

KPI cards and chart data are live formulas against the Insights sheet (not hardcoded numbers), so the dashboard updates if a category is hand-edited later in Sheets.

```
python3 scripts/build_dashboard_workbook.py
```

To use it: open the file directly, or upload it to Google Drive and open with Google Sheets — it imports as a normal multi-tab spreadsheet and can replace/merge into the "make a copy of this document" working copy from the assignment.

Formula correctness was checked with the `formulas` Python library (an Excel formula evaluator) rather than by eye — it caught one real bug during development: a `COUNTIFS(...,"<>")` "not-blank" idiom that didn't evaluate as expected, replaced with a formula that sums the already-verified category counts instead. Chart layout/rendering itself hasn't been visually confirmed since no spreadsheet app was available in this environment — worth a visual pass after opening in Sheets.
