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
- `dashboard/` — dashboard exports/screenshots
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
