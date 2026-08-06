"""
Clean and categorize the Parent at Home Survey responses.

Reads the raw survey export (Data Team Scenario.xlsx), drops empty rows,
and tags the free-text "why did/didn't you go" answers into themes so they
can be counted and charted in the dashboard.

Usage: python3 scripts/transform_data.py
"""

import csv
import re
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
SOURCE_XLSX = ROOT / "Data Team Scenario.xlsx"
OUTPUT_CSV = ROOT / "data" / "processed" / "parent_survey_cleaned.csv"

STATUS_COL = "Are you planning on going as a Parent Builder?"
GOING_COL = "What influenced your decision to go?"
NOT_GOING_COL = "What influenced your decision not to go?"
WOULD_HELP_COL = "What would help change your decision to go?"

# Order matters: a response can match several categories, but the first
# match in this list becomes its "primary" category for simple bar charts.
# Boundaries are approximate (e.g. "wants independence" vs "child asked
# parent not to come" overlap in wording) - spot-check before trusting fully.
NOT_GOING_CATEGORIES = [
    ("Cost / Finances", r"\bcost\b|\bmoney\b|\bprice\b|afford|\bbudget\b|expensive|\bexpense\b|financ"),
    ("Work Schedule / Time Off", r"\bwork\b|\bjob\b|\bschedule\b|\bbusy\b|obligation|\bpto\b|time off|vacation|\btim(e|ing)\b|availab|responsibilit|commitment"),
    ("Other Kids / Family at Home", r"other (kid|child)|younger|sibling|kids? at home|children at home|family at home|other kiddos|newborn|\bbaby\b|young ones|lot of kids|many kids"),
    ("Another Family Member Going Instead", r"(husband|wife|dad|father|mom|mother|uncle|aunt|grandpa|grandma|grandmother|grandfather) is going|another (parent|family member)"),
    ("Child Didn't Want Parent to Come", r"did ?n'?t want me|does ?n'?t want me|asked (me )?not|want(ed)? (him|her) to go alone|want(ed)? (his|her) own (space|experience)"),
    ("Wants Child to Have Independent Experience", r"independen|own experience|without (a |his |her |my )?(me|mom|dad|us|parent)|on (his|her) own|\bgrow\b|confidence|rely on (his|her)|\balone\b|\bsolo\b|autonomy|\bexperience\b"),
    ("Health / Physical / Age", r"\bhealth\b|physical|\bknee\b|bad back|back (pain|injury|problem)|\bage\b|\btoo old\b|medical|injur|illness|disab"),
    ("Other Life Commitment", r"\bmission\b|wedding|surgery|pregnan|\bmov(e|ing)\b|deploy|trip leader|another (trip|program)|graduat"),
    ("Didn't Understand the Program", r"did ?n'?t (really )?understand|not sure what|unclear|confus"),
    ("Not Needed", r"not needed|not necessary|no need|did ?n'?t need"),
]

GOING_CATEGORIES = [
    ("Wanted Shared Experience with Child", r"experience|memories|bond|together|time with|spend time"),
    ("Child/Spouse Asked Them to Go", r"asked (me|him|her)|invited"),
    ("Love of Travel / Culture", r"travel|culture|adventure|explore"),
    ("Previous Positive Experience", r"previous|been before|last year|again"),
    ("Faith / Service Motivation", r"\bspirit\b|\bfaith\b|\bserve\b|service|god|mission"),
    ("Trust in the HXP Program", r"trust|impressed by hxp|good hands|love hxp|hxp (program|offers)"),
]


def clean_text(value):
    if value is None:
        return ""
    text = str(value).replace("’", "'").replace("‘", "'")
    return re.sub(r"\s+", " ", text).strip()


def categorize(text, categories):
    text_lower = text.lower()
    matches = [label for label, pattern in categories if re.search(pattern, text_lower)]
    primary = matches[0] if matches else ("Other / Uncategorized" if text else "")
    return primary, "; ".join(matches)


def main():
    wb = openpyxl.load_workbook(SOURCE_XLSX, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(min_row=2, values_only=True))
    header = [STATUS_COL, GOING_COL, NOT_GOING_COL, WOULD_HELP_COL]

    cleaned_rows = []
    for raw_row in rows:
        status, going, not_going, would_help = (clean_text(v) for v in raw_row[:4])

        # Drop fully empty rows (no status and no text anywhere)
        if not any([status, going, not_going, would_help]):
            continue

        going_primary, going_tags = categorize(going, GOING_CATEGORIES)
        not_going_primary, not_going_tags = categorize(not_going, NOT_GOING_CATEGORIES)

        cleaned_rows.append({
            "Status": status,
            "Reason To Go": going,
            "Reason To Go - Category": going_primary,
            "Reason To Go - All Tags": going_tags,
            "Reason Not To Go": not_going,
            "Reason Not To Go - Category": not_going_primary,
            "Reason Not To Go - All Tags": not_going_tags,
            "What Would Help": would_help,
        })

    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(cleaned_rows[0].keys()))
        writer.writeheader()
        writer.writerows(cleaned_rows)

    print(f"Wrote {len(cleaned_rows)} cleaned rows to {OUTPUT_CSV}")

    # Quick sanity-check summary
    from collections import Counter
    status_counts = Counter(r["Status"] for r in cleaned_rows)
    print("\nStatus breakdown:")
    for status, count in status_counts.most_common():
        print(f"  {status or '(blank)'}: {count}")

    not_going_counts = Counter(r["Reason Not To Go - Category"] for r in cleaned_rows if r["Reason Not To Go - Category"])
    print("\nTop reasons NOT to go:")
    for cat, count in not_going_counts.most_common():
        print(f"  {cat}: {count}")

    going_counts = Counter(r["Reason To Go - Category"] for r in cleaned_rows if r["Reason To Go - Category"])
    print("\nTop reasons TO go:")
    for cat, count in going_counts.most_common():
        print(f"  {cat}: {count}")


if __name__ == "__main__":
    main()
