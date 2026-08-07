"""
Turn the cleaned/categorized survey data into decision-ready insights.

Reads data/processed/parent_survey_cleaned.csv and computes four
independent insights, each isolated in its own function and its own
labeled columns/rows in the output so they're easy to tell apart:

  1. Controllability   - is this "why not" reason something HXP can
                          act on, or a family's own choice?
  2. Soft No Signal     - does the response hint at future interest
                          ("maybe next year") rather than a firm no?
  3. Reason Stack Count - how many distinct themes are packed into one
                          response (a rough "how hard to move" proxy)?
  4. Would-Help Themes  - what do the 42 "Maybe" respondents (the
                          survey's built-in persuadable segment) say
                          would change their mind?

Outputs:
  - data/processed/parent_survey_insights.csv  (row-level, enriched)
  - data/processed/insights_summary.csv        (headline numbers,
    grouped by an "Insight" column so each of the four is easy to
    isolate/filter in Sheets)

Usage: python3 scripts/derive_insights.py
"""

import csv
import re
import sys
from pathlib import Path
from collections import Counter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from transform_data import categorize  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
CLEANED_CSV = ROOT / "data" / "processed" / "parent_survey_cleaned.csv"
INSIGHTS_CSV = ROOT / "data" / "processed" / "parent_survey_insights.csv"
SUMMARY_CSV = ROOT / "data" / "processed" / "insights_summary.csv"


# ============================================================
# INSIGHT 1: Controllability
# Maps each "Reason Not To Go" category to how much HXP can do about
# it. This is the judgment call the whole insight rests on - revisit
# it if the Parent Builder team disagrees with a placement.
# ============================================================
CONTROLLABILITY = {
    "Not Aware It Was an Option": "HXP Can Directly Fix",
    "Trip / Spot Was Already Full": "HXP Can Directly Fix",
    "Cost / Finances": "HXP Can Directly Fix",
    "Work Schedule / Time Off": "HXP Can Influence",
    "Not a Traveler / Personal Preference": "HXP Can Influence",
    "Health / Physical / Age": "HXP Can Influence",
    "Wants Child to Have Independent Experience": "Family's Own Choice",
    "Child Didn't Want Parent to Come": "Family's Own Choice",
    "Other Kids / Family at Home": "Family's Own Choice",
    "Another Family Member Going Instead": "Family's Own Choice",
    "Other Life Commitment": "Family's Own Choice",
    "Not Needed": "Family's Own Choice",
    "Family's Own Choice (No Elaboration)": "Family's Own Choice",
    "Other / Uncategorized": "Unknown",
}


def add_controllability(row):
    row["Not Going - Controllability"] = CONTROLLABILITY.get(row["Reason Not To Go - Category"], "")


# ============================================================
# INSIGHT 2: Soft No Signal
# Flags responses that hint the "No" isn't final - future interest,
# timing-dependent, "I would love to but...".
# ============================================================
SOFT_NO_PATTERN = re.compile(
    r"next year|maybe next|would love to|might go|another year|in the future|"
    r"would consider|for future|someday|not this (year|summer)|too late|"
    r"next time|would like to (go|consider)",
    re.IGNORECASE,
)


def add_soft_no_signal(row):
    text = row["Reason Not To Go"]
    row["Not Going - Soft No Signal"] = bool(text and SOFT_NO_PATTERN.search(text))


# ============================================================
# INSIGHT 3: Reason Stack Count
# Counts how many distinct themes are packed into one "why not"
# response, using the multi-tag list transform_data.py already built.
# More stacked reasons = plausibly a harder parent to move with any
# single fix.
# ============================================================
def add_reason_stack_count(row):
    tags_field = row["Reason Not To Go - All Tags"]
    row["Not Going - Reason Stack Count"] = (
        len([t for t in tags_field.split(";") if t.strip()]) if tags_field.strip() else 0
    )


# ============================================================
# INSIGHT 4: Would-Help Themes
# Categorizes the 33 "what would help" answers, all of which come
# from "Maybe" respondents - the survey's built-in persuadable
# segment (42 people total).
# ============================================================
WOULD_HELP_CATEGORIES = [
    ("Cost / Financial Assistance", r"\bcost\b|\bmoney\b|\bfund|\bfinanc|\bexpens\w*|cheap|\bprice\b"),
    ("Time / Scheduling", r"\btime\b|\bschedule|\bdates?\b|\bbusy\b|advance notice|free time"),
    ("Childcare for Other Kids", r"child ?care|younger kids|other (kid|child)|left unattended|unattended|care for (my |the )?(kids|children)|leave (my |the )?(kids|children) at home"),
    ("Registration / Spot Availability", r"easier to get a spot|mess up|process|\bspots?\b|registration"),
    ("More Information / Awareness", r"more info|don'?t know much|need more info|not sure what"),
    ("Waiting for Right Timing (kids' ages)", r"younger|older|age range|next kid|sibling|grow"),
    ("Health", r"\bhealth\b"),
    # A distinct signal from "no answer": the parent engaged with the question
    # but doesn't have a concrete ask yet - different from silence, worth
    # keeping separate from "Other/Uncategorized".
    ("No Specific Ask / Undecided", r"not sure|don'?t know\b|^n/?a$"),
]


def add_would_help_category(row):
    primary, tags = categorize(row["What Would Help"], WOULD_HELP_CATEGORIES)
    row["Would Help - Category"] = primary
    row["Would Help - All Tags"] = tags


def main():
    with open(CLEANED_CSV, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    for row in rows:
        add_controllability(row)
        add_soft_no_signal(row)
        add_reason_stack_count(row)
        add_would_help_category(row)

    fieldnames = list(rows[0].keys())
    with open(INSIGHTS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows)} enriched rows to {INSIGHTS_CSV}")

    summary_rows = []
    no_with_reason = [r for r in rows if r["Status"] == "No" and r["Reason Not To Go"]]
    total_no_reasons = len(no_with_reason)

    # --- Insight 1 summary: Controllability ---
    controllability_counts = Counter(r["Not Going - Controllability"] for r in no_with_reason)
    for label in ["HXP Can Directly Fix", "HXP Can Influence", "Family's Own Choice", "Unknown"]:
        count = controllability_counts.get(label, 0)
        pct = round(100 * count / total_no_reasons, 1)
        summary_rows.append(("1. Controllability", label, count))
        summary_rows.append(("1. Controllability", f"{label} (% of reasons given)", pct))

    # --- Insight 2 summary: Soft No Signal ---
    soft_no_count = sum(1 for r in no_with_reason if r["Not Going - Soft No Signal"])
    summary_rows.append(("2. Soft No Signal", "Responses hinting at future interest", soft_no_count))
    summary_rows.append(("2. Soft No Signal", "% of 'No' reasons given", round(100 * soft_no_count / total_no_reasons, 1)))

    # --- Insight 3 summary: Reason Stack Count ---
    stack_counts = Counter(r["Not Going - Reason Stack Count"] for r in no_with_reason)
    multi_blocker_count = sum(v for k, v in stack_counts.items() if k > 1)
    summary_rows.append(("3. Reason Stack Count", "Single-theme responses", stack_counts.get(1, 0)))
    summary_rows.append(("3. Reason Stack Count", "Multi-theme responses (2+ stacked reasons)", multi_blocker_count))
    summary_rows.append(("3. Reason Stack Count", "Multi-theme (% of 'No' reasons given)", round(100 * multi_blocker_count / total_no_reasons, 1)))

    # --- Insight 4 summary: Would-Help Themes ---
    maybe_count = sum(1 for r in rows if r["Status"] == "Maybe")
    summary_rows.append(("4. Would-Help Themes", "Total 'Maybe' respondents (persuadable segment)", maybe_count))
    would_help_counts = Counter(r["Would Help - Category"] for r in rows if r["Would Help - Category"])
    for label, count in would_help_counts.most_common():
        summary_rows.append(("4. Would-Help Themes", label, count))

    with open(SUMMARY_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Insight", "Metric", "Value"])
        writer.writerows(summary_rows)
    print(f"Wrote {len(summary_rows)} summary rows to {SUMMARY_CSV}")

    print("\n--- Headline numbers, grouped by insight ---")
    current_insight = None
    for insight, label, value in summary_rows:
        if insight != current_insight:
            print(f"\n{insight}")
            current_insight = insight
        print(f"  {label}: {value}")


if __name__ == "__main__":
    main()
