"""
QA check: verify data/processed/parent_survey_cleaned.csv is a faithful,
correctly-categorized transformation of the source Data Team Scenario.xlsx.

Rebuilds the expected output straight from the source workbook using the
same clean_text()/categorize() functions as transform_data.py, then diffs
it against the committed CSV row-by-row. Catches: dropped/duplicated rows,
corrupted text, stale categories (CSV not regenerated after a rule change),
and text/category inconsistencies.

Usage: python3 scripts/qa_check.py
"""

import csv
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from transform_data import (  # noqa: E402
    SOURCE_XLSX, OUTPUT_CSV, GOING_CATEGORIES, NOT_GOING_CATEGORIES,
    clean_text, categorize,
)


def build_expected_rows():
    wb = openpyxl.load_workbook(SOURCE_XLSX, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    expected = []
    for raw_row in rows:
        status, going, not_going, would_help = (clean_text(v) for v in raw_row[:4])
        if not any([status, going, not_going, would_help]):
            continue
        going_primary, going_tags = categorize(going, GOING_CATEGORIES)
        not_going_primary, not_going_tags = categorize(not_going, NOT_GOING_CATEGORIES)
        expected.append({
            "Status": status,
            "Reason To Go": going,
            "Reason To Go - Category": going_primary,
            "Reason To Go - All Tags": going_tags,
            "Reason Not To Go": not_going,
            "Reason Not To Go - Category": not_going_primary,
            "Reason Not To Go - All Tags": not_going_tags,
            "What Would Help": would_help,
        })
    return expected


def load_actual_rows():
    with open(OUTPUT_CSV, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def main():
    expected = build_expected_rows()
    actual = load_actual_rows()

    failures = []

    if len(expected) != len(actual):
        failures.append(
            f"ROW COUNT MISMATCH: source implies {len(expected)} rows, "
            f"CSV has {len(actual)} rows"
        )

    mismatched_rows = []
    for i, (exp, act) in enumerate(zip(expected, actual)):
        for key in exp:
            if exp[key] != act.get(key, "<MISSING COLUMN>"):
                mismatched_rows.append((i, key, exp[key], act.get(key)))

    if mismatched_rows:
        failures.append(f"{len(mismatched_rows)} FIELD MISMATCHES across {len(set(r[0] for r in mismatched_rows))} rows")

    # Text/category consistency: category should be non-empty iff text is non-empty
    consistency_issues = []
    for i, act in enumerate(actual):
        for text_col, cat_col in [
            ("Reason To Go", "Reason To Go - Category"),
            ("Reason Not To Go", "Reason Not To Go - Category"),
        ]:
            has_text = bool(act[text_col])
            has_cat = bool(act[cat_col])
            if has_text != has_cat:
                consistency_issues.append((i, text_col, act[text_col], act[cat_col]))

    if consistency_issues:
        failures.append(f"{len(consistency_issues)} TEXT/CATEGORY CONSISTENCY ISSUES (text present but no category, or vice versa)")

    # Duplicate rows (informational, not necessarily an error)
    seen = {}
    dup_count = 0
    for act in actual:
        key = tuple(act.values())
        seen[key] = seen.get(key, 0) + 1
    dup_count = sum(1 for v in seen.values() if v > 1)

    # Status value validity
    valid_statuses = {"Yes", "No", "Maybe"}
    bad_statuses = [act["Status"] for act in actual if act["Status"] not in valid_statuses]

    print(f"Expected rows (recomputed from source): {len(expected)}")
    print(f"Actual rows (in CSV):                   {len(actual)}")
    print(f"Exact-duplicate row groups in CSV:       {dup_count}")
    print(f"Rows with unexpected Status value:       {len(bad_statuses)}"
          + (f" -> {set(bad_statuses)}" if bad_statuses else ""))

    if failures:
        print("\nFAIL:")
        for f in failures:
            print(f"  - {f}")
        if mismatched_rows:
            print("\n  First 10 field mismatches:")
            for row_i, key, exp_val, act_val in mismatched_rows[:10]:
                print(f"    row {row_i} [{key}]: expected={exp_val!r} actual={act_val!r}")
        if consistency_issues:
            print("\n  First 10 consistency issues:")
            for row_i, col, text, cat in consistency_issues[:10]:
                print(f"    row {row_i} [{col}]: text={text!r} category={cat!r}")
        sys.exit(1)
    else:
        print("\nPASS: cleaned CSV matches source data exactly and every categorized "
              "row is internally consistent.")


if __name__ == "__main__":
    main()
