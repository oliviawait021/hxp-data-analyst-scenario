"""
QA check: verify data/processed/parent_survey_insights.csv is a faithful,
correctly-computed enrichment of parent_survey_cleaned.csv.

Rebuilds the four insight columns straight from the cleaned CSV using
the same functions as derive_insights.py, diffs against the committed
output, and checks for logic gaps that a plain diff wouldn't catch -
e.g. a category that exists in the data but has no controllability
mapping (silently falls back to blank instead of erroring).

Usage: python3 scripts/qa_check_insights.py
"""

import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from derive_insights import (  # noqa: E402
    CLEANED_CSV, INSIGHTS_CSV, CONTROLLABILITY,
    add_controllability, add_soft_no_signal, add_reason_stack_count, add_would_help_category,
)


def build_expected_rows():
    with open(CLEANED_CSV, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    for row in rows:
        add_controllability(row)
        add_soft_no_signal(row)
        add_reason_stack_count(row)
        add_would_help_category(row)
    return rows


def load_actual_rows():
    with open(INSIGHTS_CSV, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def main():
    expected = build_expected_rows()
    actual = load_actual_rows()
    failures = []

    if len(expected) != len(actual):
        failures.append(f"ROW COUNT MISMATCH: expected {len(expected)}, got {len(actual)}")

    mismatches = []
    for i, (exp, act) in enumerate(zip(expected, actual)):
        for key, exp_val in exp.items():
            exp_str = str(exp_val)
            act_val = act.get(key, "<MISSING COLUMN>")
            if exp_str != act_val:
                mismatches.append((i, key, exp_str, act_val))
    if mismatches:
        failures.append(f"{len(mismatches)} FIELD MISMATCHES (output doesn't match recomputed values - stale file?)")

    # Coverage gap: does every category actually present in the data have a
    # controllability mapping? A missing entry silently produces "" instead
    # of erroring, so this is the check that catches it.
    used_categories = {r["Reason Not To Go - Category"] for r in actual if r["Reason Not To Go - Category"]}
    unmapped = used_categories - set(CONTROLLABILITY.keys())
    if unmapped:
        failures.append(f"UNMAPPED CATEGORIES (no controllability assigned): {unmapped}")

    # Soft No should only ever be True when there's actual reason text
    bad_soft_no = [i for i, r in enumerate(actual)
                   if r["Not Going - Soft No Signal"] == "True" and not r["Reason Not To Go"].strip()]
    if bad_soft_no:
        failures.append(f"{len(bad_soft_no)} rows flagged Soft No with no reason text (rows: {bad_soft_no[:10]})")

    # Reason Stack Count should be 0 exactly when All Tags is empty, and vice versa
    bad_stack = [i for i, r in enumerate(actual)
                 if (r["Not Going - Reason Stack Count"] == "0") != (not r["Reason Not To Go - All Tags"].strip())]
    if bad_stack:
        failures.append(f"{len(bad_stack)} rows where Reason Stack Count doesn't match All Tags emptiness (rows: {bad_stack[:10]})")

    # Would Help Category should be non-empty iff What Would Help text is non-empty
    bad_would_help = [i for i, r in enumerate(actual)
                      if bool(r["What Would Help"].strip()) != bool(r["Would Help - Category"].strip())]
    if bad_would_help:
        failures.append(f"{len(bad_would_help)} rows where Would Help - Category doesn't match What Would Help emptiness (rows: {bad_would_help[:10]})")

    # Controllability should only ever be blank for rows with no not-going category,
    # and non-blank (one of 4 known labels) whenever a category exists
    known_controllability = set(CONTROLLABILITY.values())
    bad_controllability = [i for i, r in enumerate(actual)
                           if r["Not Going - Controllability"] and r["Not Going - Controllability"] not in known_controllability]
    if bad_controllability:
        failures.append(f"{len(bad_controllability)} rows with an unrecognized Controllability value (rows: {bad_controllability[:10]})")

    print(f"Expected rows (recomputed): {len(expected)}")
    print(f"Actual rows (in CSV):       {len(actual)}")
    print(f"Categories used in data:    {len(used_categories)}")
    print(f"Categories with a controllability mapping: {len(set(CONTROLLABILITY.keys()) & used_categories)}")

    if failures:
        print("\nFAIL:")
        for f in failures:
            print(f"  - {f}")
        if mismatches:
            print("\n  First 10 field mismatches:")
            for row_i, key, exp_val, act_val in mismatches[:10]:
                print(f"    row {row_i} [{key}]: expected={exp_val!r} actual={act_val!r}")
        sys.exit(1)
    else:
        print("\nPASS: insights CSV matches recomputed values, every category has a "
              "controllability mapping, and all four insight columns are internally consistent.")


if __name__ == "__main__":
    main()
