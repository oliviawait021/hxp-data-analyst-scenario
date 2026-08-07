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
    ("Trip / Spot Was Already Full", r"(was|is) (already )?full\b|already had (both|all|enough)|no (spots?|room|space) (left|available)|spots? (were|was) (taken|full)"),
    # A response that just names a family member and nothing else ("My child",
    # "my child") gives no obstacle to fix - it's a family decision by
    # definition, so it's mapped straight to that controllability bucket
    # rather than sitting in Other/Uncategorized.
    ("Family's Own Choice (No Elaboration)", r"^(my |our )?(child|son|daughter|kid)s?[\s:)\.,!]*$"),
    ("Cost / Finances", r"\bcost\b|\bmoney\b|\bprice\b|afford|\bbudget\b|expensive|\bexpense\b|financ|\bfunds?\b|\$|💵"),
    ("Work Schedule / Time Off", r"\bwork\b|\bjob\b|schedul|\bbusy\b|obligation|\bpto\b|time off|vacation|\btim(e|ing)\b|availab|responsibilit|commitment|booked (up )?with|too long (of a )?trip|can'?t make it|too much going on"),
    ("Other Kids / Family at Home", r"other (kid|child)|younger|sibling|kids? at home|children at home|family at home|other kiddos|newborn|\bbaby\b|young ones|lot of kids|many kids|\d+ kids\b|child ?care|can'?t leave|couldn'?t leave|leav(e|ing) (my |the )?family|away from (home|family)|for that long|(son|daughter|child|kid)s?[a-z '\"]{0,25}at home|little one at home|a little at home|when (my |our )?youngest (can|is|goes)"),
    ("Another Family Member Going Instead", r"(husband|wife|spouse|dad|father|mom|mother|uncle|aunt|grandma|grandpa|grandmother|grandfather|sister|brother|cousin)[a-z '\"-]{0,25}(is going|was going|went|gone|wanted to go|going (this|as)|is the parent builder|already)|another (parent|family member)|friends? (and|&) family (have|has)? ?(been|gone)|extended family (went|has gone)|previous (child|children) (have|has) gone|other (kids|children) (have|has) gone|call(ed)? \"?dibs\"?"),
    # Keyed on the CHILD's expressed wish ("she didn't want me to go"), not the parent's own
    # reasoning ("I wanted her to have it without me") - the latter belongs to the next category.
    ("Child Didn't Want Parent to Come", r"(did|does|would) ?(not|n'?t) (really )?want (me|us|a parent|the parent)|asked (me )?not|didn'?t ask me to (go|come)|prefers (i|we) (didn'?t|don'?t|not) go|said i couldn'?t|told me i couldn'?t|want(ed)? (him|her) to go alone|want(ed)? (to have )?(his|her) own (space|experience|adventure)|by (himself|herself)|not welcome"),
    ("Wants Child to Have Independent Experience", r"independen|independece|own experience|with ?out (a |his |her |my |one of )?(me|mom|dad|us|parent)|on (his|her|their) own|\bgrow(th)?\b|confidence|rely on (his|her)|\balone\b|\bsolo\b|autonomy|\bexperience\b|parent.?free|trust (himself|herself|themselves)|open up more|(builder|child|kid)'?s choice|it is better for (him|her)|better without (me|us)|comfort zone|push (themselves|himself|herself) out|decided to send (him|her|them)|influences? outside of (his|her|their) parents|outside influences|participate (as much|less)|wouldn'?t participate|reach out more|hold back"),
    ("Not a Traveler / Personal Preference", r"not (a |an )?(good |huge )?(traveler|fan of travel)|no desire|not comfortable|not good with|don'?t enjoy|not adventurous|not interested"),
    ("Health / Physical / Age", r"\bhealth\b|physical|\bknee\b|bad back|back (pain|injury|problem)|\bage\b|\btoo old\b|medical|injur|illness|disab|\bdeaf\b|hearing|keep up (physically )?with the kids"),
    ("Other Life Commitment", r"\bmission\b|wedding|surgery|pregnan|\bmov(e|ing)\b|deploy|trip leader|another (trip|program)|graduat|out of town|other travel plans|business trip|other activities"),
    ("Not Aware It Was an Option", r"did ?n'?t (really )?understand|not sure what|unclear|confus|not aware|wasn'?t aware|did ?n'?t realize|never heard (it|about)|did ?n'?t know (about|it)|thought (this|it) was (for|only)|what to expect|don'?t know anyone|^noth?ing\.?$|^norhing\.?$|did (not|n'?t) consider"),
    ("Not Needed", r"not needed|not necessary|no need|did ?n'?t need"),
]

GOING_CATEGORIES = [
    ("Personally Invited by HXP Staff/Trip Leader", r"trip leader (asked|invited)|hxp (office|staff) reach(ed|ing) out|asked by the trip leader"),
    ("Securing/Ensuring Child's Spot (Registration Timing)", r"(ensur|guarantee|secur|open(ed)?|free(d)? up) (a |my son'?s? |my daughter'?s? )?spot|late to sign|sign(ed)? up so|earlier selection|priority selection"),
    ("Concerned About Child Going Alone / Safety", r"wouldn'?t go with me|don'?t want to send.{0,20}alone|not willing to let|leave the country without me|too young to (go alone|travel alone)|health concern|medical concern"),
    # Terse answers that just name who influenced them, with no further elaboration -
    # kept isolated from "Other/Uncategorized" so it doesn't look like unstructured noise.
    ("Family Member Was the Reason (No Elaboration)", r"^(my |our )?(daughter|son|wife|husband|spouse|builder|youth|child)( and (my |our )?(daughter|son|wife|husband|spouse|builder|youth|child))?\.?,?!?$"),
    ("Wanted Shared Experience with Child", r"experience|memories|bond|together|time with|spend time|excit|do something with"),
    ("Child/Spouse Asked Them to Go", r"asked (for )?(me|him|her)|invited|want(ed)? me to (go|come)|want(ed)? (him|her) to (go|come)"),
    ("Love of Travel / Culture", r"travel|culture|adventure|explore"),
    ("Previous Positive Experience", r"previous|been.{0,15}before|last year|again|gone in the past|(took|take) turns|tradition|always (our )?intention|my turn"),
    ("Faith / Service Motivation", r"\bspirit(ual)?\b|\bfaith\b|\bserve\b|service|god|mission|humanitarian|\bpray|testimon(y|ies)"),
    ("Trust in the HXP Program", r"trust|impressed by hxp|good hands|love hxp|hxp (program|offers)|testimonial|q ?& ?a night|info(rmation)? (night|session)"),
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
