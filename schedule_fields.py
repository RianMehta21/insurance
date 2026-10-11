#!/usr/bin/env python3
"""Read ward class, geographical coverage and the lifetime benefit limit out
of the Benefit Schedule.

scrape_plans.py reads Part 6 Section 1(a) and 1(d). For a large family of
plans those clauses name no value at all - they cross-reference somewhere
else:

    (d) Choice of ward class
        The benefits described in these Terms and Benefits are subject to
        the restriction in the choice of ward class as stated in Section 3
        of Supplement 5 and the Benefit Schedule of these Terms and Benefits.

which is exactly what "Restricted (see doc)" / "Limited (see doc)" record.
The Benefit Schedule those clauses point at does name it, on a plain
label/value row, and every certified plan variant has its own PDF - so the
answer is per-variant and it is there to be read.

This script reads that row. By default it only writes candidate values to a
CSV for review, with the verbatim schedule text each one came from:

    python schedule_fields.py --cache pdf_cache --catalog out/plans_catalog.csv \
        --insurers AXA,Bupa,Blue Cross,FWD,Sun Life,HSBC Life,Manulife \
        --out data/schedule_candidates.csv

Review that file, then merge the values into the overrides with --apply:

    python schedule_fields.py --insurers ... --apply data/manual_overrides.csv

--apply is deliberately a second, separate run. It only fills plans whose
current value is the "(see doc)" placeholder, never overwrites a value a
person typed, and refuses to run across every insurer at once. To also
replace a SCRAPED value that disagrees with the schedule, add --correct-wrong;
every replacement it makes is printed. See apply().

Why a separate script rather than a fifth extractor inside scrape_plans.py:
that scraper runs weekly against all 33 insurers, and the Benefit Schedule
is the least standardised part of the document. This pass is opt-in, per
insurer, and its output is reviewed before it counts.
"""

import argparse
import csv
import os
import collections
import re
import subprocess
import sys
from collections import defaultdict

# --------------------------------------------------------------------------
# Locating the row
# --------------------------------------------------------------------------
# Insurers do not agree on what to call either row. Observed across the seven
# insurers this pass covers:
#
#   Entitled ward class      AXA, FWD, HSBC Life
#   Designated Ward Class    Manulife, Sun Life, Blue Cross
#   Restricted ward class    Bupa (Hero)
#   Covered ward class       Bupa (Global Prestige)
#   Covered Room             AIA
#   Covered room level       YF Life
#   Restricted accommodation room type:   Zurich (in the supplement, colon-
#                            separated, so the colon counts as the column gap)
#
#   Geographical limitation      AXA, Blue Cross, FWD, HSBC Life, Sun Life
#   Area of cover                Bupa
#   Territorial scope of cover   Manulife, FWD (vTheOne, vPrime Signature), Bowtie
#   Geographical cover           AIA
#
# A label can carry more than one footnote marker - Bowtie writes
# "Territorial scope of cover(1) (2)" - so the markers repeat.
WARD_LABEL = re.compile(
    r"^(?P<indent>\s*)(?:(?:Entitled|Designated|Restricted|Covered)\s+ward\s+class"
    r"|Covered\s+Room(?:\s+level)?|Restricted\s+accommodation\s+room\s+type\s*:)"
    r"(?:\s*\(\d+\))*"
    r"(?:\s*$|(?:\s{2,}|(?<=:)\s*)(?P<val>\S.*?)\s*$)", re.I)

GEO_LABEL = re.compile(
    r"^(?P<indent>\s*)(?:Geographical\s+(?:limitations?|cover)|Area\s+of\s+cover"
    r"|Territorial\s+scope\s+of\s+cover)"
    r"(?:\s*\(\d+\))*"
    r"(?:\s*$|\s{2,}(?P<val>\S.*?)\s*$)", re.I)

# Requiring end-of-line or a 2+ space column gap after the label is what keeps
# these off the ward-class ADJUSTMENT table, whose header reads "Entitled ward
# class as specified in the Benefit Schedule" - one space then "as", so it
# cannot match. That table is a different thing: it lists the penalty for
# being confined ABOVE your entitled class, and in several insurers'
# documents (Manulife's especially) it enumerates every ward class generically
# rather than this variant's, so reading it as the answer gives the wrong one.

# Page furniture that appears mid-cell in some documents (Sun Life's PDFs
# break a table cell across a page). Skipped while scanning, but does not end
# a cell.
FURNITURE = re.compile(
    r"^\s*(?:The\s+content\s+on\s+this\s+page\s+is\s+part\s+of|"
    r"\(No\.\s*[FS]\d+\)|Page\s+\d+|\d{1,3}|LBE\s+VHIS[-\w]*\s*V?[\w]*|"
    r"[A-Z]{2,}\s+VHIS[-\w]*\s+V\d\w*)\s*\.?\s*$")

# The adjustment-factor table's own header row also reads "Entitled ward
# class" / "Designated Ward Class", with a column gap after it, so it matches
# the label pattern above. Its cell is the REST of that header ("Confined ward
# class   Adjustment factor"), never a ward class, so it is identifiable by
# content and discarded. Without this the extractor returns the header of the
# penalty table instead of the schedule row.
ADJ_CELL = re.compile(
    r"adjustment|confined\s+ward\s+class|actual\s+ward\s+class"
    r"|ward\s+class\s+of\s+the\s+confinement|as\s+specified\s+in", re.I)

# Hard boundaries for a cell scan. Page headers sit far to the right, so they
# clear the right-column test and would otherwise be swallowed into the first
# row of the schedule - which is how FWD's geography cell came back holding
# "LBE VHIS-FPG V022026H vPrime Signature Medical Plan Benefit Schedule"
# instead of a territory.
STOP = re.compile(
    r"^\s*(?:BENEFIT\s+SCHEDULE\s*$|Benefit\s+Schedule\s*$"
    r"|Benefit\s+Level\s*:|Benefit\s+items?\b)", re.I)

# A schedule cell can be broken across a page, and FWD's vPrime Signature
# does exactly that: the first ward tier ends one page, four blank lines and
# a footer/page-number/document-code block follow, and the second tier opens
# the next. Page furniture therefore has to be stepped over rather than
# treated as the end of the cell - the run of blanks around it is why the
# tolerance is this wide - while a capped scan distance stops a cell that
# genuinely ended from running away into the rest of the document.
MAX_BLANK_RUN = 6
MAX_SCAN = 40

# A label can itself wrap onto a second line in the left column - FWD writes
# the geography row's label as "Territorial scope of cover" then
# "(Geographical limitation)(1)" underneath, with cell text continuing to the
# right of it on the same line. Narrow on purpose: a bare parenthetical only.
LABEL_CONT = re.compile(r"^\s*\([^)]*\)\s*(?:\(\d+\))?\s*$")
# ...and the same continuation with cell text alongside it on the same line.
# Matching the whole line is what keeps this safe: testing a fixed slice of
# the left column instead mis-cuts Manulife's ward cell, whose lines are
# indented raggedly enough that the slice lands mid-word.
LABEL_CONT_INLINE = re.compile(
    r"^\s*\([^)]*\)\s*(?:\(\d+\))?\s{2,}(?P<rest>\S.*?)\s*$")

# The real boundary between one schedule row and the next is the next row's
# LABEL, not a column position. Bupa centres its labels, so "Area of cover"
# sits further right than "Restricted ward class" directly beneath it and no
# indent test can separate them; without this the ward cell swallows the row
# above and reports two ward classes where the tiering is real but misread.
# A label inside parentheses is a continuation of the current label, not a new
# row, so it must not match here.
ROW_LABEL = re.compile(
    r"^\s*(?:Area\s+of\s+cover|Geographical\s+(?:limitations?|cover)|Covered\s+Room\b"
    r"|Restricted\s+accommodation\s+room\s+type"
    r"|Territorial\s+scope\s+of\s+cover|Extended\s+territorial\s+scope"
    r"|Deductible\b|Annual\s+Benefit\s+Limit|Lifetime\s+Benefit\s+Limit"
    r"|Aggregate\s+limit|Benefit\s+items?\b|Benefit\s+Coverage"
    r"|(?:Entitled|Designated|Restricted|Covered)\s+ward\s+class)", re.I)


def is_blank(line):
    return not line.strip()


def as_cell(raw):
    return [(len(t) - len(t.lstrip()), t.strip()) for t in raw if t.strip()]


def cell_at(lines, idx, m, answers):
    """Return the full text of the table cell whose label is on line `idx`.

    The schedules are two-column tables rendered by `pdftotext -layout`, so a
    cell is "the run of lines that have content in the right-hand column and
    nothing in the left". Walking both up and down from the label matters:
    Manulife centres the label vertically inside a tall wrapped cell, so most
    of the value sits ABOVE the word that names it -

        Designated Ward Class     Hospitals offer various accommodation ...
        ^ label on line 3 of 9    ^ value started 6 lines earlier

    - and a scan that only went downwards would return the disclaimer
    paragraph instead of the ward class.
    """
    label_indent = len(m.group("indent"))

    # Membership is decided by "is this line clear of the label column", not
    # by the exact column the value happens to start in. Insurers wobble the
    # indent of a wrapped cell by a few characters - Manulife's Deluxe ward
    # cell starts at 40 on one line and 43 on the next - and keying off the
    # first value's column drops the rest of the cell. The next row's label
    # sits back at the label column, so it still ends the scan.
    threshold = label_indent + 6

    def classify(line):
        """-> (keep, text) where keep is True/False/None (None = skip)."""
        if is_blank(line):
            return None, ""
        if STOP.match(line) or ROW_LABEL.match(line):
            return False, ""                  # next row, or end of the table
        if FURNITURE.match(line):
            return None, ""
        if LABEL_CONT.match(line):
            return None, ""
        mc = LABEL_CONT_INLINE.match(line)
        if mc:
            rest = mc.group("rest")
            return True, " " * line.rindex(rest) + rest
        ind = len(line) - len(line.lstrip())
        return (ind > threshold), line.rstrip()

    out_up, out_down = [], []

    if m.groupdict().get("val"):
        # Keep the value at its true column: split_columns reads a three-column
        # row apart by indentation, so a line stored stripped would collapse
        # into the leftmost cluster and the row would look one-dimensional.
        col = lines[idx].rindex(m.group("val"))
        out_down.append(" " * col + m.group("val"))

    blanks = 0
    for j in range(idx + 1, min(len(lines), idx + 1 + MAX_SCAN)):
        keep, text = classify(lines[j])
        if keep is None:
            blanks += 1
            if blanks >= MAX_BLANK_RUN:
                break
            continue
        if not keep:
            break
        blanks = 0
        out_down.append(text)

    # Scan upward only when the cell cannot already be read downward.
    #
    # Almost every schedule row is top-aligned: the value starts on the label's
    # own line. Scanning up from one of those walks straight into the tail of
    # the row ABOVE - AXA's ward cell came back holding the geography row's
    # "For non-Emergency Treatment: Asia ..." continuation lines, which happened
    # not to change the answer but would the moment a geography cell mentioned
    # a ward class. Reserve the upward scan for the two shapes that need it:
    # a label sitting alone on its line (Bupa, Manulife Supreme), and a label
    # whose inline text is boilerplate that answers nothing (Manulife Supreme
    # Lite, whose ward row opens with the "Hospitals offer various
    # accommodation options" advisory and keeps the class above it).
    label_alone = not m.groupdict().get("val")
    if label_alone or not answers(join_cell(as_cell(out_down))):
        blanks = 0
        for j in range(idx - 1, max(-1, idx - 1 - MAX_SCAN), -1):
            keep, text = classify(lines[j])
            if keep is None:
                blanks += 1
                if blanks >= MAX_BLANK_RUN:
                    break
                continue
            if not keep:
                break
            blanks = 0
            out_up.append(text)

    return as_cell(list(reversed(out_up)) + out_down)


def find_cells(text, label_re, answers):
    """Every schedule cell whose label matches, in document order.

    Cells that are really the adjustment table's header are dropped here
    rather than by position, because the two tables are not reliably ordered:
    Bupa prints the schedule before the supplement, HSBC after it.
    """
    lines = text.split("\n")
    found = []
    for i, line in enumerate(lines):
        m = label_re.match(line)
        if not m:
            continue
        cell = cell_at(lines, i, m, answers)
        if not cell:
            continue
        text_form = join_cell(cell)
        if not text_form or ADJ_CELL.search(text_form):
            continue
        found.append((i, cell, text_form))
    return found


def join_cell(cell):
    return squash(" ".join(t for _, t in cell))


def squash(s):
    s = s.replace("–", "-").replace("—", "-")
    s = s.replace("’", "'").replace("‘", "'")
    s = re.sub(r"\s+", " ", s).strip()
    # pdftotext keeps the PDF's own line-break hyphen, so a ward class wrapped
    # across two lines arrives as "Standard Semi- private Room" and matches
    # neither "Semi-private" nor "Private" cleanly. Rejoin only genuine
    # intra-word breaks: letter, hyphen, space, lowercase letter. An em-dash
    # introducing a clause is always followed by a capital or a digit, so it
    # is left alone.
    return re.sub(r"([A-Za-z])-\s+([a-z])", r"\1-\2", s)


# --------------------------------------------------------------------------
# Normalising the value
# --------------------------------------------------------------------------
# Every ward class the seven insurers actually name, longest first so that
# "Standard Private Room" is never matched as "Private Room" and "Standard
# Semi-private Room" never as "Semi-private Room".
WARD_TERMS = [
    # "Semi\s*-?\s*private": AIA's PDFs break the word as "Semi- Private",
    # which squash() leaves alone because the second half is capitalised.
    (r"Standard\s+Semi\s*-?\s*private\s+Room", "Standard Semi-private"),
    (r"Standard\s+Private\s+Room",           "Standard Private"),
    (r"Standard\s+Ward\s+Room",              "Standard Ward"),
    (r"Semi\s*-?\s*private\s+Room",          "Semi-private"),
    (r"General\s+Ward",                      "General Ward"),
    (r"Private\s+Room",                      "Private"),
    (r"Ward\s+Room",                         "Ward"),
    (r"\bSuite\b",                           "Suite"),
    # Blue Cross names its lowest class just "Ward", with no "Room" after it.
    # Last, so every longer name is consumed first, and guarded against the
    # phrase "ward class", which is boilerplate rather than an answer.
    (r"\bWard\b(?!\s+class)",                "Ward"),
]

# Location phrases, mapped to the short tag used in the output string.
# Order matters: the first match wins, so the narrower "except Hong Kong ..."
# forms must be tried before the bare territory they carve out of.
LOCATIONS = [
    (r"worldwide\s*\(excluding\s+Hong\s*Kong,?\s*Macau,?\s*Mainland\s*China"
     r"\s*and\s*the\s*USA\)",                                      "elsewhere"),
    (r"outside\s+Asia\s*&?\s*Australia-New\s*Zealand",             "outside Asia/Aus/NZ"),
    (r"Asia\s*&\s*Australia-New\s*Zealand\s*\(except",             "rest of Asia/Aus/NZ"),
    (r"other\s+areas\s+in\s+Asia,?\s*Australia\s*and\s*New\s*Zealand", "elsewhere"),
    (r"other\s+areas\s+worldwide",                                 "elsewhere"),
    (r"anywhere\s+else\s+in\s+the\s+world",                        "elsewhere"),
    (r"outside\s+Hong\s*Kong,?\s*Macau\s*(?:and|or)\s*Mainland\s*China", "elsewhere"),
    (r"excluding\s+Hong\s*Kong,?\s*Macau\s*(?:and|or)\s*Mainland\s*China", "elsewhere"),
    # Prudential: "anywhere else in Asia except Hong Kong, Macau or mainland
    # China". Missing this made the branch read as Hong Kong itself.
    (r"except\s+Hong\s*Kong,?\s*Macau\s*(?:and|or)\s*Mainland\s*China", "elsewhere"),
    (r"(?:except|excluding)\s+Hong\s*Kong,?\s*Australia\s*and\s*New\s*Zealand",  "elsewhere"),
    # Chow Tai Fook ChampCare. The "outside" form must come first: it names
    # the same four places.
    (r"outside\s+Hong\s*Kong,?\s*Macau,?\s*Australia\s*and\s*New\s*Zealand",
                                                                   "elsewhere"),
    (r"Hong\s*Kong,?\s*Macau,?\s*Australia\s*and\s*New\s*Zealand", "HK/Macau/Australia/NZ"),
    (r"Hong\s*Kong,?\s*Macau,?\s*Taiwan\s*and\s*mainland\s*China", "HK/Macau/Taiwan/China"),
    (r"Hong\s*Kong,?\s*Macau\s*(?:or|and)\s*Mainland\s*China",     "HK/Macau/China"),
    (r"Hong\s*Kong,?\s*Australia\s*and\s*New\s*Zealand",           "HK/Australia/NZ"),
]


def ward_terms_in(s):
    """Ward classes named in `s`, in document order, longest-match-first so
    "Standard Private Room" is never read as "Private Room"."""
    hits, scan = [], s
    for pat, label in WARD_TERMS:
        for mm in re.finditer(pat, scan, re.I):
            hits.append((mm.start(), label))
        scan = re.sub(pat, lambda m: " " * (m.end() - m.start()), scan, flags=re.I)
    return [label for _, label in sorted(hits)]


def location_of(s):
    for pat, tag in LOCATIONS:
        if re.search(pat, s, re.I):
            return tag
    return None


def ward_answers(text):
    return bool(ward_terms_in(strip_definitions(text)))


def geo_answers(text):
    return bool(match_geo(text))


def strip_definitions(cell):
    """Remove the parenthetical room definitions and the advisory paragraph.

    Both are boilerplate that repeats the ward class names, so leaving them in
    makes a single-tier cell look like a multi-tier one.
    """
    cell = re.sub(r"\(i\.e\.,?.*?\)", " ", cell, flags=re.I)
    cell = re.sub(r"Hospitals\s+offer\s+various\s+accommodation.*$", " ",
                  cell, flags=re.I)
    cell = re.sub(r"If\s+you\s+are\s+unsure.*$", " ", cell, flags=re.I)
    return squash(cell)


def split_columns(cell):
    """Bupa prints the ward row as THREE columns - label | where | which ward:

        Restricted ward class   For Hong Kong, Macau,        Semi-private Room
                                Taiwan and mainland China
                                For other areas in Asia,     Standard Private Room
                                Australia and New Zealand

    Flattening that to one string interleaves the two columns and the ward
    classes end up detached from the places they apply to. Where the deeper
    indent cluster holds nothing but ward classes, read the columns apart and
    pair them in document order.
    """
    if len(cell) < 3:
        return None
    indents = sorted({ind for ind, _ in cell})
    if len(indents) < 2:
        return None
    # Cluster the indents at their widest gap. Taking the single deepest
    # indent instead splits Bupa's own ward column in half, because
    # "Semi-private Room" and "Standard Private Room" are centred and so do
    # not start at the same column.
    gaps = [(indents[i + 1] - indents[i], i) for i in range(len(indents) - 1)]
    width, at = max(gaps)
    if width < 8:
        return None
    split = indents[at + 1]
    right = [t for ind, t in cell if ind >= split]
    left = [t for ind, t in cell if ind < split]
    if not right or not left:
        return None
    # The right column must be exactly one ward class per line and nothing
    # else - otherwise this is not the layout we think it is, and guessing
    # from a misread table is precisely the failure mode worth avoiding.
    stripper = "|".join(p for p, _ in WARD_TERMS)
    for t in right:
        if len(ward_terms_in(t)) != 1:
            return None
        if re.sub(stripper, "", t, flags=re.I).strip():
            return None
    locs = re.split(r"(?=For\s+)", squash(" ".join(left)), flags=re.I)
    locs = [l for l in locs if l.strip()]
    if len(locs) != len(right):
        return None
    branches = []
    for loc_text, ward_text in zip(locs, right):
        loc = location_of(loc_text)
        wards = ward_terms_in(ward_text)
        if not loc or not wards:
            return None
        branches.append((loc, wards[0]))
    return branches


def urgency_split(part):
    """Blue Cross's Love Yourself sets the ward class by urgency WITHIN each
    territory ("(i) Emergency Treatment: Ward / (ii) Non-Emergency: Ward"),
    so a branch can carry two answers rather than one."""
    em = re.search(r"(?<!Non-)(?<!Non\s)\bEmergency\s+Treatment\s*:\s*(.{0,40})", part, re.I)
    non = re.search(r"Non-Emergency\s+Treatment\s*:\s*(.{0,40})", part, re.I)
    if not (em and non):
        return None
    ew = ward_terms_in(em.group(1))
    nw = ward_terms_in(non.group(1))
    if not ew or not nw:
        return None
    if ew[0] == nw[0]:
        return nw[0]
    return f"{nw[0]}, {ew[0]} for emergency"


def normalise_ward(cell):
    """-> (value, confidence) using the agreed both-tier representation.

    Bupa/Blue Cross/FWD/Manulife all vary the ward class by where the
    Confinement happens, and the tiers differ, so collapsing to one is a real
    loss of meaning: Manulife Supreme (Advance) is Semi-private in Hong Kong
    but Standard Private everywhere else.
    """
    cols = split_columns(cell)
    if cols:
        return "; ".join(f"{w} ({loc})" for loc, w in cols), "tiered"

    body = strip_definitions(join_cell(cell))

    branches = []
    # "or Confinement outside Asia for Emergency Treatment" continues the
    # branch before it rather than opening a new one - FWD's vPrime pairs one
    # ward class with two territories that way - so a split there would strand
    # the ward class in a fragment of its own and make the row look ambiguous.
    parts = re.split(r"(?<!or )(?=(?:Confinement\s+in\b|Confinement\s+outside\b"
                     r"|Confinement\s+anywhere\b"
                     r"|For\s+(?:Hong|other|anywhere)))", body, flags=re.I)
    for part in parts:
        if not part.strip():
            continue
        loc = location_of(part)
        if not loc:
            continue
        ward = urgency_split(part)
        if ward is None:
            wards = ward_terms_in(part)
            ward = wards[0] if wards else None
        if ward:
            branches.append((loc, ward))

    if len(branches) >= 2:
        seen, uniq = {}, []
        for loc, w in branches:
            if loc in seen:
                # The same place named twice with different wards means a
                # branch's location was misread. Keeping the first silently
                # dropped Prudential's "Private (elsewhere)" tier this way.
                if seen[loc] != w:
                    return "", f"conflicting tiers for {loc}: {seen[loc]} / {w}"
                continue
            seen[loc] = w
            uniq.append(f"{w} ({loc})")
        return "; ".join(uniq), "tiered"

    hits = ward_terms_in(body)
    if len(set(hits)) == 1:
        return hits[0], "single"
    if len(set(hits)) > 1:
        return "", f"ambiguous: {sorted(set(hits))}"
    return "", "no ward term found"


GEO_TERMS = [
    (r"Worldwide\s+excluding\s+(?:the\s+)?(?:United\s+States|USA)", "Worldwide excl. USA"),
    (r"Worldwide\s+except\s+(?:the\s+)?(?:United\s+States|USA)",    "Worldwide excl. USA"),
    (r"Asia\s*(?:,|&|and)?\s*(?:including\s+)?Australia\s*(?:and|-|&)?\s*New\s*Zealand",
     "Asia incl. Australia & NZ"),
    (r"Greater\s+China",                                            "Greater China"),
    (r"\bWorldwide\b",                                              "Worldwide"),
    (r"\bAsia\b",                                                   "Asia"),
    (r"Hong\s*Kong\s*(?:and|/|,)?\s*Macau",                         "Hong Kong/Macau"),
]


def normalise_geo(cell):
    """-> (value, confidence).

    AXA and FWD split the scope by treatment urgency: a narrow scope for
    ordinary treatment, worldwide for emergencies. Both halves are recorded.
    The leading "Except for psychiatric treatments ... -" carve-out is dropped;
    it names a benefit item, not the plan's scope, and it is what made the old
    clause-level regex return "Hong Kong/Macau" for AXA WiseGuard Pro, whose
    actual limitation is Greater China.
    """
    body = squash(cell)
    body = re.sub(r"^Except\s+for\s+.*?[-–]\s*", " ", body, flags=re.I)
    body = re.sub(r"^Except\s+for\s+.*?,\s*(?=all\s+benefits)", " ", body, flags=re.I)
    body = re.sub(r"Except\s+the\s+benefit\s+items\s+below.*$", " ", body, flags=re.I)
    body = re.sub(r"\(\d+\)", " ", body)
    body = squash(body)

    non_em = re.search(r"For\s+non[\s-]?Emergency\s+Treatment\s*:\s*(.+?)"
                       r"(?=For\s+Emergency|$)", body, re.I)
    em = re.search(r"(?<!non[\s-])(?<!non)For\s+Emergency\s+Treatment\s*:\s*(.+?)$",
                   body, re.I)
    if non_em and em:
        a = match_geo(non_em.group(1))
        b = match_geo(em.group(1))
        if a and b:
            return (f"{a} ({b} for emergency)" if a != b else a), "split"
        return "", f"split but unmatched: {non_em.group(1)[:40]!r}/{em.group(1)[:40]!r}"

    v = match_geo(body)
    if v:
        return v, "single"
    return "", "no geography term found"


def match_geo(s):
    for pat, label in GEO_TERMS:
        if re.search(pat, s, re.I):
            return label
    return None


# Some plans have no geography row at all and put the answer in the
# supplement's prose instead - Asia Elite and Chubb's VHIS Series Flexi
# ("subject to a geographical limitation of Asia"), Allianz Ultra Medical
# Protect ("payable only if the relevant expenses are incurred in Asia").
# Only consulted when the schedule has no row, and labelled as such so the
# review can tell the two sources apart.
GEO_PROSE = re.compile(
    r"subject\s+to\s+(?:a|the)\s+geographical\s+limitation\s+of\s+(?P<a>[A-Z][\w ,&-]{2,60}?)\s*\."
    r"|payable\s+only\s+if\s+the\s+relevant\s+expenses\s+are\s+incurred\s+in\s+(?P<b>[A-Z][\w ,&-]{2,60}?)\s*\."
    # YF Life: "all benefits ... shall be applicable in Asia."
    r"|all\s+benefits\s+described\s+in\s+these\s+Terms\s+and\s+Benefits\s+shall\s+be\s+applicable\s+in\s+(?P<c>[A-Z][\w ,&-]{2,60}?)\s*\.")


def geo_from_prose(text):
    """-> (evidence, value, confidence) from the supplement's wording."""
    found = {}
    for m in GEO_PROSE.finditer(squash(text)):
        v = match_geo(m.group("a") or m.group("b") or m.group("c"))
        if v:
            found.setdefault(v, m.group(0))
    if len(found) == 1:
        v, ev = next(iter(found.items()))
        return ev, v, "prose (no schedule row)"
    if found:
        return "", "", f"prose disagrees: {sorted(found)}"
    return "", "", "no geography row in schedule"


# --------------------------------------------------------------------------
# Lifetime Benefit Limit
# --------------------------------------------------------------------------
# Same shape of gap as ward and geography: Part 6 Section 1(b) says only "All
# benefits ... are subject to the Lifetime Benefit Limit as stated in the
# Benefit Schedule", which scrape_plans.py records as "Capped (see doc)". The
# schedule row names the figure:
#
#     Lifetime Benefit Limit for benefit       $20,000,000 per life
#     items I (a) - (r) and II (a) - (h)
#
# The value is always one short amount (or "Nil"), so this does not need the
# general cell reader above - which in fact cannot read Manulife's row: the
# label wraps over four left-column lines and the amount is centred on a line
# with nothing at all in the label column. Instead: start at the label, read
# the right-hand column for a few lines, stop at a blank line or the next row.
#
# Manulife and Sun Life print a bare "$". The currency is then the plan's own,
# which is why normalise_lifetime() takes it; a stated currency that disagrees
# with the plan's is reported, not trusted.
LIFETIME_LABEL = re.compile(r"^(?P<indent>\s*)Lifetime\s+Benefit\s+Limit\b", re.I)
# A line that starts with the label but is really mid-sentence prose: Sun Life
# wraps "...Annual Benefit Limit and the" / "Lifetime Benefit Limit." so the
# label opens a line and is followed straight by punctuation or a verb.
LIFETIME_PROSE = re.compile(
    r"Lifetime\s+Benefit\s+Limit\s*(?:[.;,]|(?:as|and|is|shall|of)\b)", re.I)
LIFETIME_AMOUNT = re.compile(r"(HKD|USD|HK\$|US\$|\$)\s*([\d,]{5,})", re.I)
LIFETIME_NIL = re.compile(r"^(?:nil|no\s+limit|unlimited|not\s+applicable)\b", re.I)
LIFETIME_SCAN = 6


def _right_column(line, label_indent):
    """Text in the value column: after a 2+ space gap, clear of the label."""
    # Test every gap, not just the first: HSBC prints the amount to the right
    # of the label's wrapped second line ("items (a) - (l) and Enhanced
    # HKD20,000,000"), so the leading indent is a gap that comes first.
    for mm in re.finditer(r"\s{2,}(?=\S)", line):
        if mm.end() > label_indent + 6:
            return line[mm.end():].rstrip()
    return ""


def amount_row_cells(text, label, prose):
    """Every schedule row's value text for `label`, in document order.

    Shared by every one-figure row (lifetime limit, annual limit): the label
    starts its line, and its value is whatever sits in the right-hand column
    for a few lines, up to a blank line or the next row."""
    lines = text.split("\n")
    out = []
    for i, line in enumerate(lines):
        m = label.match(line)
        if not m or prose.search(line):
            continue
        indent = len(m.group("indent"))
        parts = [_right_column(line, indent)]
        for j in range(i + 1, min(len(lines), i + 1 + LIFETIME_SCAN)):
            nxt = lines[j]
            if is_blank(nxt) or STOP.match(nxt) or ROW_LABEL.match(nxt):
                break
            # A line with nothing in the label column is all value (Manulife's
            # centred amount); otherwise only its right-hand column is.
            parts.append(nxt.strip() if not nxt[:indent + 6].strip()
                         else _right_column(nxt, indent))
        out.append(squash(" ".join(p for p in parts if p)))
    return out


def lifetime_cells(text):
    """Every Lifetime Benefit Limit schedule row's value text, in order."""
    return amount_row_cells(text, LIFETIME_LABEL, LIFETIME_PROSE)


def normalise_lifetime(cell, currency):
    """-> (value, confidence). Value is a whole number in the plan's currency,
    "None" for a schedule that says Nil, or "" when nothing usable is there."""
    # The first value in the cell is the answer. The scan can run on into the
    # next row's figures - Blue Cross follows a "Nil" lifetime row with
    # "$1,000 per Day Case Procedure" - so anything after it is only flagged.
    if LIFETIME_NIL.match(cell):
        extra = LIFETIME_AMOUNT.search(cell)
        return "None", "nil" + (" (cell runs on into other figures)" if extra else "")
    m = LIFETIME_AMOUNT.search(cell)
    if not m:
        return "", "no amount"
    stated = m.group(1).upper().replace("HK$", "HKD").replace("US$", "USD")
    if stated != "$" and stated != currency:
        return "", f"currency {stated} is not the plan's {currency}"
    rest = LIFETIME_AMOUNT.search(cell, m.end())
    return (m.group(2).replace(",", ""),
            "single" + (" (cell runs on into other figures)" if rest else ""))


def first_amount(cells, currency, missing):
    """-> (value, confidence, evidence). The first row that yields a value
    wins: AIA prints a second schedule for an optional benefit further on,
    with its own "Nil" lifetime row, and appends the Standard Plan's
    HKD420,000 annual limit after its own. Neither is the plan's figure. How
    many distinct values the document holds goes into the confidence, so a
    disagreement is visible in review."""
    if not cells:
        return "", missing, ""
    read = [(normalise_lifetime(c, currency), c) for c in cells]
    hits = [(v, conf, c) for (v, conf), c in read if v]
    if not hits:
        (_, conf), c = read[0]
        return "", conf, c
    value, conf, evidence = hits[0]
    others = sorted({v for v, _, _ in hits} - {value})
    if others:
        conf += f" (document also states {', '.join(others)})"
    return value, conf, evidence


def read_lifetime(text, currency):
    return first_amount(lifetime_cells(text), currency,
                        "no lifetime row in schedule")


# --------------------------------------------------------------------------
# Annual Benefit Limit
# --------------------------------------------------------------------------
# scrape_plans.py reads this from the clause and, on AIA's and Chubb's Flexi
# plans, picks up the Standard Plan's HKD420,000 - printed in the same PDF
# because a Flexi plan must restate Standard Plan terms. 38 Flexi plans showed
# that figure instead of their own HKD5,000,000-12,000,000. The plan's own
# schedule row comes first in the document, so first_amount() takes it.
ANNUAL_LABEL = re.compile(r"^(?P<indent>\s*)Annual\s+Benefit\s+Limit\b", re.I)
ANNUAL_PROSE = re.compile(
    r"Annual\s+Benefit\s+Limit\s*(?:[.;,]|(?:as|and|is|shall|of|has)\b)", re.I)
STANDARD_ANNUAL_LIMIT = "420000"


def read_annual(text, currency):
    # "Nil" means no overall annual cap, exactly as on the lifetime row: AXA
    # Smart Medicare, Prudential FlexiChoice and Manulife First, among others,
    # limit individual benefits but not the year as a whole.
    return first_amount(amount_row_cells(text, ANNUAL_LABEL, ANNUAL_PROSE),
                        currency, "no annual limit row in schedule")


# --------------------------------------------------------------------------
# Coinsurance
# --------------------------------------------------------------------------
# VHIS mandates coinsurance on one benefit only: Prescribed Diagnostic Imaging
# Tests (30% on the Standard Plan). Flexi plans state their own rate on that
# schedule row, and most of the ones the scraper left blank say "Fully
# covered" there - 0%, a phrase scrape_plans.py does not recognise.
IMAGING_LABEL = re.compile(
    # "Diagnostic" alone: AIA wraps the label as "(i) Prescribed Diagnostic"
    # with "Imaging Tests" on the next line.
    r"^(?P<indent>\s*)(?:\(\w{1,4}\)\s*)?Prescribed\s+Diagnostic\b", re.I)
NO_PROSE = re.compile(r"(?!)")
COINS_PCT = re.compile(r"(\d{1,2})\s*%\s*(?:of\s+(?:the\s+)?Eligible\s+Expenses\s+)?Co-?insurance", re.I)
COINS_NONE = re.compile(r"Fully\s+covered|No\s+Co-?insurance|Co-?insurance\s*:?\s*(?:Nil|0\s*%)", re.I)


SCHEDULE_HEADING = re.compile(r"^\s*BENEFIT\s+SCHEDULE\s*$|^\s*Benefit\s+Schedule\s*$")
IMAGING_SCAN = 14
# The next schedule item - "(j) Prescribed Non-surgical Cancer Treatments" -
# ends the imaging row. Blank lines do not: FlexiCare prints "Fully covered"
# then, after a gap, "Performed during Confinement: subject to 30%
# Coinsurance", and stopping at the gap reported 0% for a plan that charges
# 30% on inpatient scans.
NEXT_ITEM = re.compile(r"^\s*\((?:[a-z]{1,3}|[ivx]{1,5})\)\s+\S", re.I)


def read_coinsurance(text):
    """-> (percent, confidence, evidence) from the imaging row.

    Only a row that states one rate gives a figure. Chow Tai Fook's TopCare
    charges 30% during Confinement and 20% as a Day Patient; a single number
    for that would be wrong whichever one was picked, so it is reported as
    varying and left for the plan document."""
    lines = text.split("\n")
    # Part 6 describes the imaging benefit in prose, mentioning coinsurance in
    # general terms, long before the schedule states this plan's rate. Start
    # at the first schedule heading; a document without one is scanned whole.
    first = next((i for i, l in enumerate(lines) if SCHEDULE_HEADING.match(l)), 0)
    for i, line in enumerate(lines):
        if i < first or not IMAGING_LABEL.match(line):
            continue
        block = [line]
        for nxt in lines[i + 1:i + 1 + IMAGING_SCAN]:
            if NEXT_ITEM.match(nxt) or STOP.match(nxt):
                break
            block.append(nxt)
        cell = squash(" ".join(block))
        rates = {m.group(1) for m in COINS_PCT.finditer(cell)}
        if COINS_NONE.search(cell):
            rates.add("0")
        if len(rates) == 1:
            return rates.pop(), "imaging row", cell
        if len(rates) > 1:
            return "", f"varies by setting {sorted(rates, key=int)}", cell
        # Only the plan's own (first) row counts. AIA centres "Fully covered"
        # between its (i) and (j) rows, so its row reads empty - and carrying
        # on would reach the Standard Plan schedule AIA appends, whose 30% is
        # not this plan's rate.
        return "", "imaging row states no rate of its own", cell
    return "", "no imaging row in schedule", ""

# --------------------------------------------------------------------------
# Cross-check: the ward-class adjustment table
# --------------------------------------------------------------------------
ADJ_HEADER = re.compile(
    r"(?:Entitled|Designated)\s+[Ww]ard\s+[Cc]lass(?:\s+as\s+specified)?", re.I)


def adjustment_designated(text):
    """Ward classes named in the LEFT column of the adjustment-factor table.

    Returns a set. Only a cross-check, and only a usable one when the table
    names exactly one class: several insurers (Manulife notably) print a
    generic table listing every ward class the product line offers, identical
    across levels, so a multi-value result carries no information about THIS
    variant and is reported as inconclusive rather than as a conflict.
    """
    # Anchor on the supplement's numbered SECTION HEADING, at the start of a
    # line. Searching the phrase anywhere lands on Part 6's prose instead -
    # "...subject to the restriction in the choice of ward class as stated in
    # Section 3 of Supplement 5..." - which is the cross-reference, not the
    # table, and yields nothing.
    heads = [m.start() for m in re.finditer(
        r"^\s*(?:\(?[0-9IVX]+\)?\.?\s*)?"
        r"(?:(?:Limitation|Restriction)s?\s+(?:on|in)\s+(?:the\s+)?choice"
        r"\s+of\s+ward\s+class"
        r"|Choice\s+of\s+ward\s+class\s+and\s+adjustment)",
        text, re.I | re.M)]
    for start in heads:
        found = set()
        for line in text[start:start + 3000].split("\n"):
            s = squash(line)
            if not s:
                continue
            # A row of the adjustment table, not a sentence: it opens with a
            # ward class and carries a percentage or a second ward class.
            for pat, label in WARD_TERMS:
                if re.match(pat, s, re.I):
                    if re.search(r"%|\bnil\b", s, re.I) or len(ward_terms_in(s)) > 1:
                        found.add(label)
                    break
        if found:
            return found
    return set()


# --------------------------------------------------------------------------
def pdf_text(cache, cert):
    path = os.path.join(cache, f"{cert}.pdf")
    if not os.path.exists(path):
        return None
    try:
        return subprocess.run(["pdftotext", "-layout", path, "-"],
                              capture_output=True, text=True,
                              timeout=120).stdout
    except Exception:
        return None


# The two placeholders this pass exists to replace. A plan showing anything
# else already has an answer from the PDF clause or the certified plan name,
# and is left alone: several plans whose clause reads "not subject to any
# restriction in the choice of ward class" still print a "Designated Ward
# Class" in their schedule, because there it sets the room-and-board benefit
# limit and names the pricing tier rather than restricting anything. Manulife
# First and Bupa MyFlexi are both that shape, and overwriting their honest
# "No restriction" with a ward class read out of the schedule would be wrong.
WARD_PLACEHOLDER = "Restricted (see doc)"
GEO_PLACEHOLDER = "Limited (see doc)"
LIFETIME_PLACEHOLDER = "Capped (see doc)"

# Values --correct-wrong must never replace, however loudly the schedule
# seems to disagree. "No restriction" is not a gap or a bad guess: it is the
# Part 6(d) clause answering the VHIS question outright ("All benefits ...
# are not subject to any restriction in the choice of ward class"). Plans
# shaped that way can still print a "Designated Ward Class" in their
# schedule, where it sets the room-and-board benefit limit and names the
# pricing tier - Manulife First's levels are literally called "(Semi-private
# Room)", "(Private Room)" - while only an optional supplementary benefit
# carries a real ward restriction. Manulife First (9) and Bupa MyFlexi (3)
# are both that shape; replacing their "No restriction" would state a
# restriction the certified terms do not impose.
NEVER_REPLACE = {"ward": {"No restriction"},
                 # Likewise "None" is Part 6(b) saying outright that the plan
                 # is "not subject to any Lifetime Benefit Limit".
                 "lifetime": {"None"}}


def apply(candidates, overrides_path, catalog, benefits_path, correct_wrong=False):
    """Merge reviewed candidates into manual_overrides.csv.

    By default only fills a field whose effective value today is the
    placeholder, and never touches a cell a person typed.

    With correct_wrong, it also replaces a scraped value that disagrees with
    the schedule. That is a separate flag because it is a different and
    riskier act: replacing an admitted gap is safe, while overwriting a value
    the workbook currently states confidently needs the evidence looked at
    first. It exists because the clause-level regex can match text that is
    not the answer - AXA WiseGuard Pro's territorial clause names Hong Kong
    and Macau only inside the list of benefits EXCLUDED from the limitation,
    and all 24 variants consequently read "Hong Kong/Macau" when the schedule
    says Greater China, Asia or Worldwide excluding USA depending on level.
    Every replacement is printed, so a wrong one is visible rather than
    silent.
    """
    import csv as _csv

    def current(cert):
        """What the workbook shows for this plan now, per make_workbook's
        trust order: scraped value, else the certified plan-level label."""
        s = scraped.get(cert, {})
        c = catalog.get(cert, {})
        ward = (s.get("ward_restriction") or "").strip()
        geo = (s.get("geographical_coverage") or "").strip()
        if ward == WARD_PLACEHOLDER:
            ward = (c.get("ward_type") or "").strip() or ward
        if geo == GEO_PLACEHOLDER:
            geo = (c.get("geographical_coverage") or "").strip() or geo
        life = (s.get("lifetime_benefit_limit") or "").strip()
        annual = (s.get("annual_benefit_limit") or "").strip()
        coins = (s.get("coinsurance_pct") or "").strip()
        return ward, geo, life, annual, coins

    scraped = {r["certification_no"]: r
               for r in _csv.DictReader(open(benefits_path, encoding="utf-8-sig"))}

    with open(overrides_path, encoding="utf-8-sig", newline="") as f:
        rd = _csv.DictReader(f)
        rows = list(rd)
        cols = list(rd.fieldnames or [])
    by_cert = {r["certification_no"]: r for r in rows}

    filled = collections.Counter()
    skipped, corrected = [], []
    for cand in candidates:
        cert = cand["certification_no"]
        cat = catalog.get(cert)
        if not cat:
            continue
        cur_ward, cur_geo, cur_life, cur_annual, cur_coins = current(cert)
        # On a Flexi plan the Standard Plan's annual limit is the scraper
        # reading the wrong schedule (see ANNUAL_LABEL), so it is treated as
        # a gap, not an answer. A Standard plan's 420,000 is its real limit.
        annual_gap = (STANDARD_ANNUAL_LIMIT if cat.get("plan_type") != "Standard"
                      else WARD_PLACEHOLDER)
        want = {}
        for field, cur, placeholder in (("ward", cur_ward, WARD_PLACEHOLDER),
                                        ("geography", cur_geo, GEO_PLACEHOLDER),
                                        ("lifetime", cur_life, LIFETIME_PLACEHOLDER),
                                        ("annual_limit", cur_annual, annual_gap),
                                        ("coinsurance", cur_coins, None)):
            new = cand.get(field, "")
            if not new:
                continue
            if cur in (placeholder, ""):
                want[field] = new
            elif (correct_wrong and new != cur and field != "coinsurance"
                  and cur not in NEVER_REPLACE.get(field, ())):
                want[field] = new
                if (by_cert.get(cert, {}).get(field) or "").strip() != new:
                    corrected.append((cert, field, cur, new))
        if not want:
            if any(cand.get(f) for f in ("ward", "geography", "lifetime",
                                         "annual_limit", "coinsurance")):
                skipped.append((cert, cur_ward, cur_geo))
            continue

        # A cell that already holds this value needs nothing done to it, so
        # re-running --apply is a no-op rather than a rewrite. Without this
        # the second run reports every row as "updated" and re-lists the
        # corrections it made the first time, which reads like new activity.
        want = {k: v for k, v in want.items()
                if (row_existing := by_cert.get(cert)) is None
                or (row_existing.get(k) or "").strip() != v}
        if not want:
            continue

        row = by_cert.get(cert)
        if row is None:
            row = {c: "" for c in cols}
            row.update(certification_no=cert, insurer=cat.get("insurer", ""),
                       plan_name=cat.get("plan_name", ""),
                       plan_level=cat.get("plan_level_raw", ""),
                       plan_doc_url=cat.get("plan_doc_url", ""))
            rows.append(row)
            by_cert[cert] = row
            filled["new rows"] += 1
        else:
            filled["existing rows updated"] += 1
        for k, v in want.items():
            if (row.get(k) or "").strip():
                continue                      # never overwrite a typed value
            row[k] = v
            filled[k] += 1
        note = "read from Benefit Schedule"
        if not (row.get("note") or "").strip():
            row["note"] = note

    # utf-8-sig, matching worklist.py: this file is opened in Excel, and
    # dropping the BOM makes Excel mis-read the first column header.
    with open(overrides_path, "w", encoding="utf-8-sig", newline="") as f:
        w = _csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c, "") for c in cols})
    return filled, skipped, corrected


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--catalog", default="out/plans_catalog.csv")
    ap.add_argument("--cache", default="pdf_cache")
    ap.add_argument("--insurers", default="",
                    help="comma-separated insurer name prefixes; blank = all")
    ap.add_argument("--out", default="data/schedule_candidates.csv")
    ap.add_argument("--benefits", default="data/plan_benefits.csv")
    ap.add_argument("--apply", metavar="OVERRIDES_CSV",
                    help="merge the reviewed values into this overrides file. "
                         "Requires --insurers: this pass is reviewed one "
                         "insurer group at a time, and applying it blind "
                         "across all 33 is exactly the mistake to avoid.")
    ap.add_argument("--certs", default="",
                    help="comma-separated certification numbers: read only "
                         "these. The weekly workflow passes the plans it just "
                         "scraped, so a re-certified plan is read the same "
                         "week instead of falling back to '(see doc)'.")
    ap.add_argument("--correct-wrong", action="store_true",
                    help="also replace a scraped value that disagrees with "
                         "the schedule, not just a '(see doc)' placeholder. "
                         "Every replacement is printed.")
    args = ap.parse_args()
    if args.apply and not (args.insurers or args.certs):
        raise SystemExit("--apply requires --insurers or --certs "
                         "(refusing to apply to every insurer at once)")
    if args.apply and args.correct_wrong and args.certs:
        raise SystemExit("--correct-wrong is for a reviewed run, not the "
                         "unattended weekly one")

    with open(args.catalog, encoding="utf-8-sig") as f:
        catalog = list(csv.DictReader(f))

    prefixes = [p.strip() for p in args.insurers.split(",") if p.strip()]
    certs = {c.strip() for c in args.certs.split(",") if c.strip()}
    rows = [r for r in catalog
            if (not prefixes or any(r["insurer"].startswith(p) for p in prefixes))
            and (not certs or r["certification_no"] in certs)]

    out = []
    for r in rows:
        cert = r["certification_no"]
        text = pdf_text(args.cache, cert)
        if text is None:
            out.append(dict(certification_no=cert, insurer=r["insurer"],
                            plan_name=r["plan_name"], plan_level=r["plan_level_raw"],
                            ward="", ward_conf="pdf missing", ward_evidence="",
                            geography="", geo_conf="pdf missing", geo_evidence="",
                            lifetime="", lifetime_conf="pdf missing",
                            lifetime_evidence="", annual_limit="",
                            annual_conf="pdf missing", annual_evidence="",
                            coinsurance="", coins_conf="pdf missing",
                            coins_evidence="", ward_crosscheck=""))
            continue

        life, lconf, lcell = read_lifetime(text, r["currency"])
        annual, aconf, acell = read_annual(text, r["currency"])
        coins, cconf, ccell = read_coinsurance(text)

        wcells = find_cells(text, WARD_LABEL, ward_answers)
        gcells = find_cells(text, GEO_LABEL, geo_answers)

        if wcells:
            wcell = join_cell(wcells[0][1])
            ward, wconf = normalise_ward(wcells[0][1])
        else:
            wcell, ward, wconf = "", "", "no ward row in schedule"

        if gcells:
            gcell = gcells[0][2]
            geo, gconf = normalise_geo(gcell)
        else:
            gcell, geo, gconf = geo_from_prose(text)

        # Advisory only, never a gate. The supplement's adjustment table is a
        # genuinely independent naming of the ward class in some insurers'
        # documents (Sun Life, Blue Cross) and pure boilerplate in others:
        # Manulife prints the same full ladder at every benefit level, and
        # Bupa illustrates the rule with a fixed "Ward Room" example that does
        # not track the level. Bupa's own Section 3(b) defers to "the
        # restricted ward class ... specified in the Benefit Schedule", so
        # where the two differ the schedule governs and the table is noise.
        adj = adjustment_designated(text)
        if not adj:
            xcheck = "no adjustment table"
        elif len(adj) > 1:
            xcheck = f"no signal (table lists the full ladder: {sorted(adj)})"
        else:
            only = next(iter(adj))
            xcheck = ("agrees" if only in ward
                      else f"differs (table illustrates {only}; schedule governs)")

        out.append(dict(
            certification_no=cert, insurer=r["insurer"],
            plan_name=r["plan_name"], plan_level=r["plan_level_raw"],
            ward=ward, ward_conf=wconf, ward_evidence=wcell[:400],
            geography=geo, geo_conf=gconf, geo_evidence=gcell[:400],
            lifetime=life, lifetime_conf=lconf, lifetime_evidence=lcell[:400],
            annual_limit=annual, annual_conf=aconf, annual_evidence=acell[:400],
            coinsurance=coins, coins_conf=cconf, coins_evidence=ccell[:400],
            ward_crosscheck=xcheck))

    cols = ["certification_no", "insurer", "plan_name", "plan_level",
            "ward", "ward_conf", "ward_crosscheck", "ward_evidence",
            "geography", "geo_conf", "geo_evidence",
            "lifetime", "lifetime_conf", "lifetime_evidence",
            "annual_limit", "annual_conf", "annual_evidence",
            "coinsurance", "coins_conf", "coins_evidence"]
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for row in out:
            w.writerow(row)

    if args.apply:
        cat_by_cert = {r["certification_no"]: r for r in catalog}
        filled, skipped, corrected = apply(out, args.apply, cat_by_cert,
                                                args.benefits, args.correct_wrong)
        print(f"applied to {args.apply}")
        for k, n in sorted(filled.items()):
            print(f"    {n:4}  {k}")
        if corrected:
            print(f"    {len(corrected):4}  REPLACED a value that was already stated:")
            seen = set()
            for cert, field, cur, new in corrected:
                key = (field, cur, new)
                if key in seen:
                    continue
                seen.add(key)
                n = sum(1 for c in corrected if (c[1], c[2], c[3]) == key)
                print(f"          {field:9} {cur!r} -> {new!r}  ({n} plans)")
        if skipped:
            print(f"    {len(skipped):4}  left alone (already had a real value)")
            for cert, w, g in skipped[:5]:
                print(f"          {cert}  ward={w!r} geo={g!r}")
        return 0

    ok_w = sum(1 for r in out if r["ward"])
    ok_g = sum(1 for r in out if r["geography"])
    ok_l = sum(1 for r in out if r["lifetime"])
    print(f"{len(out)} plans -> {args.out}")
    print(f"  ward extracted:      {ok_w}/{len(out)}")
    print(f"  geography extracted: {ok_g}/{len(out)}")
    print(f"  lifetime extracted:  {ok_l}/{len(out)}")
    print(f"  annual extracted:    {sum(1 for r in out if r['annual_limit'])}/{len(out)}")
    print(f"  coinsurance read:    {sum(1 for r in out if r['coinsurance'])}/{len(out)}")
    import collections as _c
    print("  ward cross-check (advisory):")
    for k, n in _c.Counter(r["ward_crosscheck"].split(" (")[0]
                           for r in out if r["ward"]).most_common():
        print(f"    {n:4}  {k}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
