#!/usr/bin/env python3
"""
Extract citation instances from a zip file of Federal Reserve Tealbook PDFs.

Usage:
    python tealbook_zip_citation_pipeline.py /path/to/tealbooks.zip \
        --output tealbook_citations.xlsx

What it does
------------
- Unzips PDFs
- Parses filename metadata:
    FOMC20190130tealbooka20190118.pdf -> Date=20190130, Year=2019, Tealbook=A
- Extracts text page by page
- Detects citation instances in body text and footnote-like text
- Expands matches to fuller bibliographic citation text where possible
- Captures context = previous sentence + containing sentence + next sentence
- Infers section from page text
- Classifies Source_Type, Paper_Type, and Pub_Name
- Writes one row per citation instance to Excel

Notes
-----
This is a rule-based baseline designed to scale across many PDFs.
It will not catch every citation perfectly, especially layout-heavy pages,
boxes, or unusual footnotes. The cleanest production upgrade is to add an
LLM review step on candidate pages after this first-pass extraction.
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from dataclasses import asdict, dataclass, replace
from pathlib import Path
from typing import List

import fitz  # PyMuPDF
import pandas as pd


OUTPUT_COLUMNS = [
    "Date",
    "Year",
    "Tealbook",
    "Meeting_Number",
    "Citation",
    "Context",
    "Section",
    "Page",
    "Printed_Page",
    "Source_Type",
    "Paper_Type",
    "Pub_Name",
]

SECTION_PATTERNS = [
    "Domestic Economic Developments",
    "Financial Market Developments",
    "International Developments",
    "Monetary Policy Strategies",
    "Balance Sheet & Income",
    "Domestic Financial Developments",
    "International Financial Developments",
    "Staff Review of the Economic Situation",
    "Staff Review of the Financial Situation",
    "Economic Outlook",
    "Risks and Uncertainty",
    "Monetary Policy Alternatives",
]

# Newer Tealbooks (roughly 2010+) print a short section label in a rotated
# tab down the page margin (e.g. "Domestic Econ Devel & Outlook"). In
# PyMuPDF's default text-extraction reading order this rotated text
# consistently lands at the very end of the page, between the "Page N of
# NN" footer and the final "Authorized for Public Release" line -- see
# extract_sidebar_tag(). These are the exact tab strings confirmed against
# a real 2015 Tealbook; SIDEBAR_TAG_HINTS is a fallback for tabs not in this
# list (e.g. Book B / Bluebook-style policy tealbooks), matched as an exact
# standalone line anywhere on the page since they are distinctive enough
# not to collide with ordinary body subheadings.
KNOWN_SIDEBAR_TAGS = [
    "Domestic Econ Devel & Outlook",
    "Int\u2019l Econ Devel & Outlook",
    "Int'l Econ Devel & Outlook",
    "Financial Developments",
    "Risks & Uncertainty",
    "Greensheets",
    "Monetary Policy Strategies",
    "Policy Alternatives",
]

PAGE_FOOTER_RE = re.compile(r"Page\s+\d{1,3}\s+of\s+\d{1,3}")
AUTHORIZED_RELEASE_RE = re.compile(r"Authorized\s+for\s+Public\s+Release", re.IGNORECASE)

FED_RESEARCH_MARKERS = [
    "board of governors of the federal reserve system",
    "federal reserve board",
    "federal reserve bank of",
    "finance and economics discussion series",
    "feds",
    "ifdp",
    "feds notes",
    "frbny staff report",
    "frbny staff reports",
    "memorandum to the fomc",
]

GOV_IO_MARKERS = [
    "u.s. treasury",
    "congressional budget office",
    "cbo",
    "bank for international settlements",
    "bis",
    "ecb",
    "european central bank",
    "imf",
    "international monetary fund",
    "world bank",
    "oecd",
    "organization for economic cooperation and development",
]

PRIVATE_SECTOR_MARKERS = [
    "blue chip",
    "survey of professional forecasters",
    "goldman sachs",
    "morgan stanley",
    "j.p. morgan",
    "jp morgan",
    "bank of america",
    "barclays",
    "nomura",
    "ubs",
    "dealer outlook",
    "consensus forecast",
]

ACADEMIC_MARKERS = [
    "journal of",
    "review of",
    "quarterly journal of",
    "american economic review",
    "econometrica",
    "brookings papers",
    "university press",
    "nber",
    "cepr",
    "iza",
    "cesifo",
]

JOURNAL_PATTERNS = [
    r"\bJournal of [A-Z][A-Za-z&\- ,:;]+",
    r"\bReview of [A-Z][A-Za-z&\- ,:;]+",
    r"\bQuarterly Journal of [A-Z][A-Za-z&\- ,:;]+",
    r"\bAmerican Economic Review\b",
    r"\bEconometrica\b",
    r"\bBrookings Papers on Economic Activity\b",
]

PUBNAME_PATTERNS = [
    r"(Journal of [A-Za-z&\- ,:;]+)",
    r"(Review of [A-Za-z&\- ,:;]+)",
    r"(Quarterly Journal of [A-Za-z&\- ,:;]+)",
    r"(American Economic Review)",
    r"(Brookings Papers on Economic Activity)",
    r"(NBER Working Paper No\.? ?\d+)",
    r"(FEDS Notes)",
    r"(IFDP)",
    r"(FRBNY Staff Reports?)",
    r"(Finance and Economics Discussion Series)",
    r"(Policy Research Working Paper(?: No\.? ?\d+)?)",
]

CITATION_PATTERNS = [
    re.compile(r"\b[A-Z][A-Za-z'\-]+(?:,? [A-Z]\.)?(?: et al\.)? \((?:19|20)\d{2}[a-z]?\)"),
    re.compile(r"\b[A-Z][A-Za-z'\-]+ and [A-Z][A-Za-z'\-]+ \((?:19|20)\d{2}[a-z]?\)"),
    re.compile(r"\((?:[A-Z][A-Za-z'\-]+(?: et al\.)?|[A-Z][A-Za-z'\-]+ and [A-Z][A-Za-z'\-]+),? (?:19|20)\d{2}[a-z]?\)"),
    re.compile(r"\b(?:IMF|OECD|CBO|ECB|BIS|Blue Chip|Survey of Professional Forecasters) \((?:19|20)\d{2}[a-z]?\)"),
    re.compile(
        r"[A-Z][^\n]{0,250}?\((?:19|20)\d{2}[a-z]?\)[^\n]{0,500}?"
        r"(?:Journal of|Review of|Quarterly Journal|Working Paper|Staff Reports?|FEDS|IFDP|"
        r"IMF|OECD|Blue Chip|Brookings Papers|University Press|American Economic Review)"
        r"[^\n]{0,300}"
    ),
]

FOOTNOTE_START_RE = re.compile(r"(?:^|\n)\s*\d+\s+")
FOOTNOTE_BLOCK_RE = re.compile(r"(?:^|\n)\s*\d+\s+[A-Z].{20,}", re.DOTALL)


@dataclass
class CitationRow:
    Date: str
    Year: str
    Tealbook: str
    Meeting_Number: str
    Citation: str
    Context: str
    Section: str
    Page: str
    Printed_Page: str
    Source_Type: str
    Paper_Type: str
    Pub_Name: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract citation instances from a zip of Tealbook PDFs.")
    parser.add_argument("zip_path", help="Path to input zip file containing PDFs")
    parser.add_argument("--output", required=True, help="Output .xlsx file path")
    parser.add_argument(
        "--keep-extracted",
        action="store_true",
        help="Keep the extracted PDF directory instead of removing it afterward",
    )
    return parser.parse_args()


def parse_filename_metadata(filename: str):
    m_date = re.search(r"FOMC_(\d{4})-(\d{2})-(\d{2})", filename)
    if not m_date:
        return "", "", ""

    year, month, day = m_date.groups()
    date = f"{year}{month}{day}"

    tb = ""

    # New format
    m_tb = re.search(r"Tealbook_([AB])", filename, re.IGNORECASE)
    if m_tb:
        tb = m_tb.group(1).upper()

    # Old formats
    elif re.search(r"Greenbook", filename, re.IGNORECASE):
        tb = "A"
    elif re.search(r"Bluebook", filename, re.IGNORECASE):
        tb = "B"

    return date, year, tb


def normalize_whitespace(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def normalize_citation_whitespace(text: str) -> str:
    """
    More aggressive normalization for bibliography/citation extraction:
    converts single line breaks into spaces so wrapped journal titles
    are less likely to be cut off.
    """
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"(?<!\n)\n(?!\n)", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def split_sentences(text: str) -> List[str]:
    text = normalize_whitespace(text)
    if not text:
        return []
    parts = re.split(r"(?<=[\.\?!])\s+(?=(?:\(|\"|[A-Z]))", text)
    return [p.strip() for p in parts if p.strip()]


CONTEXT_WINDOW_WORDS = 30


def _nth_span(text: str, pattern: str, occurrence: int, flags: int = 0) -> "tuple[int, int] | None":
    """Return the (start, end) span of the `occurrence`-th (1-indexed) match
    of `pattern` in `text`. Falls back to the last available match if there
    are fewer matches than `occurrence` (keeps context extraction working
    even if occurrence counting is slightly off across text transformations,
    rather than returning nothing).
    """
    spans = [(m.start(), m.end()) for m in re.finditer(pattern, text, flags)]
    if not spans:
        return None
    idx = min(max(occurrence, 1), len(spans)) - 1
    return spans[idx]


def find_citation_anchor_span(text: str, citation: str, occurrence: int = 1) -> "tuple[int, int] | None":
    """Locate where a (possibly reconstructed) citation actually sits in the
    normalized page text, so real surrounding words can be pulled for the
    Context column.

    `occurrence` picks which repeated mention this is (1st, 2nd, 3rd, ...)
    -- important because the same short citation (e.g. "Taylor (1999)") can
    legitimately appear many times on one page, and each row's Context
    should reflect where THAT particular mention sits, not always the first.

    Full bibliographic citations (memos, working papers, journal articles)
    are rebuilt from fragments -- author list, quoted title, date -- and run
    through clean_citation_text(), so they rarely match the raw page text
    verbatim. Try several anchors, from most to least specific, so Context
    is never just the citation echoed back at itself.
    """
    # 1) Exact match. Works for most inline author-year citations, since
    # those are extracted directly from the normalized page text.
    if citation:
        span = _nth_span(text, re.escape(citation), occurrence)
        if span is not None:
            return span

    # 2) Quoted title, if present. Titles are copied verbatim from the
    # source (only whitespace-normalized), so this is the most reliable
    # anchor for reconstructed memo/note/working-paper/journal citations.
    # Match word-by-word with flexible whitespace so a line-wrapped title
    # in the source still lines up; strip trailing punctuation from each
    # word since the reconstructed citation attaches a comma/period to the
    # title that the original page text usually does not have.
    m = re.search(r'["\u201c](.+?)["\u201d]', citation)
    if m:
        title_words = [w.strip(" ,.;:") for w in re.findall(r"\S+", m.group(1))[:10]]
        title_words = [w for w in title_words if w]
        if title_words:
            pat = r"\s+".join(re.escape(w) for w in title_words)
            span = _nth_span(text, pat, occurrence, flags=re.IGNORECASE)
            if span is not None:
                return span

    # 3) Short author/year anchor, e.g. "Taylor (1999)" out of a longer
    # reconstructed citation, or a multi-year mention expanded from
    # "Taylor (1993, 1999)" into "Taylor (1993)".
    short_anchor = get_short_anchor(citation)
    if short_anchor and short_anchor != citation:
        span = _nth_span(text, re.escape(short_anchor), occurrence)
        if span is not None:
            return span

    # 4) First capitalized surname in the citation, as a last resort (covers
    # memo/report citations with no quoted title and no author-year form).
    m = re.search(r"[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+", citation)
    if m:
        span = _nth_span(text, re.escape(m.group(0)), occurrence)
        if span is not None:
            return span

    return None


def get_context(full_text: str, citation: str, occurrence: int = 1, window_words: int = CONTEXT_WINDOW_WORDS) -> str:
    """Return the citation plus `window_words` words of real surrounding
    text on each side (e.g. "See this memo ..."), so Context always shows
    what the citation is embedded in rather than repeating the citation.

    `occurrence` (1-indexed) selects which repeated mention of this citation
    on the page this row corresponds to, so five separate "Taylor (1999)"
    rows get five different, accurate Context values instead of the same
    first-occurrence text five times.

    Word-window is used instead of sentence-splitting because sentence
    boundaries are unreliable on PDF-extracted text (abbreviations, wrapped
    lines, footnote numerals), and a fixed word count is far more robust to
    locate around any anchor found by find_citation_anchor_span.
    """
    text = normalize_citation_whitespace(full_text)
    span = find_citation_anchor_span(text, citation, occurrence=occurrence)
    if span is None:
        # Nothing in the page text could be matched to this citation at all
        # (should be rare); fall back to returning the citation itself so the
        # column is never blank.
        return citation

    start, end = span
    before_words = re.findall(r"\S+", text[:start])[-window_words:]
    after_words = re.findall(r"\S+", text[end:])[:window_words]
    middle = text[start:end]
    return normalize_whitespace(" ".join(before_words + [middle] + after_words))


def extract_sidebar_tag(page_text: str) -> str:
    """Extract the rotated section tab printed in the page margin of newer
    Tealbooks (e.g. "Domestic Econ Devel & Outlook", "Risks & Uncertainty").

    This is what the reader actually sees as "the section" for a page, and
    is distinct from -- and should take priority over -- bolded body
    subheadings like "Monetary Policy", which are subsections within a
    tab, not the tab itself.

    Two strategies, since the tab's exact position in extracted-text
    reading order is not fully consistent across page layouts:
    1. Known exact tab strings (KNOWN_SIDEBAR_TAGS), matched as a standalone
       line anywhere on the page. This is what correctly finds the tab even
       on data-table-heavy Greensheets pages, where the tab text ends up
       buried in the middle of a long run of table numbers rather than
       cleanly isolated near a footer/header anchor.
    2. A structural fallback for tabs not in that list: on ordinary
       narrative pages the tab text lands, in PyMuPDF's reading order,
       between the "Page N of NN" footer and the final "Authorized for
       Public Release" line at the very end of the page. Used only when
       strategy 1 finds nothing, so it does not need to handle table pages.

    Returns "" if no tab is found (older Tealbooks without this feature, or
    a page where it could not be reliably isolated) so callers can fall
    back to the previous heuristics.
    """
    for tag in KNOWN_SIDEBAR_TAGS:
        if re.search(rf"(?m)^[ \t]*{re.escape(tag)}[ \t]*$", page_text):
            return tag

    page_m = PAGE_FOOTER_RE.search(page_text)
    if not page_m:
        return ""
    auth_after = [
        a for a in AUTHORIZED_RELEASE_RE.finditer(page_text) if a.start() > page_m.end()
    ]
    if not auth_after:
        return ""

    between = page_text[page_m.end():auth_after[0].start()]
    lines = [l.strip() for l in between.splitlines() if l.strip()]
    if not lines:
        return ""
    # Table-heavy pages can leave stray numeric/date fragments in this
    # window; only accept it if every line looks like a short label, not
    # data.
    if any(len(l) > 45 or re.search(r"\d", l) for l in lines):
        return ""
    candidate = re.sub(r"\s+", " ", " ".join(lines)).strip()
    if not candidate or len(candidate) > 45:
        return ""
    if not re.match(r"^[A-ZÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'\u2019&.,\-\s]*$", candidate):
        return ""

    # Collapse an accidental doubled tab, e.g. "Risks & Uncertainty Risks &
    # Uncertainty" (both the left- and right-margin tabs extracted next to
    # each other), into a single instance.
    words = candidate.split()
    half = len(words) // 2
    if half > 0 and words[:half] == words[half:]:
        candidate = " ".join(words[:half])
    return candidate.strip()


def find_section(page_text: str) -> str:
    sidebar = extract_sidebar_tag(page_text)
    if sidebar:
        return sidebar

    lower = page_text.lower()
    for section in SECTION_PATTERNS:
        if section.lower() in lower:
            return section
    top_lines = [line.strip() for line in page_text.splitlines()[:12] if line.strip()]
    for line in top_lines:
        if 5 <= len(line) <= 80 and line == line.title():
            return line
    return ""


def extract_printed_page_number(page: fitz.Page) -> str:
    """Best-effort extraction of the page number as it appears printed on
    the page, for reference only.

    The Page column itself is the physical position of the page within the
    PDF file (see process_pdf), which is always unambiguous and identical in
    meaning across every Tealbook vintage. Printed page numbering, by
    contrast, takes several different forms across vintages -- confirmed
    against real examples: "Page N of NN" (2015+ Tealbook A), a bare
    "N of NN" with no "Page" word (some Bluebooks), a lone centered number
    with no "of NN" at all (older Bluebooks), and "II-25"-style roman-
    section-plus-page numbers (Greenbooks) -- and some vintages print no
    page number at all on chart-only pages. Trying to out-guess all of
    that with coordinate scoring previously produced silently wrong
    results (including negative page numbers), so this function is
    intentionally conservative: it returns "" rather than a guess whenever
    the page doesn't match one of these known, confidently-recognized
    formats, since a blank value here is far safer than a wrong one.
    """
    raw_text = page.get_text("text") or ""

    m = re.search(r"Page\s+(\d{1,3})\s+of\s+\d{1,3}", raw_text, re.IGNORECASE)
    if m:
        return m.group(1)

    # Bare "N of NN" with no literal word "Page", isolated on its own line
    # (older Bluebooks). Requiring the whole line to be just this phrase
    # avoids matching ordinary prose like "3 of 4 respondents said...".
    m = re.search(r"(?m)^[ \t]*(\d{1,3})\s+of\s+\d{1,3}[ \t]*$", raw_text)
    if m:
        return m.group(1)

    # "II-25" / "I-3" style roman-numeral-section-plus-page (Greenbooks),
    # isolated on its own line for the same reason.
    m = re.search(r"(?m)^[ \t]*([IVXLC]{1,6}-\d{1,3})[ \t]*$", raw_text)
    if m:
        return m.group(1)

    return ""


def clean_citation_text(text: str) -> str:
    text = normalize_citation_whitespace(text)
    # Common PDF extraction artifacts. Keep this conservative: these should
    # only improve display, not change detection logic.
    text = re.sub(r"([A-Za-z])-\s+([A-Za-z])", r"\1-\2", text)
    text = re.sub(r"\s+", " ", text)
    # Remove citation lead-ins only when they occur at the very beginning of
    # an already-extracted candidate. This prevents returning "Source:" or
    # "See" as part of the citation while not changing body contexts.
    text = re.sub(r"(?i)^(?:source\.?|see(?: also)?|for example(?:,)?|e\.g\.(?:,)?|note\.?)\s*[:.]?\s+", "", text)
    text = text.strip(" ;,.")
    return text

YEAR_RE = r"(?:19|20)\d{2}[a-z]?"
REPEATED_AUTHOR_DASH_RE = re.compile(r"^[\-\u2010-\u2015\u2212_]{3,}\s*")


ACADEMIC_AUTHOR_OVERRIDES = [
    "Taylor",
    "Orphanides",
    "Woodford",
    "Svensson",
    "McCallum",
    "Goodfriend",
    "Clarida",
    "Galí",
    "Gali",
    "Gertler",
    "Cochrane",
]

def has_academic_author_override(citation: str) -> bool:
    return any(re.search(rf"\b{re.escape(author)}\b", citation) for author in ACADEMIC_AUTHOR_OVERRIDES)

KNOWN_INSTITUTIONAL_AUTHORS = [
    "Board of Governors of the Federal Reserve System",
    "Congressional Budget Office",
    "Office of Management and Budget",
    "Federal Reserve Bank of New York",
    "Federal Reserve Bank of Chicago",
    "Federal Reserve Bank of San Francisco",
    "Federal Reserve Bank of Boston",
    "Federal Reserve Bank of Cleveland",
    "Federal Reserve Bank of Philadelphia",
    "Federal Reserve Bank of Richmond",
    "Federal Reserve Bank of Atlanta",
    "Federal Reserve Bank of St. Louis",
    "Federal Reserve Bank of Dallas",
    "Federal Reserve Bank of Kansas City",
    "Federal Reserve Bank of Minneapolis",
    "U.S. Department of the Treasury",
    "Department of the Treasury",
    "Treasury Department",
]

def is_known_institutional_author(citation: str) -> bool:
    c = clean_citation_text(citation).lower()
    return any(c.startswith(author.lower() + " ") or c.startswith(author.lower() + "(") for author in KNOWN_INSTITUTIONAL_AUTHORS)

def extract_known_institutional_citations_with_spans(page_text: str) -> List[tuple[int, int, str]]:
    """Extract hard-coded agency/institution citations with full author names.

    These often appear as chart/table sources, e.g. "Source. Congressional
    Budget Office (2008), ...". The ordinary surname-year parser would turn
    that into "Office (2008)", so we capture these first and remove their
    spans from the inline pass.
    """
    text = normalize_citation_whitespace(page_text)
    out: List[tuple[int, int, str]] = []
    for author in KNOWN_INSTITUTIONAL_AUTHORS:
        a = re.escape(author)
        # Keep the pattern short and source-like; stop at the first terminal
        # period before obvious page/table spillover.
        pat = re.compile(
            rf"(?is)(?:Source\.?\s*)?(?P<cite>{a}\s*\({YEAR_RE}\)\s*,?\s*.{{0,450}}?)(?:(?<=\.)\s+(?=[A-Z][a-z]|Class\s+I\s+FOMC|$)|$)"
        )
        for m in pat.finditer(text):
            cite = clean_citation_text(m.group('cite'))
            cite = re.sub(r"(?is)\s+Class\s+I\s+FOMC.*$", "", cite).strip()
            cite = re.sub(r"(?is)\s+Supervisory and Regulatory Actions.*$", "", cite).strip()
            if re.search(rf"\({YEAR_RE}\)", cite):
                out.append((m.start(), m.end(), cite))
    return out


def extract_years(year_blob: str) -> List[str]:
    """Return all years inside a parenthetical year blob, e.g. '1993, 1999'."""
    return re.findall(YEAR_RE, year_blob)


def expand_inline_citation(citation: str) -> List[str]:
    """Expand compact in-text citations into one author-year citation per year.

    Examples:
        Taylor (1993, 1999) -> Taylor (1993), Taylor (1999)
        (Taylor, 1993, 1999) -> Taylor (1993), Taylor (1999)
        Ilzetzki and others (2010) -> Ilzetzki and others (2010)
    """
    text = normalize_citation_whitespace(citation).strip()
    author_tail = r"(?:\s+and\s+[A-Z][A-Za-z'\-]+|\s+et al\.|\s+and others)?"

    m = re.fullmatch(
        rf"(?P<author>[A-Z][A-Za-z'\-]+{author_tail})\s*\((?P<years>{YEAR_RE}(?:\s*[,;]\s*{YEAR_RE})*)\)",
        text,
    )
    if m:
        return [f"{m.group('author')} ({year})" for year in extract_years(m.group('years'))]

    m = re.fullmatch(
        rf"\((?P<author>[A-Z][A-Za-z'\-]+{author_tail}),\s*(?P<years>{YEAR_RE}(?:\s*[,;]\s*{YEAR_RE})*)\)",
        text,
    )
    if m:
        return [f"{m.group('author')} ({year})" for year in extract_years(m.group('years'))]

    return [text]


def is_reference_page(page_text: str) -> bool:
    return bool(re.search(r"(?im)^\s*References\s*:?\s*$", page_text))


def find_reference_section_start(page_text: str) -> "int | None":
    """Character offset where the page's 'References' heading line begins,
    or None if there isn't one. Used to scope inline-citation scanning to
    the body text above the References section, so a page that has both
    body-text inline mentions and a full bibliography at the bottom (common
    in Tealbook appendices) doesn't lose the inline mentions to the
    reference-list branch in detect_citations().
    """
    m = re.search(r"(?im)^\s*References\s*:?\s*$", page_text)
    return m.start() if m else None


def looks_like_reference_start(line: str) -> bool:
    line = line.strip()
    if not line:
        return False
    if REPEATED_AUTHOR_DASH_RE.match(line) and re.search(rf"\({YEAR_RE}\)", line):
        return True
    # Typical author-date bibliography entry: "Taylor, John B. (1993) ..."
    return bool(re.match(rf"^[A-Z][A-Za-z'\-]+(?:,|\s+and\s+|\s+et al\.).*?\({YEAR_RE}\)", line))


# A looser "this line opens a new author list" shape than
# looks_like_reference_start -- "Surname, " at the very start of the line --
# with no requirement that "(YEAR)" also appear on this same line. Used only
# as a fallback (see split_reference_entries below) for entries whose author
# list is long enough to wrap onto a second line before the year shows up,
# e.g. a 9-author memo citation. looks_like_reference_start alone would
# never fire on either line in that case (no year on line 1, and line 2
# typically starts mid-list with a first name like "Stephen Meyer, ..." not
# "Surname, "), so without this fallback the whole entry is silently lost
# whenever it isn't already being accumulated as a continuation of a prior
# entry -- most obviously the very first entry in a References section.
AUTHOR_LIST_CONTINUATION_START_RE = re.compile(r"^[A-Z][A-Za-z'\-]+,\s+[A-Z]")


def get_reference_author_prefix(entry: str) -> str:
    m = re.match(rf"^(.+?\({YEAR_RE}\))", entry)
    if not m:
        return ""
    before_year = entry[: entry.find("(")].strip()
    return before_year


def split_reference_entries(page_text: str) -> List[str]:
    """Extract structured bibliography entries from a References section.

    This handles wrapped entries and repeated-author dash entries such as
    '----- (1999)' by inheriting the previous author string.
    """
    lines = page_text.splitlines()
    start_idx = None
    for i, line in enumerate(lines):
        if re.match(r"^\s*References\s*:?\s*$", line, re.IGNORECASE):
            start_idx = i + 1
            break
    if start_idx is None:
        return []

    entries: List[str] = []
    current: List[str] = []
    previous_author = ""

    def flush_current():
        nonlocal current, previous_author
        if not current:
            return
        entry = clean_citation_text(" ".join(current))
        entry = re.sub(r"\s+", " ", entry)
        if entry and re.search(rf"\({YEAR_RE}\)", entry):
            entries.append(entry)
            author = get_reference_author_prefix(entry)
            if author and not REPEATED_AUTHOR_DASH_RE.match(author):
                previous_author = author
        current = []

    for raw_line in lines[start_idx:]:
        line = raw_line.strip()
        if not line:
            flush_current()
            continue
        if re.match(r"^Class\s+I\s+FOMC\b", line):
            flush_current()
            break

        if looks_like_reference_start(line):
            flush_current()
            if REPEATED_AUTHOR_DASH_RE.match(line) and previous_author:
                line = REPEATED_AUTHOR_DASH_RE.sub(previous_author + " ", line)
            current = [line]
        elif current:
            current.append(line)
        elif AUTHOR_LIST_CONTINUATION_START_RE.match(line):
            # No entry currently open, and this doesn't look like a
            # complete one-line reference start either -- but it does open
            # a "Surname, " author list, so start accumulating rather than
            # dropping it. flush_current()'s own year-present check still
            # guards against keeping this if a "(YEAR)" never actually
            # turns up on a later wrapped line.
            current = [line]

    flush_current()
    return entries


# NOTE_WORD covers memo, memorandum, and note -- including "staff note" and
# "briefing note", since the \b before the group anchors on the word "note"
# itself regardless of the descriptive word in front of it (e.g. "staff",
# "briefing"). It intentionally does NOT match bare "Note:" chart captions or
# "note 4" cross-references, because every branch below still requires the
# structural marker (to the Committee/FOMC, or by <Capitalized Author>) right
# next to it.
NOTE_WORD = r"(?:memo(?:randum)?|note)"

# Shared "who the memo was sent to" fragment. Most memos go "to the FOMC"/
# "to the Committee", but some go directly "to the Board of Governors"
# instead (e.g. staff research memos addressed to the Board rather than
# routed through the FOMC).
MEMO_BODY_RE = r"(?:Federal\s+Open\s+Market\s+Committee|FOMC|Committee|Board\s+of\s+Governors(?:\s+of\s+the\s+Federal\s+Reserve\s+System)?)"

STRICT_MEMO_CITATION_RE = re.compile(
    r"(?:"
    rf"{NOTE_WORD}\s+to\s+(?:the\s+)?{MEMO_BODY_RE}"
    rf"|{NOTE_WORD}\s+by\s+[A-ZÀ-ÖØ-Þ]"
    rf"|(?:[A-Z][a-z]+\s+\d{{1,2}},\s+)?(?:19|20)\d{{2}}[a-z]?\s+{NOTE_WORD}\s+by\s+[A-ZÀ-ÖØ-Þ]"
    r"|Board\s+of\s+Governors\s+of\s+the\s+Federal\s+Reserve\s+System"
    r"|Division\s+of\s+Research\s+and\s+Statistics"
    r")",
    re.IGNORECASE,
)


def has_strict_memo_signal(text: str) -> bool:
    """True for actual memo citations, not generic phrases like 'memorandum items'."""
    return bool(STRICT_MEMO_CITATION_RE.search(text))





def has_title_first_memo_signal(text: str) -> bool:
    """True for title-first memo/note citations, including date-only forms.

    Requires the full structure: (memo|memorandum|note|staff note|briefing
    note) -> "Title" -> by AUTHORS -> (distributed|sent) to the FOMC/Committee.
    This is what distinguishes a real cited internal document from a stray
    sentence like "Note that inflation..." or "See note 4" -- those never
    have a quoted title followed by "by <Author>" followed by a
    distributed/sent-to-Committee marker, so they never match.

    Example: See the note on "Title" by Author that was distributed to the FOMC on March 11.
    """
    return bool(re.search(
        rf"(?is)\b{NOTE_WORD}\s+(?:entitled\s+|on\s+)?[\"“].+?[\"”]\s+by\s+[A-ZÀ-ÖØ-Þ].{{0,320}}?\b(?:that\s+was\s+)?(?:distributed|sent)\s+to\s+(?:the\s+)?(?:FOMC|Federal\s+Open\s+Market\s+Committee|Committee)\b",
        text,
    ))

def starts_like_author_citation(citation: str) -> bool:
    """True when a long footnote/reference candidate begins with an author list.

    This prevents prose such as "This extension incorporates... the memo to the
    Committee by Robert Tetlow..." from being accepted as the citation.
    The actual memo is extracted separately by an author-anchored pattern.
    """
    citation = clean_citation_text(citation)
    if not citation:
        return False

    # Bibliography style: King, T. B.; Levin, A. T.; and Perli, R., "Title"...
    if re.match(rf"^{INITIAL_AUTHOR_LIST_RE}\s*,?\s*[\"“]", citation):
        return True

    # Last, First style: Taylor, John B. (1993) ...
    if re.match(r"^[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+\s*,\s*[A-ZÀ-ÖØ-Þ]", citation):
        return True

    # Full-name style: Robert Tetlow, "Title"... or A. Thomas King and Jane Doe (2007)...
    full_name = rf"{PERSON_NAME_RE}(?:\s*,\s*{PERSON_NAME_RE})*(?:\s*,?\s+and\s+{PERSON_NAME_RE})?"
    if re.match(rf"^{full_name}\s*,\s*[\"“]", citation):
        return True
    if re.match(rf"^{full_name}\s*\({YEAR_RE}\)", citation):
        return True

    return False

def is_full_bibliographic_citation(citation: str) -> bool:
    """Return True only for full bibliographic-style citations.

    This avoids two recurring false positives: chart/table notes that contain
    source-like words, and ordinary prose containing the standalone word
    'memorandum'. A memo is treated as a full citation only when it has a
    real citation marker such as 'memo by ...', 'memorandum to the FOMC',
    or Board/Fed institutional source language.
    """
    citation = clean_citation_text(citation)
    c = citation.lower()

    if not citation:
        return False

    # Hard-coded institutional authors are valid agency/government citations
    # even though the ordinary surname-year test would see only the final word
    # of the organization name (for example, Office (2008)).
    if is_known_institutional_author(citation) and re.search(rf"\({YEAR_RE}\)", citation):
        return True

    has_signal = bool(FULL_CITATION_SIGNAL_RE.search(citation) or has_strict_memo_signal(citation))
    if not has_signal:
        return False

    # Obvious chart/table/page-note spillover, not bibliography.
    if re.search(r"\b(ratio scale|blue shaded|class i fomc|page \d+ of \d+|nominal trade-weighted|real major currencies|journal of commerce index)\b", c):
        return False
    if re.match(r"(?i)^(?:percent|ratio|index|monthly|quarterly|note\.|memo\s+taylor|near-term prescriptions|inflation objective memo)\b", citation):
        return False

    # Do not accept prose merely because it contains words such as
    # 'Call Report memorandum items'.
    if re.search(r"(?i)\bmemorandum\b", citation) and not has_strict_memo_signal(citation):
        return False

    paren_year = re.search(rf"\({YEAR_RE}\)", citation)
    if paren_year:
        prefix = citation[:paren_year.start()].strip(" ,;.")
        suffix = citation[paren_year.end():].strip()

        # The material immediately before the year should be an author list,
        # allowing both "First Last" and "Last, First" bibliography styles.
        prefix_tail = prefix[-320:]
        has_author_before_year = bool(
            AUTHOR_LIST_BEFORE_YEAR_RE.search(prefix_tail)
            or re.search(r"[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+\s*,\s*[A-ZÀ-ÖØ-Þ]", prefix_tail)
        )
        if not has_author_before_year:
            return False

        # Full citations normally have a title/source marker after the year.
        # Plain rule labels like "Taylor (1993) rule" should remain inline citations.
        has_full_after_year = bool(
            re.search(r"^[,\.]", suffix)
            or re.search(r'["“”]', suffix)
            or re.search(r"(?i)journal|review|working paper|unpublished paper|discussion paper|staff report|feds|ifdp|brookings|university press|conference series|economic letter|nber|pp\.|vol\.|forthcoming|www\.|http|memorandum to|memo to", suffix)
            or has_strict_memo_signal(suffix)
        )
        return has_full_after_year

    # Non-parenthetical-year full citations: require a quoted title plus a
    # source/date marker, and require the extracted string to START with the
    # author list. Author-anchored memo/working-paper patterns construct this
    # clean form; generic surrounding prose should be rejected.
    return bool(
        starts_like_author_citation(citation)
        and re.search(YEAR_RE, citation)
        and re.search(r'["“].+?["”]', citation)
        and (
            re.search(r"(?i)working paper|unpublished paper", citation)
            or has_strict_memo_signal(citation)
        )
    )


FULL_CITATION_SIGNAL_RE = re.compile(
    r"journal|review|quarterly journal|working paper|unpublished paper|discussion paper|staff report|"
    r"feds|feds note|feds notes|ifdp|brookings|university press|conference series|"
    r"monetary policy rules|federal reserve bank|economic letter|dallas fed|nber|"
    r"federal open market committee|pp\.|vol\.|forthcoming|www\.|http",
    re.IGNORECASE,
)

# Unicode-aware name patterns for trimming prose before long footnote citations.
# They are intentionally broad because Tealbook footnotes include accents and
# full first-name author strings.
NAME_TOKEN_RE = r"(?:[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+|[A-Z]\.)"
PERSON_NAME_RE = rf"{NAME_TOKEN_RE}(?:\s+{NAME_TOKEN_RE}){{0,5}}"
AUTHOR_LIST_BEFORE_YEAR_RE = re.compile(
    rf"(?P<authors>{PERSON_NAME_RE}(?:\s*,\s*{PERSON_NAME_RE})*(?:\s*,?\s+and\s+{PERSON_NAME_RE})?)\s*$"
)

def trim_to_author_start_before_year(candidate: str) -> str:
    """If prose precedes an author list, start at the author list."""
    candidate = clean_citation_text(candidate)
    year_match = re.search(rf"\({YEAR_RE}\)", candidate)
    if not year_match:
        return trim_to_author_start_for_nonparen_year(candidate)

    prefix = candidate[:year_match.start()]
    suffix = candidate[year_match.start():]

    # Remove everything through the last explicit citation cue.
    cue_matches = list(re.finditer(
        r"(?i)(?:^|[\.;:,]\s+)(?:among[^,;:.]{0,120},\s+)?(?:for example,?\s+|for instance,?\s+)?(?:see(?: also)?|e\.g\.|such as|including)\s+",
        prefix,
    ))
    if cue_matches:
        prefix = prefix[cue_matches[-1].end():]

    # Also handle phrases like "based on the results in Jack Hadley..."
    in_matches = list(re.finditer(r"(?i)\b(?:in|by)\s+(?=[A-ZÀ-ÖØ-Þ])", prefix))
    if in_matches:
        prefix = prefix[in_matches[-1].end():]

    # Bibliography-style author lists sometimes end with the cited author's
    # last name before the year, e.g.
    #   Reifschneider, David and John C. Williams (2000)
    # AUTHOR_LIST_BEFORE_YEAR_RE (PERSON_NAME_RE already permits a
    # single-token name) handles this correctly, along with ordinary full
    # "First Last, First Last, and First Last" lists -- checked first, so a
    # second author's surname doesn't get mistaken for the start of an
    # inverted "Lastname, Firstname" entry and swallow the first author's
    # given name (e.g. "Stefania D'Amico, Don Kim, and Min Wei" incorrectly
    # trimmed down to "D'Amico, Don Kim, and Min Wei").
    m = AUTHOR_LIST_BEFORE_YEAR_RE.search(prefix.strip())
    if m:
        prefix = m.group('authors')
    else:
        bib_author = re.search(
            rf"(?P<authors>[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+\s*,\s*{PERSON_NAME_RE}(?:\s*,\s*{PERSON_NAME_RE})*(?:\s*,?\s+and\s+{PERSON_NAME_RE})?)\s*$",
            prefix.strip(),
        )
        if bib_author:
            prefix = bib_author.group('authors')

    return clean_citation_text(prefix.strip() + " " + suffix.lstrip())


def trim_to_author_start_for_nonparen_year(candidate: str) -> str:
    """Trim prose before full citations whose year appears at the end, not in parentheses.

    Example: "See King, T. B.; Levin, A. T; and Perli, R., ... Working Paper, July 2007"
    should start at "King, T. B.; ...".
    """
    candidate = clean_citation_text(candidate)
    cue_matches = list(re.finditer(
        r"(?i)(?:^|[\.;:,]\s+)(?:for further (?:details|information)[^,;:.]{0,120},\s+)?(?:see(?: also)?|e\.g\.|such as|including)\s+",
        candidate,
    ))
    if cue_matches:
        candidate = candidate[cue_matches[-1].end():]

    # "memo by Benson Durham" -> start at the author after "by".
    by_matches = list(re.finditer(r"(?i)\bby\s+(?=[A-ZÀ-ÖØ-Þ])", candidate[:220]))
    if by_matches and re.search(r"(?i)memo", candidate):
        candidate = candidate[by_matches[-1].end():]

    # Drop common lead-ins left at the beginning.
    candidate = re.sub(r"(?i)^(?:for further (?:details|information)[^,;:.]{0,120},\s+)?", "", candidate).strip()
    return clean_citation_text(candidate)


LEADING_CITATION_CUE_RE = re.compile(
    r"(?is)^.*?(?:\bsee\b|\bsee also\b|\bfor example,?\s+see\b|\bfor instance,?\s+see\b|\be\.g\.,?|\bsuch as\b|\bincluding\b)\s+",
)


def strip_leading_footnote_prose(candidate: str) -> str:
    """Remove explanatory prose before the actual bibliographic citation.

    Footnotes often embed a full citation inside a prose sentence.  This
    function trims to the author list immediately before the first year, so
    rows begin with the author names rather than phrases like "Among recent
    works, see" or "This assumption is based on the results in".
    """
    candidate = trim_to_author_start_before_year(candidate)

    # Drop common lead-in phrases left at the beginning as a final cleanup.
    candidate = re.sub(
        r"(?i)^(?:among[^,;:.]{0,120},\s+)?(?:for example,?\s+|for instance,?\s+)?(?:see(?: also)?|e\.g\.|such as|including)\s+",
        "",
        candidate,
    ).strip()

    return clean_citation_text(candidate)


def trim_full_citation_tail(candidate: str) -> str:
    """Trim prose/figure spillover around a long footnote citation."""
    candidate = strip_leading_footnote_prose(candidate)

    # Cut off obvious non-citation material that PyMuPDF sometimes appends
    # from nearby figures/tables after a footnote citation.
    cut_patterns = [
        r"\.?\s*\)?\s*End\s+footnote\s+\d+.*$",
        r"\.?\s*\)?\s*\[End\s+footnote\s+\d+\.?\].*$",
        r"\s+Gap<.*$",
        r"\s+Incentive\s+to\s+refinance.*$",
        r"\s+Figure\s+\d+\b.*$",
        r"\s+Chart\s+\d+\b.*$",
        r"\s+Table\s+\d+\b.*$",
        r"\.\s+(?:Domestic|International|Financial|Monetary)\s+(?:Econ|Economic|Developments|Policy).*$",
        r"\s+\d+(?:\.\d+)?\s+\d+(?:\.\d+)?\s+\d+(?:\.\d+)?\s+.*$",
        r"\s+\d+\s+(?=(?:See|For example|For instance|[A-Z][a-z]+\s+[A-Z][a-z]+\s+and\s+))",
    ]
    for pat in cut_patterns:
        candidate = re.sub(pat, "", candidate).strip()

    # Keep Fed memo citations through the division name before applying generic date endings.
    if re.search(r"(?i)memorandum|memo", candidate) and "Division of Research and Statistics" in candidate:
        return clean_citation_text(candidate)

    # Prefer clean endings when a recognizable bibliographic terminus exists.
    ending_patterns = [
        r"^(.+?https?://\S+?\.pdf)\b.*$",
        r"^(.+?https?://.+?\.pdf)\b.*$",
        r"^(.+?www\.\S+?\.pdf)\b.*$",
        r"^(.+?www\..+?\.pdf)\b.*$",
        r"^(.+?(?:https?://|www\.)\S+?)(?:\.\s+[A-Z].*)?$",
        r"^(.+?pp\.\s*\d+\s*(?:[-–—−]|\s+)\s*\d+\)?)(?:\.?(?:\s+.*)?)?$",
        r"^(.+?forthcoming\.)\s+.*$",
        r"^(.+?FEDS Notes?,\s+[A-Z][a-z]+\s+\d{1,2}\.)\s+.*$",
        r"^(.+?Economic Letter,\s+vol\.\s*\d+\s*\([^)]*\).*?(?:\.pdf|\.))\s+.*$",
        r"^(.+?Working Paper,\s*(?:[A-Z][a-z]+\s+)?(?:\d{1,2},\s*)?\d{4}\.?)\s+.*$",
        r"^(.+?unpublished paper,\s*[^.]*?\d{4}\.?)\s+.*$",
        # A trailing "Month Day[, Year]" right after the institution name is
        # itself the citation's date, not trailing junk -- must be tried
        # before the more generic memo patterns below, which would
        # otherwise treat the day number as strippable filler when there's
        # no terminating period (e.g. this citation was trimmed at a
        # "; and" sibling-citation boundary rather than a sentence end).
        r"^(.+?(?:Board of Governors of the Federal Reserve System|Federal Open Market Committee)[^.]*?,\s*[A-Z][a-z]+\s+\d{1,2}(?:,\s*\d{4})?)\.?\s*.*$",
        r"^(.+?memo(?:randum)?\s+to\s+the\s+Federal\s+Open\s+Market\s+Committee.*?(?:Division of Research and Statistics|Board of Governors of the Federal Reserve System)[^.]*\.?)\s+.*$",
        r"^(.+?memo(?:randum)?[^.]*?(?:Board of Governors of the Federal Reserve System|Federal Open Market Committee)[^.]*\.?)\s+.*$",
        r"^(.+?memo(?:randum)?[^.]*?,\s*(?:[A-Z][a-z]+\s+\d{1,2},\s*)?\d{4}\.?)\s+.*$",
    ]
    for pat in ending_patterns:
        m = re.match(pat, candidate, flags=re.IGNORECASE)
        if m:
            candidate = m.group(1).strip()
            break

    return clean_citation_text(candidate)




INITIAL_AUTHOR_LIST_RE = (
    r"[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+,\s*(?:[A-Z]\.?\s*){1,4}"
    r"(?:\s*;\s*[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+,?\s*(?:[A-Z]\.?\s*){1,4})*"
    r"(?:\s*;?\s*and\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+,?\s*(?:[A-Z]\.?\s*){1,4})?"
)

def strip_leading_see_cue(text: str) -> str:
    """Remove prose before the actual citation when it is introduced by See/etc."""
    text = clean_citation_text(text)
    text = re.sub(r"(?is)^.*?\b(?:see|see also|for further details,? see|for further information(?: regarding [^,]+)?,? see)\s+", "", text)
    return clean_citation_text(text)



def normalize_memo_date(date_text: str, default_year: str = "") -> str:
    """Normalize memo dates and optionally fill in a missing year.

    Examples:
        March 9 2009 -> March 9, 2009
        March 11 with default_year=2009 -> March 11, 2009
    """
    date_text = re.sub(r"\s+", " ", date_text.strip().strip(" .,;"))
    date_text = re.sub(rf"([A-Z][a-z]+\s+\d{{1,2}})\s+({YEAR_RE})\b", r"\1, \2", date_text)
    if default_year and not re.search(YEAR_RE, date_text):
        date_text = f"{date_text}, {default_year}"
    return date_text


def normalize_memo_body(body_text: str) -> str:
    """Normalize a matched recipient body (see MEMO_BODY_RE) to one of a
    small set of canonical forms, so the same institution doesn't produce
    several differently-worded Citation strings for the same underlying
    memo (e.g. "FOMC" vs "the Federal Open Market Committee").
    """
    b = re.sub(r"\s+", " ", body_text.strip()).lower()
    if b == "fomc":
        return "FOMC"
    if "board of governors" in b:
        return "Board of Governors of the Federal Reserve System"
    return "Federal Open Market Committee"

def build_memo_citation(authors: "str | None", title: str, body: str, date: "str | None") -> str:
    """Build a memo citation string, with author and date both optional.

    The structural signal that makes a match trustworthy is the
    combination of a quoted title, the memo/note word, and an explicit
    institutional recipient (FOMC/Committee/Board) -- not the presence of
    an author or date, which real Tealbook prose sometimes omits on
    repeat/informal references to a memo already introduced earlier.
    Rather than drop those instances, this omits only the piece that's
    actually missing so the citation is never fabricated, just
    incomplete.
    """
    title = title.strip().strip(" ,.")
    prefix = f'{clean_citation_text(authors)}, ' if authors else ""
    suffix = f", {date}" if date else ""
    return f'{prefix}"{title}," memo to {body}{suffix}'


def extract_nonparen_full_citations(text: str, default_year: str = "") -> List[str]:
    """Extract full working-paper and memo citations without parenthetical years."""
    out: List[str] = []

    # Working paper / unpublished paper with date year at the end.
    # This covers both full-name authors and initial-list staff citations such as
    # "King, T. B.; Levin, A. T; and Perli, R., "..." Working paper, July 2007."
    wp_pat = re.compile(
        rf"(?is)(?:^|\bSee\s+|for further (?:details|information)[^,.]{{0,120}},?\s+see\s+)"
        rf"(?P<cite>(?:{INITIAL_AUTHOR_LIST_RE}|[A-ZÀ-ÖØ-Þ][^\n;]{{0,320}}?)\s*,?\s*[\"“].+?[\"”]\.?"
        rf"\s*(?:Working\s+paper|Working\s+Paper|unpublished\s+paper),\s*(?:[A-Z][a-z]+\s+)?(?:\d{{1,2}},\s*)?{YEAR_RE}\.?)"
    )
    for m in wp_pat.finditer(text):
        c = strip_leading_see_cue(m.group('cite'))
        c = trim_full_citation_tail(c)
        if is_full_bibliographic_citation(c):
            out.append(c)

    # "memo/note by AUTHOR, \"TITLE\"" with a date before the memo/note word.
    memo_by_pat = re.compile(
        rf"(?is)(?P<date>[A-Z][a-z]+\s+\d{{1,2}},\s+{YEAR_RE})\s+{NOTE_WORD}\s+by\s+"
        rf"(?P<authors>[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+(?:\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+)*)"
        rf"\s*,\s*[\"“](?P<title>.+?)[\"”]"
    )
    for m in memo_by_pat.finditer(text):
        title = m.group('title').strip().strip(" ,.")
        c = f"{m.group('authors')}, \"{title},\" memorandum by {m.group('authors')}, {m.group('date')}"
        if is_full_bibliographic_citation(c):
            out.append(clean_citation_text(c))

    # "memo/note to the Committee by AUTHORS, \"TITLE,\" DATE"
    memo_to_pat = re.compile(
        rf"(?is){NOTE_WORD}\s+to\s+the\s+(?P<body>Committee|Federal Open Market Committee)\s+by\s+"
        rf"(?P<authors>[A-ZÀ-ÖØ-Þ][^\"“”]{{0,260}}?)\s*,\s*[\"“](?P<title>.+?)[\"”]\s*,?\s*"
        rf"(?P<date>[A-Z][a-z]+\s+\d{{1,2}},\s+{YEAR_RE})"
    )
    for m in memo_to_pat.finditer(text):
        body = "Federal Open Market Committee" if m.group('body').lower() == 'committee' else m.group('body')
        authors = clean_citation_text(m.group('authors'))
        title = m.group('title').strip().strip(" ,.")
        c = f"{authors}, \"{title},\" memorandum to the {body}, {m.group('date')}"
        if is_full_bibliographic_citation(c):
            out.append(clean_citation_text(c))

    # Title-first memo citations, e.g.
    #   The memo "Further Information ..." by Eileen Mauskopf and David Reifschneider,
    #   sent to FOMC, March 9 2009, ...
    # These otherwise get missed because the authors appear after the title.
    title_first_memo_pat = re.compile(
        rf"(?is)(?:^|\b(?:The|the|See|see)\s+)?"
        rf"{NOTE_WORD}\s+(?:entitled\s+|on\s+)?[\"“](?P<title>.+?)[\"”]"
        rf"\s+by\s+(?P<authors>[A-ZÀ-ÖØ-Þ][^\"“”.;]{{0,260}}?)"
        rf"\s*,?\s*(?:that\s+was\s+)?(?:sent|distributed)\s+to\s+(?:the\s+)?(?P<body>FOMC|Federal\s+Open\s+Market\s+Committee|Committee)"
        rf"(?:\s+on)?\s*,?\s*(?P<date>[A-Z][a-z]+\s+\d{{1,2}}(?:,?\s+{YEAR_RE})?)"
    )
    for m in title_first_memo_pat.finditer(text):
        authors = clean_citation_text(m.group('authors'))
        title = m.group('title').strip().strip(" ,.")
        body_raw = m.group('body') or 'FOMC'
        body = 'FOMC' if body_raw.upper() == 'FOMC' else 'Federal Open Market Committee'
        date = normalize_memo_date(m.group('date'), default_year=default_year)
        c = f'{authors}, "{title}," memo to {body}, {date}'
        if is_full_bibliographic_citation(c):
            out.append(clean_citation_text(c))

    # Authorless version of title_first_memo_pat: bare "memo \"TITLE,\"
    # distributed/sent to (the)? BODY on DATE" -- no "titled"/"entitled"
    # connector word (so memo_titled_pat doesn't match) and no "by
    # AUTHORS" clause at all. Example:
    #   the memo "Reinvestment Proposal," distributed to the FOMC on
    #   April 21, 2017.
    memo_bare_title_nodate_author_pat = re.compile(
        rf"(?is)(?:^|\b(?:The|the|See|see)\s+)?"
        rf"{NOTE_WORD}\s+[\"“](?P<title>.+?)[\"”]"
        rf"\s*,?\s*(?:that\s+was\s+)?(?:sent|distributed)\s+to\s+(?:the\s+)?(?P<body>{MEMO_BODY_RE})"
        rf"(?:\s+on)?\s*,?\s*(?P<date>[A-Z][a-z]+\s+\d{{1,2}}(?:,?\s+{YEAR_RE})?)"
    )
    for m in memo_bare_title_nodate_author_pat.finditer(text):
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        date = normalize_memo_date(m.group('date'), default_year=default_year)
        c = build_memo_citation(None, title, body, date)
        out.append(clean_citation_text(c))

    # "memo/note to the FOMC, \"TITLE\" by AUTHORS (DATE)" -- title comes
    # right after the body (FOMC/Committee), before the authors, with the
    # date in parentheses at the end. Distinct from memo_to_pat above,
    # where the authors come before the title instead. Example:
    #   see the memo to the FOMC, "Large-scale Asset Purchases and
    #   Inflation Expectations in the FRB/US Model" by David
    #   Reifschneider and John Roberts (April 20, 2009).
    memo_to_title_first_pat = re.compile(
        rf"(?is){NOTE_WORD}\s+to\s+the\s+(?P<body>FOMC|Committee|Federal\s+Open\s+Market\s+Committee)\s*,\s*"
        rf"[\"“](?P<title>.+?)[\"”]\s+by\s+"
        rf"(?P<authors>[A-ZÀ-ÖØ-Þ][^\"“”.;()]{{0,260}}?)"
        rf"\s*\(\s*(?P<date>[A-Z][a-z]+\s+\d{{1,2}},\s+{YEAR_RE})\s*\)"
    )
    for m in memo_to_title_first_pat.finditer(text):
        body_raw = m.group('body')
        body = "FOMC" if body_raw.upper() == "FOMC" else "Federal Open Market Committee"
        authors = clean_citation_text(m.group('authors'))
        title = m.group('title').strip().strip(" ,.")
        c = f'{authors}, "{title}," memo to {body}, {m.group("date")}'
        if is_full_bibliographic_citation(c):
            out.append(clean_citation_text(c))

    # "memo (by AUTHORS)? titled/entitled \"TITLE\" (by AUTHORS)? (sent|
    # distributed) to (the)? BODY (on)? DATE" -- "titled"/"entitled" as the
    # title connector, with authors either before or after the title.
    # Examples:
    #   the memo by Christopher Erceg, ... titled "An Overview of Simple
    #   Policy Rules ...," sent to the Committee on July 18, 2012.
    #   see the memo titled "Options for Continuation of Open-Ended Asset
    #   Purchases in 2013" by Board and FRBNY staff sent to the Committee
    #   on November 30, 2012
    #   see the memo entitled "Exit Strategy Considerations" (by K. Femia
    #   and J. Remache ...) that was sent to the Committee (corrected) on
    #   March 12, 2013.
    memo_titled_pat = re.compile(
        rf"(?is){NOTE_WORD}\s+"
        rf"(?:by\s+(?P<authors_pre>[A-Za-zÀ-ÖØ-öø-ÿ][^\"“”;()]{{0,260}}?)\s+)?"
        rf"(?:titled|entitled)\s+[\"“](?P<title>.+?)[\"”]"
        rf"(?:\s*\(?\s*by\s+(?P<authors_post>[A-Za-zÀ-ÖØ-öø-ÿ][^\"“”;()]{{0,260}}?)\s*\)?)?"
        rf"\s*,?\s*(?:that\s+was\s+)?(?:{NOTE_WORD}\s+)?(?:sent|distributed)\s+to\s+(?:the\s+)?(?P<body>{MEMO_BODY_RE})"
        rf"(?:\s+\(corrected\))?"
        rf"(?:\s+on)?\s*,?\s*(?P<date>[A-Z][a-z]+\s+\d{{1,2}}(?:,?\s+{YEAR_RE})?)"
    )
    for m in memo_titled_pat.finditer(text):
        authors = m.group('authors_pre') or m.group('authors_post')
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        date = normalize_memo_date(m.group('date'), default_year=default_year)
        c = build_memo_citation(authors, title, body, date)
        out.append(clean_citation_text(c))

    # "memo (that was)? sent/distributed to (the)? BODY on DATE
    # titled/entitled \"TITLE\" by AUTHORS" -- same as memo_titled_pat but
    # with the sent/distributed-to-DATE clause coming BEFORE the title
    # instead of after. Example:
    #   See the memo sent to the Committee on January 20, 2015 entitled
    #   "Reducing the Aggregate Capacity of the ON RRP Facility" by
    #   Deborah Leonard, Josh Frost, Jane Ihrig, and Gretchen Weinbach.
    memo_sent_first_pat = re.compile(
        rf"(?is){NOTE_WORD}\s+(?:that\s+was\s+)?(?:sent|distributed)\s+to\s+(?:the\s+)?(?P<body>{MEMO_BODY_RE})"
        rf"(?:\s+\(corrected\))?\s+(?:on\s+)?(?P<date>[A-Z][a-z]+\s+\d{{1,2}}(?:,?\s+{YEAR_RE})?)\s+"
        rf"(?:titled|entitled)\s+[\"“](?P<title>.+?)[\"”]"
        rf"\s+by\s+(?P<authors>[A-Za-zÀ-ÖØ-öø-ÿ][^\"“”;()]{{0,260}}?)\s*\."
    )
    for m in memo_sent_first_pat.finditer(text):
        authors = clean_citation_text(m.group('authors'))
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        date = normalize_memo_date(m.group('date'), default_year=default_year)
        c = f'{authors}, "{title}," memo to {body}, {date}'
        out.append(clean_citation_text(c))

    # "memo to (the)? BODY titled \"TITLE\"[, by AUTHORS,]? dated DATE" --
    # "dated DATE" instead of a sent/distributed-to clause. Example:
    #   the memo to the FOMC titled "Recent Movements in Longer-Term
    #   Treasury Yields...," by the staff at the Board and the Federal
    #   Reserve Bank of New York, dated July 14, 2017.
    memo_to_titled_dated_pat = re.compile(
        rf"(?is){NOTE_WORD}\s+to\s+(?:the\s+)?(?P<body>{MEMO_BODY_RE})\s+(?:titled|entitled)\s+"
        rf"[\"“](?P<title>.+?)[\"”]\s*,?\s*(?:by\s+(?P<authors>[A-Za-zÀ-ÖØ-öø-ÿ][^\"“”;()]{{0,260}}?)\s*,?\s*)?"
        rf"dated\s+(?P<date>[A-Z][a-z]+\s+\d{{1,2}}(?:,?\s+{YEAR_RE})?)"
    )
    for m in memo_to_titled_dated_pat.finditer(text):
        authors = m.group('authors')
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        date = normalize_memo_date(m.group('date'), default_year=default_year)
        c = build_memo_citation(authors, title, body, date)
        out.append(clean_citation_text(c))

    # "memo \"TITLE\" prepared for (the)? BODY in DATE (AUTHORS)" -- authors
    # in a trailing parenthetical after a month/year-only date. Example:
    #   The memo "Approaches to Clarifying the Conditionality in the
    #   Committee's Forward Guidance" prepared for the FOMC in September
    #   2011 (Brian Doyle, Spence Hilton, ... of the Board of Governors
    #   and the Federal Reserve Bank of New York) provides...
    memo_prepared_for_pat = re.compile(
        rf"(?is){NOTE_WORD}\s+[\"“](?P<title>.+?)[\"”]\s+prepared\s+for\s+(?:the\s+)?(?P<body>{MEMO_BODY_RE})\s+in\s+"
        rf"(?P<date>[A-Z][a-z]+\s+{YEAR_RE})\s*\(\s*(?P<authors>[A-ZÀ-ÖØ-Þ][^()]{{0,300}}?)\)"
    )
    for m in memo_prepared_for_pat.finditer(text):
        authors = clean_citation_text(m.group('authors'))
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        c = f'{authors}, "{title}," memo to {body}, {m.group("date")}'
        out.append(clean_citation_text(c))

    # "\"TITLE,\" a memo(randum)? written by AUTHORS, which was sent/
    # distributed to (the)? BODY on DATE" -- title leads the sentence, no
    # author precedes it. Example:
    #   "Deflation, Low Inflation, and the Conduct of Monetary Policy," a
    #   memorandum written by Douglas Elmendorf, Dave Reifschneider, and
    #   David Wilcox, which was sent to the FOMC on June 13, 2003.
    memo_title_leads_written_by_pat = re.compile(
        rf"(?is)[\"“](?P<title>.+?)[\"”]\s*,?\s*a\s+{NOTE_WORD}\s+written\s+by\s+"
        rf"(?P<authors>[A-Za-zÀ-ÖØ-öø-ÿ][^\"“”;()]{{0,260}}?)\s*,\s*which\s+was\s+(?:sent|distributed)\s+to\s+"
        rf"(?:the\s+)?(?P<body>{MEMO_BODY_RE})\s+on\s+(?P<date>[A-Z][a-z]+\s+\d{{1,2}}(?:,?\s+{YEAR_RE})?)"
    )
    for m in memo_title_leads_written_by_pat.finditer(text):
        authors = clean_citation_text(m.group('authors'))
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        date = normalize_memo_date(m.group('date'), default_year=default_year)
        c = f'{authors}, "{title}," memo to {body}, {date}'
        out.append(clean_citation_text(c))

    # "AUTHOR, \"TITLE,\" memo(randum) to (the)? BODY, DATE" -- natural
    # bibliography-style order (author and title lead, the memo/note word
    # and recipient come third), as opposed to memo_by_pat/memo_to_pat
    # above where the memo/note word leads. Example:
    #   Yuriy Kitsul, "A Review of Market- and Survey-Based Measures of
    #   Medium- and Longer-Term Inflation Expectations," memorandum to the
    #   Board of Governors, December 10, 2014.
    memo_bib_style_pat = re.compile(
        rf"(?is)(?P<authors>[A-ZÀ-ÖØ-Þ][^\"“”()\n]{{0,200}}?)\s*,\s*[\"“](?P<title>.+?)[\"”]\s*,?\s*"
        rf"{NOTE_WORD}\s+to\s+(?:the\s+)?(?P<body>{MEMO_BODY_RE})\s*,\s*"
        rf"(?P<date>[A-Z][a-z]+\s+\d{{1,2}}(?:,?\s+{YEAR_RE})?)"
    )
    for m in memo_bib_style_pat.finditer(text):
        authors = strip_leading_see_cue(clean_citation_text(m.group('authors')))
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        date = normalize_memo_date(m.group('date'), default_year=default_year)
        c = f'{authors}, "{title}," memo to {body}, {date}'
        out.append(clean_citation_text(c))

    # "\"TITLE,\" memo(randum) to (the)? BODY from AUTHOR[, dated DATE]" --
    # title leads, "from AUTHOR" instead of "by AUTHOR" (a Bluebook-era
    # phrasing). Examples:
    #   "Committee Discussion of the Wording of the Announcement,"
    #   memorandum to the Committee from Vincent Reinhart, dated
    #   October 22, 2003.
    #   "A Potential Change in the Wording of the Risk Assessment,"
    #   memorandum to the Federal Open Market Committee from Vincent
    #   Reinhart, August 5, 2004.
    memo_from_author_pat = re.compile(
        rf'(?is)[\"“](?P<title>.+?)[\"”]\s*,?\s*{NOTE_WORD}\s+to\s+(?:the\s+)?(?P<body>{MEMO_BODY_RE})\s+from\s+'
        rf'(?P<authors>[A-Za-zÀ-ÖØ-öø-ÿ][^\"“”;()]{{0,200}}?)\s*,?\s*(?:dated\s+)?'
        rf'(?P<date>[A-Z][a-z]+\s+\d{{1,2}}(?:,?\s+{YEAR_RE})?)'
    )
    for m in memo_from_author_pat.finditer(text):
        authors = clean_citation_text(m.group('authors'))
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        date = normalize_memo_date(m.group('date'), default_year=default_year)
        c = f'{authors}, "{title}," memo to {body}, {date}'
        out.append(clean_citation_text(c))

    # "[DATE] memo to (the)? BODY, \"TITLE\" by AUTHORS" -- date (if any)
    # comes as a prefix rather than a trailing "sent/distributed/dated"
    # clause; no connector word between BODY and the quoted title, just a
    # comma. Example:
    #   See the September 9 memo to the Committee, "Gauging the Effective
    #   Stance of Monetary Policy," by Jean-Philippe Laforte, Dave
    #   Reifschneider, John Roberts, and Tom Tallarini.
    memo_to_body_comma_title_pat = re.compile(
        rf'(?is)(?:(?P<date>[A-Z][a-z]+\s+\d{{1,2}}(?:,\s+{YEAR_RE})?)\s+)?{NOTE_WORD}\s+to\s+(?:the\s+)?(?P<body>{MEMO_BODY_RE})\s*,\s*'
        rf'[\"“](?P<title>.+?)[\"”]\s+by\s+(?P<authors>[A-Za-zÀ-ÖØ-öø-ÿ][^\"“”;()]{{0,260}}?)\s*\.'
    )
    for m in memo_to_body_comma_title_pat.finditer(text):
        authors = clean_citation_text(m.group('authors'))
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        date_raw = m.group('date') or default_year
        if not date_raw:
            continue
        date = normalize_memo_date(date_raw, default_year=default_year)
        c = f'{authors}, "{title}," memo to {body}, {date}'
        out.append(clean_citation_text(c))

    # "(FOMC|Board) memo titled \"TITLE\" (MONTH YEAR)" -- the institution
    # appears as an adjective directly before "memo" (not "memo to X"),
    # with a bare parenthetical month/year date and no author at all.
    # Example:
    #   the FOMC memo titled "The Federal Reserve's Long-Run Operating
    #   Regime" (November 2018).
    #   the January 18, 2018, Board memo titled "Implications of U.S.
    #   Yield Curve Flattening..."
    memo_adjective_titled_pat = re.compile(
        rf"(?is)(?:(?P<date_prefix>[A-Z][a-z]+\s+\d{{1,2}},\s+{YEAR_RE})\s*,\s*)?"
        rf"(?P<body>FOMC|Board)\s+{NOTE_WORD}\s+titled\s+[\"“](?P<title>.+?)[\"”]"
        rf"(?:\s*\(\s*(?P<date_suffix>[A-Z][a-z]+\s+{YEAR_RE})\s*\))?"
    )
    for m in memo_adjective_titled_pat.finditer(text):
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        date_raw = m.group('date_prefix') or m.group('date_suffix')
        date = normalize_memo_date(date_raw, default_year=default_year) if date_raw else None
        c = build_memo_citation(None, title, body, date)
        out.append(clean_citation_text(c))

    # "[DATE,]? [staff]? memo to (the)? BODY titled \"TITLE.\"" -- date (if
    # any) as a prefix, no author, sentence ends right after the title.
    # Example:
    #   the August 31, 2017, staff memo to the FOMC titled "Preliminary
    #   Assessment of Effects of Hurricane Harvey on the U.S. Economy."
    memo_date_prefix_to_body_titled_pat = re.compile(
        rf"(?is)(?:(?P<date>[A-Z][a-z]+\s+\d{{1,2}},\s+{YEAR_RE})\s*,\s*)?"
        rf"(?:staff\s+)?{NOTE_WORD}\s+to\s+(?:the\s+)?(?P<body>{MEMO_BODY_RE})\s+titled\s+"
        rf"[\"“](?P<title>.+?)[\"”]"
    )
    for m in memo_date_prefix_to_body_titled_pat.finditer(text):
        date_raw = m.group('date')
        if not date_raw:
            continue  # needs at least the date prefix as a signal here, since
            # this pattern has no author requirement and no trailing
            # sent/distributed/dated clause to anchor on
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        date = normalize_memo_date(date_raw, default_year=default_year)
        c = build_memo_citation(None, title, body, date)
        out.append(clean_citation_text(c))

    # "memo entitled \"TITLE,\" to (the)? BODY by AUTHORS" -- has an author
    # but no date anywhere. Example:
    #   refer to the memorandum entitled "Legislative Initiative for
    #   Additional Federal Reserve Balance Sheet Management Tools," to the
    #   Federal Open Market Committee by Brian Madigan, James Clouse,
    #   Scott Alvarez and Sophia Allison.
    memo_entitled_to_body_by_nodate_pat = re.compile(
        rf"(?is){NOTE_WORD}\s+entitled\s+[\"“](?P<title>.+?)[\"”]\s*,?\s*to\s+(?:the\s+)?(?P<body>{MEMO_BODY_RE})\s+by\s+"
        rf"(?P<authors>[A-Za-zÀ-ÖØ-öø-ÿ][^\"“”;()]{{0,260}}?)\s*\."
    )
    for m in memo_entitled_to_body_by_nodate_pat.finditer(text):
        authors = clean_citation_text(m.group('authors'))
        title = m.group('title').strip().strip(" ,.")
        body = normalize_memo_body(m.group('body'))
        c = build_memo_citation(authors, title, body, None)
        out.append(clean_citation_text(c))

    return out


def is_likely_citation_footnote(block: str) -> bool:
    """Reject chart/table numeric fragments before expensive footnote parsing."""
    text = normalize_citation_whitespace(block)
    text = re.sub(r"^\s*\d+\s+", "", text).strip()
    low = text.lower()
    if not text or len(text) < 25:
        return False
    has_citation_shape = bool(
        re.search(rf"(?i)\bsee(?: also)?\b|for further (?:details|information)|{NOTE_WORD}\s+(?:to|by)", text)
        or re.search(r"[\"“].+?[\"”]", text)
        or (re.search(rf"\({YEAR_RE}\)", text) and FULL_CITATION_SIGNAL_RE.search(text))
    )
    if not has_citation_shape:
        return False
    # Reject pure chart/table fragments, but do not reject a valid footnote simply
    # because PyMuPDF appended the page footer ("Class I FOMC ... Page X of Y").
    if re.search(r"\b(ratio scale|blue shaded|selected interest rates|treasury yield curve|dollar exchange rate indexes|journal of commerce index)\b", low):
        if not re.search(r"[\"“].+?[\"”]", text) and not re.search(r"(?i)working paper|unpublished paper|memorandum|memo\s+(?:to|by)", text):
            return False
    return True

def extract_full_footnote_citations(block: str, default_year: str = "") -> List[str]:
    """Extract long bibliographic citations from a numbered footnote block.

    Handles both standard author-year forms and Tealbook staff citations like
    "King, T. B.; Levin, A. T; and Perli, R., ... Working Paper, July 2007"
    where the year appears only at the end.
    """
    text = normalize_citation_whitespace(block)
    text = re.sub(r"^\s*\d+\s+", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    if not is_likely_citation_footnote(block):
        return []
    if not (FULL_CITATION_SIGNAL_RE.search(text) or has_strict_memo_signal(text) or has_title_first_memo_signal(text)):
        return []

    citations: List[str] = []

    citations.extend(extract_nonparen_full_citations(text, default_year=default_year))

    # Standard full citations with parenthetical years.
    #
    # A single footnote often lists several citations separated by "; and"
    # or ";", e.g. "See X (2007), ...; and Y (2011), ...". Without a
    # boundary-aware trim, the fixed-size windows below would either merge
    # both citations into one Citation string, or -- for the second
    # citation -- start mid-way through the first one's trailing text and
    # fail validation, silently dropping it (or leaving only a fragment of
    # its author list to be picked up later by the generic short-citation
    # scanner, e.g. "Wei (2011)" instead of "Canlin Li and Min Wei
    # (2011)"). Trimming each window at the nearest sibling-citation
    # boundary keeps each citation self-contained.
    #
    # Some footnotes instead chain full citations as separate sentences --
    # "...see Smith (2019), ... https://... .  For a related mechanism,
    # see Jones (2018), ... .  See also Lee (2018), ...." -- with no
    # semicolon anywhere. Without a matching boundary there, a later
    # citation's look-back window can walk into an earlier sibling's own
    # text and re-derive *that* citation instead, which then gets dropped
    # as a duplicate -- silently losing the later citation(s) entirely
    # (this is how a footnote with 3 full citations chained this way was
    # only yielding 2). The period-based alternative below treats a
    # sentence boundary immediately followed by a "see"/"see also" cue
    # (optionally preceded by a short "For ..., " lead-in) as the same
    # kind of sibling boundary as ";", without simply matching any period
    # so it doesn't fire mid-citation.
    SIBLING_SEP_RE = re.compile(
        r";\s*(?:and\s+|see\s+also\s+)?"
        r"|\.\s+(?:For\s+[a-z][^,]{0,80},\s+)?(?:see\s+also|see)\s+",
        flags=re.IGNORECASE,
    )
    # Right-trim needs to confirm a real author-list-then-"(YEAR)" follows
    # the separator -- otherwise an unrelated semicolon inside a single
    # citation's own institution/date text (rare, but possible) could
    # truncate it early. Reuses PERSON_NAME_RE/the same comma-plus-"and"
    # author-list shape as AUTHOR_LIST_BEFORE_YEAR_RE elsewhere, so 3+
    # author lists like "A, B, and C (YEAR)" are recognized, not just
    # a single "A and B" pair.
    NEXT_SIBLING_START_RE = re.compile(
        rf"(?:;\s*(?:and\s+|see\s+also\s+)?"
        rf"|\.\s+(?:For\s+[a-z][^,]{{0,80}},\s+)?(?:see\s+also|see)\s+)"
        rf"{PERSON_NAME_RE}(?:\s*,\s*{PERSON_NAME_RE})*(?:\s*,?\s+and\s+{PERSON_NAME_RE})?\s*\({YEAR_RE}\)",
        flags=re.IGNORECASE,
    )
    for year_match in re.finditer(rf"\({YEAR_RE}\)", text):
        left_window = text[max(0, year_match.start() - 420):year_match.start()]
        right_window = text[year_match.start():min(len(text), year_match.end() + 1600)]

        # If a preceding sibling citation's "; and" separator appears in
        # left_window, this citation's author list starts right after it
        # (we already know a valid year follows, at year_match itself, so
        # we don't need to re-confirm one here).
        last_sep_end = None
        for sep_match in SIBLING_SEP_RE.finditer(left_window):
            last_sep_end = sep_match.end()
        if last_sep_end is not None:
            left_window = left_window[last_sep_end:]

        # If a following sibling citation's "; and AUTHOR (YEAR)" boundary
        # appears in right_window, stop before it.
        next_boundary = NEXT_SIBLING_START_RE.search(right_window, year_match.end() - year_match.start())
        if next_boundary:
            right_window = right_window[:next_boundary.start()]

        raw = left_window + right_window
        c = trim_full_citation_tail(raw)
        if is_full_bibliographic_citation(c):
            citations.append(c)

    # Working papers or memos whose only year is in the date at the end.
    nonparen_patterns = [
        rf"(?:See\s+)?[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+,\s*(?:[A-Z]\.\s*)+(?:;\s*[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+,?\s*(?:[A-Z]\.\s*)+)*(?:;?\s*and\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+,?\s*(?:[A-Z]\.\s*)+)?,\s*[“\"].+?[”\"]\.?\s*(?:Working Paper|unpublished paper),\s*(?:[A-Z][a-z]+\s+)?(?:\d{{1,2}},\s*)?{YEAR_RE}\.?,?",
        rf"(?:see\s+)?(?:the\s+)?(?:[A-Z][a-z]+\s+\d{{1,2}},\s+{YEAR_RE}\s+)?{NOTE_WORD}\s+by\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+(?:\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+)*,\s*[“\"].+?[”\"]\.?",
        rf"[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+(?:,?\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+)*(?:,\s*and\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+(?:\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+)*)?,\s*[“\"].+?[”\"]\.?\s*{NOTE_WORD}[^.]*?(?:Federal Open Market Committee|Board of Governors of the Federal Reserve System)[^.]*\.?,?",
    ]
    for pat in nonparen_patterns:
        for m in re.finditer(pat, text, flags=re.IGNORECASE):
            c = strip_leading_see_cue(m.group(0))
            c = trim_full_citation_tail(c)
            if is_full_bibliographic_citation(c):
                citations.append(c)

    # Conservative semicolon fallback for multiple full citations in one footnote.
    for piece in re.split(r"(?<=;)\s+", text):
        if (FULL_CITATION_SIGNAL_RE.search(piece) or has_strict_memo_signal(piece)) and (re.search(rf"\({YEAR_RE}\)", piece) or re.search(YEAR_RE, piece)):
            c = trim_full_citation_tail(piece)
            if is_full_bibliographic_citation(c):
                citations.append(c)

    # De-duplicate only artifacts created by multiple extraction passes within this one footnote.
    cleaned = [clean_citation_text(c) for c in citations if clean_citation_text(c)]
    # Drop partial extractions when a longer extraction contains them.
    filtered: List[str] = []
    for c in cleaned:
        cl = c.lower()
        if any(cl != other.lower() and cl in other.lower() for other in cleaned):
            continue
        filtered.append(c)

    seen = set()
    final: List[str] = []
    for c in filtered:
        key = c.lower()
        if key not in seen:
            seen.add(key)
            final.append(c)
    return final

def get_candidate_citations(text: str) -> List[str]:
    text = re.sub(r"^\s*\d+\s+", "", text)
    text = text.replace("\n", " ")

    anchors = list(re.finditer(r"[A-Z][A-Za-z'\-]+.*?\((?:19|20)\d{2}[a-z]?\)", text))
    if not anchors:
        return []

    citations = []

    for i, match in enumerate(anchors):
        start = match.start()
        next_anchor_start = anchors[i + 1].start() if i + 1 < len(anchors) else len(text)

        semicolon = re.search(r";", text[start:next_anchor_start])
        if semicolon:
            end = start + semicolon.start()
        else:
            end = next_anchor_start

        candidate = text[start:end].strip(" .,;")

        # Accept if:
        # 1. Has publication signal (full citation)
        # OR
        # 2. Looks like structured citation (author + year + title quotes)

        if (
            re.search(r"Journal|Review|Working Paper|FEDS|IFDP|Brookings|Hutchins", candidate)
            or re.search(r"“|”|\"", candidate)  # title quotes
        ):
            citations.append(candidate)

    return citations

def extract_citation_from_context(text: str) -> List[str]:
    """
    Extract clean citation(s) from context.
    Supports both full bibliographic entries and compact multi-year in-text citations.
    """

    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)

    citations: List[str] = []

    # Full bibliographic records. Include repeated-author-dash-normalized entries
    # that already arrive from split_reference_entries().
    if is_full_bibliographic_citation(text):
        citations.append(clean_citation_text(text))

    # Author (year[, year]) citations.
    author_year_pat = re.compile(
        rf"\b[A-Z][A-Za-z'\-]+(?:\s+and\s+[A-Z][A-Za-z'\-]+|\s+et al\.|\s+and others)?\s*\({YEAR_RE}(?:\s*[,;]\s*{YEAR_RE})*\)"
    )
    for match in author_year_pat.finditer(text):
        citations.extend(expand_inline_citation(match.group(0)))

    # Parenthetical citations: (Author, year[, year])
    paren_pat = re.compile(
        rf"\([A-Z][A-Za-z'\-]+(?:\s+and\s+[A-Z][A-Za-z'\-]+|\s+et al\.|\s+and others)?,\s*{YEAR_RE}(?:\s*[,;]\s*{YEAR_RE})*\)"
    )
    for match in paren_pat.finditer(text):
        citations.extend(expand_inline_citation(match.group(0)))

    seen = set()
    final: List[str] = []
    for c in citations:
        c = clean_citation_text(c)
        key = c.lower()
        if key not in seen:
            seen.add(key)
            final.append(c)

    return final

def get_short_anchor(citation: str) -> str:
    """
    Extract a shorter anchor from a full citation so it can still be located
    in the surrounding page text for context extraction.
    """
    # Full bibliography: "Taylor, John B. (1993) ..." -> "Taylor"
    m = re.search(rf"\b([A-Z][A-Za-z'\-]+)(?:,\s*[^()]+)?\s*\({YEAR_RE}\)", citation)
    if m:
        return m.group(1)

    m = re.search(
        rf"\b([A-Z][A-Za-z'\-]+(?:,? [A-Z]\.)?(?: and [A-Z][A-Za-z'\-]+)?(?: et al\.)? \({YEAR_RE}(?:\s*[,;]\s*{YEAR_RE})*\))",
        citation,
    )
    if m:
        return m.group(1)

    m = re.search(rf"\((?:[A-Z][A-Za-z'\-]+(?: et al\.)?|[A-Z][A-Za-z'\-]+ and [A-Z][A-Za-z'\-]+),? {YEAR_RE}(?:\s*[,;]\s*{YEAR_RE})*\)", citation)
    if m:
        return m.group(0)

    return citation[:80]

def extract_footnote_blocks(page_text: str) -> List[tuple[int, str]]:
    """
    Split page text into footnote-like numbered blocks.
    """
    blocks: List[tuple[int, str]] = []
    matches = list(FOOTNOTE_START_RE.finditer(page_text))

    for i, match in enumerate(matches):
        start = match.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(page_text)
        block = page_text[start:end].strip()
        if len(block) >= 20:
            blocks.append((start, block))

    return blocks


def expand_full_citation(page_text: str, start: int, end: int) -> str:
    """
    Expand a detected citation match to a fuller bibliographic span.

    Strategy:
    - expand outward within the surrounding paragraph/footnote block
    - stop at double newlines or likely next footnote starts
    - normalize line wraps so journal titles continue across lines
    """
    text = page_text

    left = start
    while left > 0:
        if text[max(0, left - 2):left] == "\n\n":
            break
        if re.search(r"\n\s*\d+\s+$", text[max(0, left - 12):left]):
            break
        left -= 1

    right = end
    while right < len(text):
        if text[right:right + 2] == "\n\n":
            break
        if re.match(r"\n\s*\d+\s+[A-Z]", text[right:right + 16]):
            break
        if text[right] == ";":   # <-- ADD THIS LINE
            right += 1
            break
        right += 1

    candidate = text[left:right].strip()
    candidate = clean_citation_text(candidate)

    return candidate




def citation_anchor_surname_year(citation: str) -> tuple[str, str] | None:
    """Return a conservative (surname, year) anchor for suppressing inline
    matches that are actually inside a full bibliographic footnote/reference.
    """
    m = re.search(rf"\((?P<year>{YEAR_RE})\)", citation)
    if not m:
        return None
    year = m.group("year")
    prefix = citation[:m.start()]
    # Last capitalized token before the year is usually the cited surname.
    names = re.findall(r"[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+", prefix)
    if not names:
        return None
    return (names[-1].lower(), year.lower())


def inline_match_looks_inside_full_reference(text: str, start: int, end: int, match_text: str, full_anchors: set[tuple[str, str]]) -> bool:
    """Suppress false duplicate inline matches inside full footnote citations.

    Example: a full footnote row for "Reifschneider, David and John C. Williams (2000)..."
    can otherwise produce a second row "Williams (2000)".  We only suppress when
    the same surname-year anchor appears in a local window that has full-reference
    signals such as quoted title, journal/source, page numbers, or footnote closure.
    This preserves repeated ordinary in-text mentions on the same page.
    """
    anchor = citation_anchor_surname_year(match_text)
    if not anchor or anchor not in full_anchors:
        return False

    window = text[max(0, start - 280):min(len(text), end + 900)]
    return bool(re.search(
        r"[\"“”]|\bJournal\b|\bReview\b|\bEconometrica\b|\bVol\.\b|\bvol\.\b|\bpp\.\b|"
        r"\bWorking\s+[Pp]aper\b|\bunpublished\s+paper\b|\bmemo(?:randum)?\b|"
        r"End\s+(?:box\s+)?footnote|\[End\s+(?:box\s+)?footnote",
        window,
        flags=re.IGNORECASE,
    ))

def detect_citations(page_text: str, default_year: str = "") -> List[tuple[str, int]]:
    """Detect citation instances on a page.

    Returns a list of (citation_text, occurrence) tuples. `occurrence` is the
    1-indexed count of that exact citation string on this page (e.g. the 2nd
    time "Taylor (1999)" appears), so downstream context extraction can find
    the correct surrounding text for each individual mention rather than
    always the first one.

    This intentionally returns repeated in-text instances on the same page.
    The only suppression is for inline author-year matches that occur inside a
    footnote block already captured as a full bibliographic citation, which
    prevents false double counts such as a full Reifschneider and Williams
    reference plus a separate Williams (2000) parsed from inside that reference.
    """

    results: List[tuple[str, int]] = []
    full_seen: set[str] = set()
    full_inline_anchors: set[tuple[str, str]] = set()
    occurrence_counts: dict[str, int] = {}

    def add(c: str):
        c = clean_citation_text(c)
        if not c:
            return
        # Deduplicate only full bibliographic/footnote/reference artifacts on
        # a page. Repeated ordinary in-text mentions are intentionally kept.
        if is_full_bibliographic_citation(c):
            key = re.sub(r"\s+", " ", c.lower()).strip(" .,)]")
            if key in full_seen:
                return
            full_seen.add(key)
            anchor = citation_anchor_surname_year(c)
            if anchor:
                full_inline_anchors.add(anchor)
        occ_key = c.lower()
        occurrence_counts[occ_key] = occurrence_counts.get(occ_key, 0) + 1
        results.append((c, occurrence_counts[occ_key]))

    # If the page has a References section, extract its entries as full
    # bibliographic rows first -- but don't return early. A page can (and in
    # Tealbook appendices often does) also carry ordinary inline mentions in
    # the body text above that heading, e.g. "...proposed by Taylor (1993,
    # 1999)." and "...see Orphanides (2003)." earlier on the same page as a
    # References list. Those are genuinely separate citation instances and
    # should still be counted.
    #
    # To avoid the opposite failure -- a reference entry like
    # "Blaise, Emily (2005). ... Journal of ..." getting matched a second
    # time as a bare inline "Blaise (2005)" -- everything from the
    # References heading to the end of the page is excluded from every
    # scan below (institutional, footnote, and plain inline). Only the body
    # text above the heading is scanned for inline citations.
    scan_text = page_text
    if is_reference_page(page_text):
        for entry in split_reference_entries(page_text):
            add(entry)
        ref_start = find_reference_section_start(page_text)
        if ref_start is not None:
            scan_text = page_text[:ref_start]

    inline_source = scan_text

    # Hard-coded institutional source citations, such as Congressional Budget
    # Office (2008), should keep the full agency name and should not later
    # reappear as Office (2008) or System (2008).
    for _start, _end, inst_cite in extract_known_institutional_citations_with_spans(scan_text):
        add(inst_cite)
    for author in KNOWN_INSTITUTIONAL_AUTHORS:
        inline_source = re.sub(re.escape(author) + r"\s*\(", " ", inline_source)

    # Full citations in numbered footnotes need different treatment from
    # ordinary in-text citations. If a footnote has at least one full citation,
    # remove that exact footnote span from the later inline pass to avoid
    # parsing author-year strings inside the full reference as separate
    # citations. Use span-based removal instead of str.replace(block), because
    # block is stripped/normalized by extraction and may not exactly match the
    # raw page text in boxes/footers.
    footnote_remove_spans: List[tuple[int, int]] = []
    for pos, block in extract_footnote_blocks(scan_text):
        fulls = extract_full_footnote_citations(block, default_year=default_year)
        if fulls:
            for c in fulls:
                add(c)

            raw_start = scan_text.find(block, max(0, pos - 5))
            if raw_start == -1:
                raw_start = scan_text.find(block)
            if raw_start == -1:
                raw_start = max(0, pos)
                raw_end = min(len(scan_text), pos + len(block))
            else:
                raw_end = raw_start + len(block)
            footnote_remove_spans.append((raw_start, raw_end))

    if footnote_remove_spans:
        chars = list(inline_source)
        for start, end in sorted(footnote_remove_spans, reverse=True):
            start = max(0, min(start, len(chars)))
            end = max(start, min(end, len(chars)))
            chars[start:end] = " " * (end - start)
        inline_source = "".join(chars)

    text = normalize_citation_whitespace(inline_source)

    # Memo/note/working-paper citations often appear as inline body text or
    # boxed notes rather than inside a numbered footnote -- for example a
    # highlighted box reading (across several lines):
    #   "Title"
    #   by AUTHORS
    #   distributed to the Committee, DATE
    # extract_full_footnote_citations() above only looks inside footnote
    # blocks (text starting with a footnote number), so it misses these.
    # Run the same title-first / memo-by / memo-to patterns across the whole
    # normalized page text as well. add() already de-duplicates identical
    # full-citation text, so anything also caught by the footnote pass above
    # is not double-counted.
    for c in extract_nonparen_full_citations(text, default_year=default_year):
        add(c)

    author_tail = r"(?:\s+and\s+[A-Z][A-Za-z'\-]+|\s+et al\.|\s+and others)?"
    author_year_pat = re.compile(
        rf"\b[A-Z][A-Za-z'\-]+{author_tail}\s*\({YEAR_RE}(?:\s*[,;]\s*{YEAR_RE})*\)"
    )
    for match in author_year_pat.finditer(text):
        if inline_match_looks_inside_full_reference(text, match.start(), match.end(), match.group(0), full_inline_anchors):
            continue
        for c in expand_inline_citation(match.group(0)):
            add(c)

    paren_pat = re.compile(
        rf"\([A-Z][A-Za-z'\-]+{author_tail},\s*{YEAR_RE}(?:\s*[,;]\s*{YEAR_RE})*\)"
    )
    for match in paren_pat.finditer(text):
        if inline_match_looks_inside_full_reference(text, match.start(), match.end(), match.group(0), full_inline_anchors):
            continue
        for c in expand_inline_citation(match.group(0)):
            add(c)

    return results

def classify_source_type(citation: str) -> str:
    c = citation.lower()

    # Institutional/Fed signals take precedence over author overrides so that
    # Fed memos or Board materials are not mislabeled as Academic just because
    # they mention an academic-style author name.
    if is_known_institutional_author(citation):
        if c.startswith("board of governors of the federal reserve system") or c.startswith("federal reserve bank of"):
            return "Fed Research"
        return "Government / International Organization"
    if any(marker in c for marker in FED_RESEARCH_MARKERS):
        return "Fed Research"
    # Internal Board/FOMC memos and notes (e.g. "memo to FOMC", "note to the
    # Committee") are Fed Research regardless of phrasing/author, even when
    # they don't literally contain "memorandum to the fomc".
    if has_strict_memo_signal(citation):
        return "Fed Research"
    if any(marker in c for marker in GOV_IO_MARKERS):
        return "Government / International Organization"

    # Stable academic author overrides.  This catches compact rule labels such
    # as Taylor (1993), Taylor (1999), Orphanides (2003), etc., even when the
    # surrounding text does not include a journal or working-paper signal.
    if has_academic_author_override(citation):
        return "Academic"

    if any(marker in c for marker in PRIVATE_SECTOR_MARKERS):
        return "Private Sector / Market"
    if any(marker in c for marker in ACADEMIC_MARKERS):
        return "Academic"
    return "Unknown"


def classify_paper_type(citation: str, source_type: str) -> str:
    c = citation.lower()

    if is_known_institutional_author(citation):
        return "Government Report"

    if any(re.search(pat, citation) for pat in JOURNAL_PATTERNS):
        return "Journal Article"
    if any(marker in c for marker in [
        "working paper",
        "discussion paper",
        "policy research working paper",
        "occasional paper",
        "technical paper",
        "feds notes",
        "ifdp",
        "working paper,",
        "unpublished paper",
    ]):
        return "Working Paper"
    if has_strict_memo_signal(citation):
        return "Staff Memo / Note"
    if "staff report" in c or "staff reports" in c:
        return "Staff Report"
    if "report" in c and source_type == "Government / International Organization":
        return "Government Report"
    if any(marker in c for marker in [
        "university press",
        " ed.",
        " eds.",
        " in ",
        "chapter",
    ]):
        return "Book / Chapter"
    if any(marker in c for marker in [
        "survey",
        "forecast",
        "consensus",
        "blue chip",
    ]):
        return "Market Forecast / Survey"
    return "Unknown"


def extract_pub_name(citation: str) -> str:
    citation = normalize_citation_whitespace(citation)

    patterns = [
        # Journal with volume + pages
        r"(Journal of [A-Za-z&\- ,:;]+?, vol\. ?\d+[^;\n]*)",
        r"(Review of [A-Za-z&\- ,:;]+?, vol\. ?\d+[^;\n]*)",
        r"(Quarterly Journal of [A-Za-z&\- ,:;]+?, vol\. ?\d+[^;\n]*)",

        # fallback: just journal name if volume missing
        r"(Journal of [A-Za-z&\- ,:;]+)",
        r"(Review of [A-Za-z&\- ,:;]+)",
        r"(Quarterly Journal of [A-Za-z&\- ,:;]+)",

        r"(American Economic Review[^;\n]*)",
        r"(Econometrica[^;\n]*)",
        r"(Brookings Papers on Economic Activity[^;\n]*)",

        r"(NBER Working Paper No\.? ?\d+)",
        r"(FEDS Notes[^;\n]*)",
        r"(IFDP)",
        r"(FRBNY Staff Reports?[^;\n]*)",
        r"(Finance and Economics Discussion Series)",
        r"(Policy Research Working Paper(?: No\.? ?\d+)?)",
    ]

    for pattern in patterns:
        match = re.search(pattern, citation, re.IGNORECASE)
        if match:
            return match.group(1).strip(" .,;:")

    return ""


def extract_page_text(page: fitz.Page) -> str:
    text = page.get_text("text")
    return normalize_whitespace(text)




def full_citation_year(citation: str) -> str:
    """Best-effort year for reconciling duplicate full citations."""
    m = re.search(rf"\(({YEAR_RE})\)", citation)
    if m:
        return m.group(1).lower()
    years = re.findall(YEAR_RE, citation)
    return years[-1].lower() if years else ""


def _surname_from_segment(segment: str) -> str:
    """Return the likely surname from one author-list segment."""
    segment = clean_citation_text(segment)
    # Bibliography style: "Reifschneider, David" or "King, T. B."
    m = re.match(r"^([A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+)\s*,", segment)
    if m:
        return m.group(1).lower()
    names = re.findall(r"[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+", segment)
    return names[-1].lower() if names else ""


def full_citation_author_key(citation: str) -> str:
    """Conservative author key for full-reference reconciliation.

    It intentionally uses only author information before the citation year/title,
    so one footnote found by two extraction paths maps to the same key, while
    ordinary repeated in-text instances are not processed here.
    """
    citation = clean_citation_text(citation)
    low = citation.lower()

    for author in KNOWN_INSTITUTIONAL_AUTHORS:
        if low.startswith(author.lower()):
            return re.sub(r"[^a-z0-9]+", "_", author.lower()).strip("_")

    year_match = re.search(rf"\({YEAR_RE}\)", citation)
    if year_match:
        prefix = citation[:year_match.start()]
    else:
        # Non-parenthetical-year memos/working papers: stop before the title.
        q = re.search(r"[\"“]", citation)
        prefix = citation[:q.start()] if q else citation[:260]

    prefix = clean_citation_text(prefix)
    prefix = re.sub(r"(?i)^(?:source|see(?: also)?|for example|e\.g\.|note)\s*[:.]?\s+", "", prefix)

    # Split common author-list forms.
    parts = re.split(r"\s*;\s*|\s*,?\s+and\s+", prefix)
    surnames = [_surname_from_segment(part) for part in parts if _surname_from_segment(part)]

    # For "Reifschneider, David and John C. Williams", the split gives
    # Reifschneider plus Williams.  For unsplit full-name lists, take first/last.
    if not surnames:
        names = re.findall(r"[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'\-]+", prefix)
        if names:
            surnames = [names[-1].lower()]

    if len(surnames) >= 2:
        return "|".join([surnames[0], surnames[-1]])
    return surnames[0] if surnames else re.sub(r"[^a-z0-9]+", "_", prefix.lower()).strip("_")[:60]


def full_citation_title_key(citation: str) -> str:
    """Short title/source signature to avoid over-collapsing same author/year."""
    citation = clean_citation_text(citation)
    m = re.search(r"[\"“](.+?)[\"”]", citation)
    if m:
        title = m.group(1)
    else:
        # Institution/report citations may not use quotes; use words after year.
        ym = re.search(rf"\({YEAR_RE}\)", citation)
        title = citation[ym.end():] if ym else citation
    title = re.sub(r"[^a-z0-9]+", " ", title.lower()).strip()
    words = title.split()
    return "_".join(words[:8])


def full_citation_reconcile_key(row: CitationRow) -> str:
    citation = clean_citation_text(row.Citation)
    year = full_citation_year(citation)
    author_key = full_citation_author_key(citation)
    title_key = full_citation_title_key(citation)
    return f"{row.Date}|{row.Tealbook}|{row.Page}|{author_key}|{year}|{title_key}"


def full_citation_quality_score(row: CitationRow) -> int:
    """Score candidate full citations; higher is cleaner/better."""
    c = clean_citation_text(row.Citation)
    low = c.lower()
    score = 0

    if starts_like_author_citation(c) or is_known_institutional_author(c):
        score += 6
    if re.search(r"[\"“].+?[\"”]", c):
        score += 4
    if any(pat in low for pat in ["journal", "review", "econometrica", "working paper", "unpublished paper"]):
        score += 4
    if has_strict_memo_signal(c):
        score += 4
    if re.search(r"\bpp\.\s*\d+", c, re.IGNORECASE):
        score += 2
    if re.search(r"\bvol\.\s*\d+", c, re.IGNORECASE):
        score += 2
    if row.Paper_Type and row.Paper_Type != "Unknown":
        score += 2
    if row.Source_Type and row.Source_Type != "Unknown":
        score += 1

    # Penalize spillover/body text and extraction artifacts.
    bad_patterns = [
        r"end\s+(?:box\s+)?footnote",
        r"\[end\s+(?:box\s+)?footnote",
        r"\bhowever,",
        r"\bmoreover,",
        r"\bconsequently,",
        r"\bdown sharply\b",
        r"\blast page of\b",
        r"\bclass i fomc\b",
        r"\bgap<",
        r"\bincentive to refinance\b",
        r"\bfigure\s+\d+\b",
        r"\bchart\s+\d+\b",
        r"\btable\s+\d+\b",
    ]
    for pat in bad_patterns:
        if re.search(pat, low, flags=re.IGNORECASE):
            score -= 8

    if re.match(r"(?i)^(?:this|the|however|moreover|consequently|in contrast|with these|for more discussion)\b", c):
        score -= 8
    if len(c) > 700:
        score -= 4
    if len(c) > 1100:
        score -= 8

    return score


def is_reconcilable_full_row(row: CitationRow) -> bool:
    """Only full bibliographic/reference rows are reconciled.

    Ordinary in-text citations, including repeated Taylor (1999) mentions, are
    left untouched even if Source_Type is Academic from an author override.
    """
    if row.Paper_Type == "Unknown":
        return False
    return is_full_bibliographic_citation(row.Citation) or row.Paper_Type in {
        "Journal Article",
        "Working Paper",
        "Staff Memo / Note",
        "Staff Report",
        "Government Report",
        "Book / Chapter",
        "Market Forecast / Survey",
    }


def reconcile_full_citation_rows(rows: List[CitationRow]) -> List[CitationRow]:
    """Deduplicate only full-reference artifacts on the same page.

    Multiple in-text mentions are citation instances and remain as-is.  Full
    footnote/reference rows are sources, so if two extractors find the same
    author/year/title on the same page, keep the cleanest version.
    """
    best_by_key: dict[str, tuple[int, int, CitationRow]] = {}
    passthrough: List[tuple[int, CitationRow]] = []

    for idx, row in enumerate(rows):
        if not is_reconcilable_full_row(row):
            passthrough.append((idx, row))
            continue

        key = full_citation_reconcile_key(row)
        score = full_citation_quality_score(row)
        # Prefer higher score; on ties, prefer shorter citation to avoid spillover.
        tie = -len(clean_citation_text(row.Citation))
        current = best_by_key.get(key)
        if current is None or (score, tie) > (current[0], current[1]):
            best_by_key[key] = (score, tie, row)

    kept_full = [(rows.index(item[2]) if item[2] in rows else 10**9, item[2]) for item in best_by_key.values()]
    combined = passthrough + kept_full
    combined.sort(key=lambda x: x[0])
    return [row for _idx, row in combined]

def process_pdf(pdf_path: Path) -> List[CitationRow]:
    date, year, tealbook = parse_filename_metadata(pdf_path.name)
    if not date or not tealbook:
        print(f"Warning: could not fully parse filename metadata for {pdf_path.name}", file=sys.stderr)

    rows: List[CitationRow] = []

    try:
        with fitz.open(pdf_path) as doc:
            for page_idx in range(len(doc)):
                page = doc[page_idx]
                page_text = extract_page_text(page)
                if not page_text:
                    continue

                section = find_section(page_text)
                # Page is the physical position of this page within the PDF
                # file (1-indexed) -- always unambiguous, requires no
                # parsing, and is what "go to page" in a PDF viewer uses.
                # Printed_Page is a separate, best-effort attempt to read
                # the page number as it appears on the page itself, left
                # blank when the format can't be confidently recognized
                # (see extract_printed_page_number).
                physical_page = str(page_idx + 1)
                printed_page = extract_printed_page_number(page)
                citations = detect_citations(page_text, default_year=year)

                for raw_citation, occurrence in citations:
                    # Always pull real surrounding text for Context (e.g. "See
                    # this memo ..."), including for full bibliographic
                    # citations -- get_context() locates them via a quoted
                    # title / short author-year / surname anchor, so it no
                    # longer just echoes the citation back at itself.
                    # `occurrence` ensures repeated mentions of the same
                    # citation each get their own accurate surrounding text.
                    context = get_context(page_text, raw_citation, occurrence=occurrence)

                    # detect_citations() already expands multi-year mentions and
                    # parses reference-list entries. Do not re-extract from the
                    # context here, or one compact mention can duplicate rows.
                    cleaned_citations = [raw_citation]

                    for citation in cleaned_citations:
                        source_type = classify_source_type(citation)
                        paper_type = classify_paper_type(citation, source_type)
                        pub_name = extract_pub_name(citation)

                        rows.append(
                            CitationRow(
                                Date=date,
                                Year=year,
                                Tealbook=tealbook,
                                Meeting_Number="",  # filled in by assign_meeting_numbers() once all files are known
                                Citation=citation,
                                Context=context,
                                Section=section,
                                Page=physical_page,
                                Printed_Page=printed_page,
                                Source_Type=source_type,
                                Paper_Type=paper_type,
                                Pub_Name=pub_name,
                            )
                        )
    except Exception as e:
        print(f"Warning: skipping unreadable PDF {pdf_path}: {e}", file=sys.stderr)

    return reconcile_full_citation_rows(rows)


def is_real_pdf(path: Path) -> bool:
    name = path.name
    parts = set(path.parts)

    if path.suffix.lower() != ".pdf":
        return False
    if name.startswith("._") or name.startswith("."):
        return False
    if "__MACOSX" in parts:
        return False

    return True


def extract_zip(zip_path: Path, out_dir: Path) -> List[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as zf:
        for member in zf.infolist():
            member_path = Path(member.filename)
            name = member_path.name

            if name.startswith("._") or name.startswith("."):
                continue
            if "__MACOSX" in member_path.parts:
                continue

            zf.extract(member, out_dir)

    return sorted(p for p in out_dir.rglob("*.pdf") if is_real_pdf(p))


def sanitize_worksheet_formulas(ws) -> None:
    """Undo openpyxl's automatic formula detection for ordinary text.

    openpyxl marks any string cell whose value starts with "=" as a
    formula the moment it's assigned, regardless of whether it's an actual
    formula. Tealbook footnotes are full of model equations (e.g.
    "y = C + I + G + NX"), and the Context column's word-window can land
    right on a stray "=" token, so this fires occasionally and always on
    real data. The resulting "formula" isn't valid, so Excel silently
    strips it on open and reports a repair -- which can blank or truncate
    that cell's content, not just show a scary dialog.

    Resetting the cell's data_type back to plain string preserves the
    exact original text (including the leading "="): openpyxl already
    stores the untouched string as cell.value even when data_type is "f",
    so nothing needs to be rewritten, only reclassified.
    """
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type == "f" and isinstance(cell.value, str):
                cell.value = str(cell.value)
                cell.data_type = "s"


CHAINED_YEAR_ARTIFACT_RE = re.compile(r"(?i)^chained\s*\(\s*(?:19|20)\d{2}[a-z]?\s*\)$")


def drop_chained_year_artifacts(rows: List[CitationRow]) -> List[CitationRow]:
    """Drop rows whose Citation is just "Chained (YEAR)" -- an inline-citation
    false positive from table captions like "Billions of chained (2005)
    dollars", which matches the ordinary Author (Year) shape closely enough
    to get picked up. Applied only as a final filter on the assembled rows,
    not inside detection itself, so it can never suppress a genuine
    citation partway through extraction -- it only removes rows whose
    Citation text, in full, is exactly this artifact.
    """
    return [row for row in rows if not CHAINED_YEAR_ARTIFACT_RE.match(row.Citation.strip())]


def assign_meeting_numbers(rows: List[CitationRow]) -> List[CitationRow]:
    """Number each FOMC meeting within a year in chronological order: the
    year's earliest Tealbook date is meeting 1, the next distinct date is
    meeting 2, and so on. This can't be hardcoded off a specific
    month/day, since meeting dates shift from year to year (and across
    different eras of the FOMC's meeting calendar).

    Book A and Book B for the same meeting share a Date and therefore the
    same Meeting_Number, since they're companion documents for one
    meeting, not two separate ones.

    Numbering is computed relative to the Dates actually present in the
    rows passed in. If a year's Tealbooks are processed across multiple
    separate runs rather than all together, numbering won't be consistent
    between runs unless every run includes that year's complete set of
    dates.
    """
    dates_by_year: "dict[str, set[str]]" = {}
    for row in rows:
        dates_by_year.setdefault(row.Year, set()).add(row.Date)

    # Dates are YYYYMMDD strings, which sort chronologically as plain text.
    meeting_number_by_year_date: "dict[tuple[str, str], int]" = {}
    for year, dates in dates_by_year.items():
        for i, date in enumerate(sorted(dates), start=1):
            meeting_number_by_year_date[(year, date)] = i

    return [
        replace(row, Meeting_Number=str(meeting_number_by_year_date[(row.Year, row.Date)]))
        for row in rows
    ]


# Every "Taylor (1993)" / "Taylor (1999)" mention in a Tealbook -- whether a
# short inline mention or a full reference-list entry -- refers to one of
# exactly two specific, well-known papers; this is standard usage across
# every Tealbook, not something that needs per-citation classification:
#   Taylor (1993): Taylor, John B., "Discretion versus Policy Rules in
#     Practice," Carnegie-Rochester Conference Series on Public Policy,
#     Vol. 39 (December), pp. 195-214.
#   Taylor (1999): Taylor, John B., "A Historical Analysis of Monetary
#     Policy Rules," in John B. Taylor, ed., Monetary Policy Rules,
#     University of Chicago Press, pp. 319-341.
TAYLOR_1993_RE = re.compile(r"(?i)\bTaylor\b[^()]{0,80}\(1993\)")
TAYLOR_1999_RE = re.compile(r"(?i)\bTaylor\b[^()]{0,80}\(1999\)")


def apply_taylor_rule_hardcodes(rows: List[CitationRow]) -> List[CitationRow]:
    """Hardcode Paper_Type/Pub_Name for Taylor (1993)/(1999) mentions.

    This is more accurate than per-citation classification, which has no
    way to recover the publication venue from a short inline mention like
    "Taylor (1999)" or "inertial Taylor (1999) rule" at all -- and more
    consistent than relying on each occurrence's own (sometimes
    incomplete) surrounding text. Applied as a final pass on the
    assembled rows, like drop_chained_year_artifacts, so it never
    interferes with detection.
    """
    updated = []
    for row in rows:
        if TAYLOR_1993_RE.search(row.Citation):
            row = replace(
                row,
                Paper_Type="Journal Article",
                Pub_Name="Carnegie-Rochester Conference Series on Public Policy",
            )
        elif TAYLOR_1999_RE.search(row.Citation):
            row = replace(row, Paper_Type="Book / Chapter", Pub_Name="Monetary Policy Rules")
        updated.append(row)
    return updated


def main() -> int:
    args = parse_args()

    zip_path = Path(args.zip_path).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()
    extract_dir = output_path.parent / f"{output_path.stem}__extracted"

    if not zip_path.exists():
        print(f"Input zip not found: {zip_path}", file=sys.stderr)
        return 1

    pdfs = extract_zip(zip_path, extract_dir)
    if not pdfs:
        print("No valid PDFs found in zip file.", file=sys.stderr)
        return 1

    all_rows: List[CitationRow] = []
    for pdf in pdfs:
        print(f"Processing {pdf.name}...", file=sys.stderr)
        all_rows.extend(process_pdf(pdf))

    before_count = len(all_rows)
    all_rows = drop_chained_year_artifacts(all_rows)
    dropped = before_count - len(all_rows)
    if dropped:
        print(f"Dropped {dropped} 'Chained (YEAR)' artifact row(s).", file=sys.stderr)

    all_rows = assign_meeting_numbers(all_rows)
    all_rows = apply_taylor_rule_hardcodes(all_rows)

    df = pd.DataFrame([asdict(row) for row in all_rows], columns=OUTPUT_COLUMNS)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        sanitize_worksheet_formulas(writer.sheets["Sheet1"])

    print(f"Wrote {len(df)} citation rows to {output_path}")

    if not args.keep_extracted:
        try:
            import shutil
            shutil.rmtree(extract_dir)
        except Exception as e:
            print(f"Warning: could not remove extracted directory {extract_dir}: {e}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
