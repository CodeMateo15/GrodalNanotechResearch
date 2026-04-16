"""
Convert txt files into Excel format matching Combined_community_worksheet.xlsx.

V5: Adds Newspaper source, replaces Business RTF with Business TXT,
    fixes Science separator threshold, adds Molecular Engineering /
    Financial / Commerce / Application keywords, dual date columns
    with match percentage.

HOW TO USE:
1. Place this script in the same folder as your .txt files
2. Run: python create_excelV5.py
"""

import re
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment

# ── CONFIGURATION ────────────────────────────────────────────────────────────

FILE_SOURCE_MAP = {
    "government.txt":           ("Government",        1, "government"),
    "Science_news.txt":         ("Science News",       2, "after_label"),
    "Science_research.txt":     ("Science Research",   3, "after_label"),
    "Business_press.txt":       ("Business Press",     4, "business_press"),
    "Business.txt":             ("Business",           5, "business_press"),
    "Old Python Files/Business_2005.rtf": ("Business", 5, "business_press"),
    "futurists.txt":            ("Futurists",          6, "futurist"),
    "Newspapers_1984-2005.txt": ("Newspapers",         7, "newspaper"),
}

YEAR_MIN = 1983
YEAR_MAX = 2005

OUTPUT_FILE = "output.xlsx"

# Reference worksheet with dates for sources that lack them in the raw text
REFERENCE_WORKSHEET = "Combined community worksheet.xlsx"

# ── SEPARATORS ───────────────────────────────────────────────────────────────

SEP_ASTERISKS = re.compile(r'^\s*\*{19,}\s*$', re.MULTILINE)
SEP_BUSINESS  = re.compile(r'^\s*Article\s+\d+\s+\*{19,}', re.MULTILINE)
SEP_NEWSPAPER = re.compile(
    r'(?:^\s*Article\s+\d+\s+\*{19,}|^\s*\*\s+\*\s+\*\s*$)',
    re.MULTILINE
)

# ── SECTION LABELS (Science-style) ──────────────────────────────────────────

SECTION_LABELS = re.compile(
    r'^(Reports?|Policy\s+Forum|News|Perspectives?|Reviews?|Letters?|'
    r'Editorial|Research\s+Article|Brief\s+Communications?|Brevia|'
    r'Random\s+Samples?|Newsmakers?|News\s+Focus|'
    r'News\s+of\s+the\s+Week|Findings|'
    r'This\s+Week\s+in\s+Science|Essays?|Corrections?)$',
    re.IGNORECASE
)

# ── REFERENCES ───────────────────────────────────────────────────────────────

REFERENCES_START = re.compile(
    r'^(References(\s+and\s+Notes)?|Bibliography|Notes|'
    r'Supporting\s+(Online\s+)?Material|SOM\s+Text|'
    r'Acknowledgements?|Supplementary\s+Materials?)$',
    re.IGNORECASE
)

# ── BUSINESS PRESS ───────────────────────────────────────────────────────────

COPYRIGHT_LINE = re.compile(r'^\(?\s*Copyright|\(c\)\s*\d{4}', re.IGNORECASE)

BUSINESS_TAIL_JUNK = re.compile(
    r'^\s*(Document\s+\S+|More Like This|Retrieving article\(s\)\.\.\.'
    r'|(?:Article\s+\d+\s+)?Next Article)\s*$',
    re.IGNORECASE
)

BUSINESS_HEADER_JUNK = re.compile(
    r'^\s*(Retrieving article|(?:Article\s+\d+\s+)?Next Article|'
    r'D&B Report|Financial Snapshot|Company Quick Search|'
    r'Company Data from|Add to Company List|'
    r'Analyst Report|Comparison Report|Quote|News|Details|Chart|'
    r'There is no related information)',
    re.IGNORECASE
)

# ── SCIENCE METADATA ────────────────────────────────────────────────────────

SCIENCE_META = re.compile(
    r'^(\d{1,5}\s*$|'
    r'DOI:\s|doi:\s|'
    r'Vol\.\s*\d|'
    r'Prev\s*\|\s*Table of Contents|'
    r'Science\s+\d{1,2}\s+\w+\s+\d{4}|'
    r'Science,\s+New\s+Series|'
    r'Copyright\s*\(?|'
    r'Originally\s+published|'
    r'\d{1,2}\s+\w+\s+\d{4}\s*$|'
    r'\d{1,2}\s+\w+\s+\d{4}\s+VOL\s+\d|'
    r'\*?\s*To whom correspondence)',
    re.IGNORECASE
)

# ── DATES ────────────────────────────────────────────────────────────────────

DATE_PATTERNS = [
    (r'\b(\d{1,2}\s+\w+\s+\d{4})\b',    '%d %B %Y'),
    (r'\b(\w+\s+\d{1,2},\s+\d{4})\b',   '%B %d, %Y'),
    (r'\b(\d{1,2}\s+\w{3}\s+\d{4})\b',  '%d %b %Y'),
    (r'\b(\w{3}\s+\d{1,2},\s+\d{4})\b', '%b %d, %Y'),
    (r'\b(\d{4}-\d{2}-\d{2})\b',         '%Y-%m-%d'),
]

MONTH_NAMES = {
    'january','february','march','april','may','june',
    'july','august','september','october','november','december',
    'jan','feb','mar','apr','jun','jul','aug','sep','oct','nov','dec'
}

# ── TOPIC KEYWORDS ───────────────────────────────────────────────────────────

TOPIC_KEYWORDS = {
    "Space":            [r'\bspace\b', r'\bsatellite', r'\baerospace', r'\borbit', r'\brocket', r'\bnasa\b', r'\bspacecraft'],
    "Electronics":      [r'\belectronic', r'\bcircuit', r'\btransistor', r'\bdiode', r'\bchip\b'],
    "Artificial Intelligence": [r'\bartificial intelligence\b', r'\bmachine learning\b', r'\bneural network', r'\bdeep learning\b', r'\bai\b', r'\bgradient descent\b'],
    "Photonics":        [r'\bphotonic', r'\boptical\b', r'\blaser', r'\bfiber optic', r'\bphoton'],
    "Biotech/Biology":  [r'\bbiotech', r'\bbiology\b', r'\bbiological\b', r'\bgenetic', r'\bgenome', r'\bprotein\b', r'\bcell\b', r'\bbacterial'],
    "Semiconductors":   [r'\bsemiconductor', r'\bsilicon\b', r'\bgallium', r'\bdoping\b', r'\bwafer'],
    "Robotics":         [r'\brobot', r'\bautonomo', r'\bmanipulator', r'\bactuat'],
    "Computers/Computing": [r'\bcomput', r'\bprocessor', r'\bsoftware\b', r'\bhardware\b', r'\bwireless\b', r'\bdigital\b', r'\bmicroprocessor'],
    "Material Science": [r'\bmaterial science', r'\bcomposite\b', r'\balloy\b', r'\bpolymer\b', r'\bceramics?\b', r'\bcoating'],
    "Cleantech":        [r'\bcleantech\b', r'\brenewable', r'\bsolar\b', r'\bwind energy', r'\bclean energy', r'\bgreen tech', r'\bhydrogen fuel'],
    "Hypertext":        [r'\bhypertext\b', r'\bhyperlink', r'\bweb page'],
    "Internet":         [r'\binternet\b', r'\bonline\b', r'\bworld wide web\b', r'\bbroadband', r'\bnetwork\b',
                         r'\bdot\.com\b', r'\bdotcom\b', r'\bdot com\b', r'\bdot-com\b'],
    "Chemistry":        [r'\bchemi', r'\breaction\b', r'\bcatalys', r'\bcompound\b', r'\bsynthes'],
    "Physics":          [r'\bphysics\b', r'\bphysical\b', r'\bquantum\b', r'\bthermodynamic', r'\bmechanics\b', r'\belectromagnet'],
    "Engineering":      [r'\bengineering\b'],
    "Nanotech":         [r'\bnanotech\w*'],
    "Nano":             [r'\bnano\w*'],
    "Molecular Manufacturing": [
        r'\bATP synthase\b', r'\bmolecular motor', r'\bmolecular machine',
        r'\bmolecular assembl', r'\bmolecular manufactur', r'\bconveyor',
        r'\bself-assembl', r'\bmolecular machinery\b', r'\bbiological machine',
        r'\bbiological motor', r'\bmolecular chaperones?\b', r'\bconveyor belt',
        r'\brobot arm', r'\bbiological nanomachin', r"nature's assembler",
        r'\brobotic assembly\b', r'\bmolecular robot',
    ],
    "Revolution":       [r'\brevolution', r'\brevolutionary\b', r'\bparadigm shift', r'\bbreakthrough', r'\btransformative\b'],
    "Financial":        [r'\$', r'\bdollar', r'\bmoney\b', r'\bfunding\b', r'\binvest', r'€', r'£', r'¥'],
    "Commerce":         [r'\bcommerce', r'\bcommercial', r'\bmarket', r'\btrade'],
    "Application":      [r'\bproduct', r'\bproduced', r'\bdevice', r'\bapplication', r'\bmanufactur'],
}


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def extract_date(text):
    """Find the first valid date in text."""
    for pattern, fmt in DATE_PATTERNS:
        for match in re.finditer(pattern, text, re.IGNORECASE):
            candidate = match.group(1)
            parts = re.split(r'[\s,]+', candidate)
            has_month = any(p.lower() in MONTH_NAMES for p in parts)
            if not has_month and fmt != '%Y-%m-%d':
                continue
            try:
                return datetime.strptime(candidate, fmt)
            except ValueError:
                pass
            # Handle odd capitalization like "JUne" → "June"
            try:
                return datetime.strptime(candidate.title(), fmt)
            except ValueError:
                continue
    return None


def count_keyword(text, patterns):
    return sum(len(re.findall(p, text, re.IGNORECASE)) for p in patterns)


def count_words(text):
    return len(text.split())


def get_non_blank_lines(text):
    return [l.strip() for l in text.splitlines() if l.strip()]


def strip_references(text):
    """Split text into (body, references) at the first references heading."""
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if REFERENCES_START.match(line.strip()):
            return "\n".join(lines[:i]).strip(), "\n".join(lines[i:]).strip()
    return text, ""


def in_year_range(year):
    if year is None:
        return True
    return YEAR_MIN <= year <= YEAR_MAX


def _compute_date_match(original, scraped):
    """Compare two dates and return a match percentage."""
    if original is None or scraped is None:
        return None
    if original.date() == scraped.date():
        return 100
    if original.year == scraped.year and original.month == scraped.month:
        return 75
    if original.year == scraped.year:
        return 50
    return 0


def make_row(original_date, scraped_date, source_num, source_name, title, body, refs,
             original_year=None):
    """Build a single output row dict."""
    # Use reference year when available, otherwise derive from dates
    if original_year is not None:
        year = original_year
    elif original_date:
        year = original_date.year
    elif scraped_date:
        year = scraped_date.year
    else:
        year = None
    row = {
        'Original Date': original_date.strftime('%d %B %Y') if original_date else None,
        'Scraped Date':  scraped_date.strftime('%d %B %Y') if scraped_date else None,
        'Date Match %':  _compute_date_match(original_date, scraped_date),
        'Year':          year,
        'Sources':       source_num,
        'Name':          source_name,
        'Word count':    count_words(body),
        'Title':         title,
        'Body':          body,
        'References':    refs,
    }
    for topic, patterns in TOPIC_KEYWORDS.items():
        row[topic] = count_keyword(body, patterns)
    # Combined column for electronics/computing/semiconductors
    row['Total Electronics/Computing'] = (
        row.get('Computers/Computing', 0) + row.get('Semiconductors', 0) + row.get('Electronics', 0)
    )
    return row


def _body_after_title(chunk, title):
    """Return the text in chunk after the first line matching title."""
    raw_lines = chunk.splitlines()
    for i, line in enumerate(raw_lines):
        if line.strip() == title:
            return "\n".join(raw_lines[i + 1:]).strip()
    return chunk


# ═════════════════════════════════════════════════════════════════════════════
# FORMAT-SPECIFIC PARSERS
# ═════════════════════════════════════════════════════════════════════════════

def _parse_ref_date(val):
    """Parse a single date value from the reference worksheet."""
    if pd.isna(val):
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, str):
        fixed = val.replace('0ct', 'Oct').replace('0CT', 'OCT')
        try:
            return pd.to_datetime(fixed).to_pydatetime()
        except Exception:
            print(f"  WARNING: Could not parse date string: '{val}'")
            return None
    return None


def load_reference_data(script_dir, source_num):
    """Load dates and years from the reference worksheet for a given source.
    Returns (dates_list, years_list) where each is positionally indexed."""
    ref_path = script_dir / REFERENCE_WORKSHEET
    if not ref_path.exists():
        return [], []
    try:
        df = pd.read_excel(ref_path)
        src_df = df[df['Sources'] == source_num].reset_index(drop=True)
        dates = [_parse_ref_date(val) for val in src_df['Date']]
        years = []
        for val in src_df['Year']:
            if pd.isna(val):
                years.append(None)
            else:
                years.append(int(val))
        return dates, years
    except Exception as e:
        print(f"  WARNING: Could not load reference data: {e}")
        return [], []


def load_reference_dates(script_dir, source_num):
    """Load dates from the reference worksheet (backward-compatible wrapper)."""
    dates, _ = load_reference_data(script_dir, source_num)
    return dates


def align_by_date(scraped_dates, ref_dates, ref_years):
    """Align parsed articles to reference rows using date-group matching.

    Groups both sequences by date, then matches within each group by order.
    Returns (aligned_dates, aligned_years) lists parallel to scraped_dates.
    """
    from collections import defaultdict

    # Build reference groups: date -> [(ref_idx, ref_date, ref_year), ...]
    ref_groups = defaultdict(list)
    for i, (rd, ry) in enumerate(zip(ref_dates, ref_years)):
        if rd is not None:
            ref_groups[rd.date()].append((i, rd, ry))

    aligned_dates = [None] * len(scraped_dates)
    aligned_years = [None] * len(scraped_dates)

    # Track how many articles we've consumed from each date group
    group_cursors = defaultdict(int)

    for out_idx, sd in enumerate(scraped_dates):
        if sd is None:
            continue
        key = sd.date()
        cursor = group_cursors[key]
        group = ref_groups.get(key, [])
        if cursor < len(group):
            _, ref_date, ref_year = group[cursor]
            aligned_dates[out_idx] = ref_date
            aligned_years[out_idx] = ref_year
            group_cursors[key] = cursor + 1
        else:
            # No more ref rows for this date; use scraped date's year
            aligned_years[out_idx] = sd.year

    return aligned_dates, aligned_years


def parse_government(content, source_name, source_num, ref_dates=None, ref_years=None):
    """Government: asterisk separators, strip 'Article N' prefix from title.
    Dates come from the reference worksheet (articles have no inline dates)."""
    chunks = [c.strip() for c in SEP_ASTERISKS.split(content) if c.strip()]
    rows = []

    for idx, chunk in enumerate(chunks):
        original_date = ref_dates[idx] if ref_dates and idx < len(ref_dates) else None
        original_year = ref_years[idx] if ref_years and idx < len(ref_years) else None
        scraped_date = extract_date(chunk)
        date = original_date or scraped_date
        year = original_year or (date.year if date else None)

        # Skip year filter when using reference dates (reference is authoritative)
        if not ref_dates and not in_year_range(year):
            continue

        lines = get_non_blank_lines(chunk)
        if not lines:
            continue

        # Skip "Article N" prefix line
        start = 0
        if re.match(r'^Article\s+\d+\s*$', lines[0], re.IGNORECASE):
            start = 1

        if start >= len(lines):
            continue

        title = lines[start]
        body_text = _body_after_title(chunk, title)
        body, refs = strip_references(body_text)

        # Fix: if body is empty but title is very long, the body was pasted into the title
        if count_words(body) == 0 and count_words(title) > 20:
            body = title
            title = title[:80].rsplit(' ', 1)[0] + '...'

        rows.append(make_row(original_date, scraped_date, source_num, source_name, title, body, refs,
                             original_year=original_year))

    return rows


def parse_after_label(content, source_name, source_num, ref_dates=None, ref_years=None):
    """Science Research / Science News: section label precedes title."""
    chunks = [c.strip() for c in SEP_ASTERISKS.split(content) if c.strip()]
    rows, skipped = [], 0

    for idx, chunk in enumerate(chunks):
        scraped_date = extract_date(chunk)
        original_date = ref_dates[idx] if ref_dates and idx < len(ref_dates) else None
        original_year = ref_years[idx] if ref_years and idx < len(ref_years) else None
        date = scraped_date or original_date
        year = date.year if date else None
        if not in_year_range(year):
            skipped += 1
            continue

        lines = get_non_blank_lines(chunk)
        if not lines:
            continue

        # Find title: line after a section label
        title = ""
        for i, line in enumerate(lines):
            if SECTION_LABELS.match(line) and i + 1 < len(lines):
                candidate = lines[i + 1]
                if not SCIENCE_META.match(candidate) and not re.match(r'^\d+$', candidate):
                    title = candidate
                    break

        # Fallback: first substantive non-metadata line
        if not title:
            for line in lines:
                if (not SCIENCE_META.match(line)
                        and not re.match(r'^\d+$', line)
                        and len(line) > 3):
                    title = line
                    break

        if not title:
            continue

        # Body after title, strip leading science metadata
        body_text = _body_after_title(chunk, title)
        body_lines = body_text.splitlines()
        while body_lines and (not body_lines[0].strip()
                              or SCIENCE_META.match(body_lines[0].strip())):
            body_lines.pop(0)

        body_text = "\n".join(body_lines).strip()
        body, refs = strip_references(body_text)

        # Fix: if body is empty but title is very long, the content ended up as the title
        if count_words(body) < 5 and count_words(title) > 20:
            body = title
            title = title[:80].rsplit(' ', 1)[0] + '...'

        if count_words(body) < 5:
            continue

        rows.append(make_row(original_date, scraped_date, source_num, source_name, title, body, refs,
                             original_year=original_year))

    if skipped:
        print(f"  (skipped {skipped} articles outside {YEAR_MIN}–{YEAR_MAX})")
    return rows


def _strip_business_tail(lines):
    """Remove trailing junk (Document ID, 'More Like This', contact blocks)."""
    while lines:
        last = lines[-1].strip()
        if not last or BUSINESS_TAIL_JUNK.match(last):
            lines.pop()
        else:
            break
    # Strip trailing contact block (has both phone AND email on same line)
    if lines and (re.search(r'\d{3}[.-]\d{3}[.-]\d{4}', lines[-1])
                  and re.search(r'\b\S+@\S+\.\S+', lines[-1])):
        lines.pop()
    return lines


def parse_business(content, source_name, source_num, ref_dates=None, ref_years=None):
    """Business Press / Business: 'Article N ****' separator, real headline as title."""
    chunks = [c.strip() for c in SEP_BUSINESS.split(content) if c.strip()]
    rows, skipped = [], 0

    for idx, chunk in enumerate(chunks):
        scraped_date = extract_date(chunk)
        original_date = ref_dates[idx] if ref_dates and idx < len(ref_dates) else None
        original_year = ref_years[idx] if ref_years and idx < len(ref_years) else None
        date = scraped_date or original_date
        year = date.year if date else None
        if not in_year_range(year):
            skipped += 1
            continue

        lines = get_non_blank_lines(chunk)
        if not lines:
            continue

        # Skip known header junk at start of chunk
        while lines and BUSINESS_HEADER_JUNK.match(lines[0]):
            lines.pop(0)
        if not lines:
            continue

        title = lines[0]  # actual headline

        # Body starts after copyright line
        raw_lines = chunk.splitlines()
        body_start = 0
        for i, line in enumerate(raw_lines):
            if COPYRIGHT_LINE.match(line.strip()):
                body_start = i + 1
                break

        if body_start == 0:
            # Fallback: skip metadata block (~first 8 lines)
            body_start = min(8, len(raw_lines))

        body_lines = list(raw_lines[body_start:])
        body_lines = _strip_business_tail(body_lines)
        body_text = "\n".join(body_lines).strip()
        body, refs = strip_references(body_text)

        if count_words(body) < 5:
            continue

        rows.append(make_row(original_date, scraped_date, source_num, source_name, title, body, refs,
                             original_year=original_year))

    if skipped:
        print(f"  (skipped {skipped} articles outside {YEAR_MIN}–{YEAR_MAX})")
    return rows


def parse_futurist(content, source_name, source_num, ref_dates=None, ref_years=None):
    """Futurists: split by ToC lines, asterisk separators, and dash separators."""
    combined_sep = re.compile(
        r'(?:^\s*Foresight Update \d+\s*-\s*Table of Contents.*$'
        r'|^\s*\*{19,}\s*$'
        r'|^\s*-{40,}\s*$)',
        re.MULTILINE
    )
    chunks = [c.strip() for c in combined_sep.split(content) if c.strip()]

    # Lines that appear in issue headers (not article content)
    header_line_re = re.compile(
        r'^(A publication of the Foresight Institute|'
        r'Preparing for future technologies|'
        r'Board of Directors|'
        r'All Rights Reserved|'
        r'Write to the Foresight Institute|'
        r'If you find information|'
        r'Editor\s|Publisher\s|'
        r'.{0,30}(President|Secretary|Treasurer)\s*$|'
        r'Box \d+.*CA\s+\d|'
        r'.{0,3}Copyright\s+\d{4}|'
        r'Foresight Institute\s*$|'
        r'\d{1,2}\s+\w+\s+\d{4}\s*$)',
        re.IGNORECASE
    )

    # Whole chunks to skip entirely
    skip_chunk_re = re.compile(
        r'^(Clippings Invited|'
        r'If you find information and clippings|'
        r'Write to the Foresight Institute)',
        re.IGNORECASE
    )

    # Detect issue header chunks to reset date propagation
    issue_header_re = re.compile(
        r'^\s*A publication of the Foresight Institute',
        re.IGNORECASE
    )

    rows, skipped = [], 0
    last_date = None
    article_idx = 0

    for chunk in chunks:
        lines = get_non_blank_lines(chunk)
        if not lines:
            continue

        # Reset last_date at issue boundaries to prevent stale date propagation
        if issue_header_re.match(lines[0]):
            last_date = None

        # Always try to capture date (even from skipped header chunks)
        chunk_date = extract_date(chunk)
        if chunk_date:
            last_date = chunk_date

        # Skip known boilerplate chunks
        if any(skip_chunk_re.match(l) for l in lines[:2]):
            continue
        if count_words(chunk) < 20:
            continue

        # Strip issue-header boilerplate from chunk start (only when
        # the chunk actually begins with a known header line)
        if lines and header_line_re.match(lines[0]):
            while lines:
                if header_line_re.match(lines[0]):
                    lines.pop(0)
                elif len(lines[0].split()) <= 2:
                    lines.pop(0)  # short name lines in header block
                elif lines[0][0].islower():
                    lines.pop(0)  # continuation of previous header paragraph
                else:
                    break

        if not lines or count_words(' '.join(lines)) < 15:
            continue

        scraped_date = chunk_date if chunk_date else last_date
        year = scraped_date.year if scraped_date else None
        if not in_year_range(year):
            skipped += 1
            continue

        original_date = ref_dates[article_idx] if ref_dates and article_idx < len(ref_dates) else None
        original_year = ref_years[article_idx] if ref_years and article_idx < len(ref_years) else None
        title = lines[0]
        body = _body_after_title(chunk, title)
        rows.append(make_row(original_date, scraped_date, source_num, source_name, title, body, "",
                             original_year=original_year))
        article_idx += 1

    if skipped:
        print(f"  (skipped {skipped} articles outside {YEAR_MIN}–{YEAR_MAX})")
    return rows


def parse_newspaper(content, source_name, source_num, ref_dates=None, ref_years=None):
    """Newspapers: split by 'Article N ****' and '* * *' sub-item separators.
    Uses date-group alignment instead of positional index for reference matching."""
    chunks = [c.strip() for c in SEP_NEWSPAPER.split(content) if c.strip()]
    skipped = 0
    last_date = None

    # First pass: parse all chunks into preliminary article data
    parsed = []
    for idx, chunk in enumerate(chunks):
        scraped_date = extract_date(chunk)
        if scraped_date:
            last_date = scraped_date
        else:
            scraped_date = last_date

        date = scraped_date
        year = date.year if date else None
        if not in_year_range(year):
            skipped += 1
            continue

        lines = get_non_blank_lines(chunk)
        if not lines:
            continue

        # Skip known header junk
        while lines and BUSINESS_HEADER_JUNK.match(lines[0]):
            lines.pop(0)
        if not lines:
            continue

        title = lines[0]

        # Body starts after copyright line if present
        raw_lines = chunk.splitlines()
        body_start = 0
        for i, line in enumerate(raw_lines):
            if COPYRIGHT_LINE.match(line.strip()):
                body_start = i + 1
                break

        if body_start == 0:
            body_text = _body_after_title(chunk, title)
        else:
            body_lines = list(raw_lines[body_start:])
            body_lines = _strip_business_tail(body_lines)
            body_text = "\n".join(body_lines).strip()

        body, refs = strip_references(body_text)

        if count_words(body) < 5:
            continue

        parsed.append({
            'scraped_date': scraped_date,
            'title': title,
            'body': body,
            'refs': refs,
        })

    # Second pass: align with reference data by date groups
    scraped_dates = [p['scraped_date'] for p in parsed]
    if ref_dates and ref_years:
        aligned_dates, aligned_years = align_by_date(scraped_dates, ref_dates, ref_years)
    else:
        aligned_dates = [None] * len(parsed)
        aligned_years = [None] * len(parsed)

    # Build final rows with aligned reference data
    rows = []
    for i, p in enumerate(parsed):
        rows.append(make_row(
            aligned_dates[i], p['scraped_date'], source_num, source_name,
            p['title'], p['body'], p['refs'],
            original_year=aligned_years[i],
        ))

    if skipped:
        print(f"  (skipped {skipped} articles outside {YEAR_MIN}–{YEAR_MAX})")
    return rows


# ═════════════════════════════════════════════════════════════════════════════
# DISPATCHER
# ═════════════════════════════════════════════════════════════════════════════

def strip_rtf(text):
    """Strip RTF markup, returning plain text."""
    # Remove RTF header up to first empty line or content
    text = re.sub(r'\{\\fonttbl[^}]*\}', '', text)
    text = re.sub(r'\{\\colortbl[^}]*\}', '', text)
    text = re.sub(r'\{\\rtf1[^}]*?\n', '', text)
    # Remove RTF control words (e.g., \f0, \fs24, \cf0, \pard, etc.)
    text = re.sub(r'\\[a-z]+\d*\s?', ' ', text)
    # Remove remaining braces
    text = text.replace('{', '').replace('}', '')
    # Convert RTF line breaks (\) to newlines
    text = text.replace('\\\n', '\n')
    # Clean up extra whitespace
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n[ \t]+', '\n', text)
    return text.strip()


def parse_articles(filepath, source_name, source_num, title_style):
    """Route to the appropriate parser."""
    script_dir = filepath.parent if filepath.suffix != '.rtf' else filepath.parent.parent

    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        content = f.read()
    content = content.replace('\r\n', '\n').replace('\r', '\n')

    # Strip RTF markup if needed
    if filepath.suffix == '.rtf':
        content = strip_rtf(content)

    # Skip reference dates for supplementary files (e.g., Business_2005.rtf)
    # whose articles don't align with the main reference worksheet
    if filepath.suffix == '.rtf':
        ref_dates, ref_years = [], []
    else:
        ref_dates, ref_years = load_reference_data(script_dir, source_num)
    if ref_dates:
        print(f"  (loaded {len(ref_dates)} reference entries from {REFERENCE_WORKSHEET})")

    if title_style == "government":
        return parse_government(content, source_name, source_num, ref_dates, ref_years)
    elif title_style == "after_label":
        return parse_after_label(content, source_name, source_num, ref_dates, ref_years)
    elif title_style == "business_press":
        return parse_business(content, source_name, source_num, ref_dates, ref_years)
    elif title_style == "futurist":
        return parse_futurist(content, source_name, source_num, ref_dates, ref_years)
    elif title_style == "newspaper":
        return parse_newspaper(content, source_name, source_num, ref_dates, ref_years)
    else:
        raise ValueError(f"Unknown title_style: {title_style}")


# ═════════════════════════════════════════════════════════════════════════════
# EXCEL OUTPUT
# ═════════════════════════════════════════════════════════════════════════════

def write_excel(all_rows, output_path):
    columns = (
        ['Original Date', 'Scraped Date', 'Date Match %', 'Year', 'Sources', 'Name',
         'Word count', 'Title', 'Body', 'References']
        + list(TOPIC_KEYWORDS.keys())
        + ['Total Electronics/Computing']
    )
    df = pd.DataFrame(all_rows, columns=columns)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        ws = writer.sheets['Sheet1']

        header_fill = PatternFill('solid', start_color='4472C4', end_color='4472C4')
        header_font = Font(name='Arial', bold=True, color='FFFFFF', size=11)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')

        name_widths = {
            'Original Date': 14, 'Scraped Date': 14, 'Date Match %': 12,
            'Year': 6, 'Sources': 8, 'Name': 18, 'Word count': 12,
            'Title': 40, 'Body': 60, 'References': 60,
        }
        all_cols = [chr(65 + i) if i < 26 else 'A' + chr(65 + i - 26)
                    for i in range(ws.max_column)]
        for col_idx, col_letter in enumerate(all_cols, start=1):
            col_name = df.columns[col_idx - 1]
            ws.column_dimensions[col_letter].width = name_widths.get(col_name, 16)

        data_font = Font(name='Arial', size=11)
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = data_font
                cell.alignment = Alignment(wrap_text=False, vertical='top')


# ═════════════════════════════════════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════════════════════════════════════

def main():
    script_dir = Path(__file__).parent
    all_rows = []

    for filename, (source_name, source_num, title_style) in FILE_SOURCE_MAP.items():
        filepath = script_dir / filename
        if not filepath.exists():
            print(f"WARNING: '{filename}' not found in {script_dir} — skipping.")
            continue

        rows = parse_articles(filepath, source_name, source_num, title_style)
        all_rows.extend(rows)
        print(f"'{filename}' ({source_name}, Sources={source_num}): {len(rows)} articles")
        # Print a few samples for spot-checking
        for r in rows[:3]:
            safe_title = r['Title'][:60].encode('ascii', 'replace').decode('ascii')
            print(f"  Sample: Scraped={r['Scraped Date']}, Original={r['Original Date']}, "
                  f"Match={r['Date Match %']}, Words={r['Word count']}, "
                  f"Title='{safe_title}'")

    if not all_rows:
        print("No articles found. Check that your source files are in the same folder.")
        return

    # ── DEDUPLICATION ───────────────────────────────────────────────────────
    before_count = len(all_rows)
    df_temp = pd.DataFrame(all_rows)
    df_temp = df_temp.drop_duplicates(subset=['Title', 'Body', 'Sources'], keep='first')
    all_rows = df_temp.to_dict('records')
    removed = before_count - len(all_rows)
    if removed:
        print(f"\nRemoved {removed} exact duplicate rows ({before_count} -> {len(all_rows)})")

    # ── DATE MATCH REPORTING (exclude Government & Futurists) ───────────────
    print("\n-- Date Match % by Source --")
    for src_num, src_name in [(2, 'Science News'), (3, 'Science Research'),
                               (4, 'Business Press'), (5, 'Business'), (7, 'Newspapers')]:
        src_data = [r['Date Match %'] for r in all_rows
                    if r['Sources'] == src_num and r['Date Match %'] is not None
                    and not (isinstance(r['Date Match %'], float) and r['Date Match %'] != r['Date Match %'])]
        if src_data:
            avg = sum(src_data) / len(src_data)
            zeros = sum(1 for v in src_data if v == 0)
            print(f"  {src_name}: Avg={avg:.1f}% ({len(src_data)} rows)")
            if zeros:
                print(f"    WARNING: {zeros} rows with 0% date match")

    output_path = script_dir / OUTPUT_FILE
    write_excel(all_rows, output_path)
    print(f"\nDone! {len(all_rows)} total rows written to '{output_path}'")


if __name__ == '__main__':
    main()
