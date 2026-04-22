"""Generate a PDF report summarizing the second round of pipeline improvements.

Covers six items from next_steps.txt completed on 2026-04-22:
  1. Days Between + tolerant date-match columns
  2. Auto-populated Qualitative Samples + AI Qualitative Samples sheets
  3. Coverage Report sheet explaining the 12,774 -> 12,206 article-count delta
  4. Multi-page PDF export of notebook visualizations (full + summary)
  5. Opt-in Claude API qualitative-coding module
  6. Newspaper separator-regex fix (recovered +95 missing articles)
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from fpdf import FPDF
from pathlib import Path

OUT_DIR = Path(__file__).parent
IMG_DIR = OUT_DIR / "_report_imgs"
IMG_DIR.mkdir(exist_ok=True)


def save_fig(fig, name):
    path = IMG_DIR / f"{name}.png"
    fig.savefig(path, dpi=180, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return str(path)


# ── Chart 1: strict vs tolerant Date Match % by source ─────────────────────

def chart_strict_vs_tolerant():
    sources  = ['Science News', 'Sci. Research', 'Bus. Press', 'Business', 'Newspapers']
    strict   = [94.2, 99.8, 100.0, 99.1, 100.0]
    tolerant = [92.6, 99.8, 100.0, 99.7, 100.0]

    x = np.arange(len(sources))
    width = 0.36

    fig, ax = plt.subplots(figsize=(8, 3.8))
    b1 = ax.bar(x - width/2, strict,   width, color='#94A3B8', label='Strict (year-bucket)', edgecolor='white')
    b2 = ax.bar(x + width/2, tolerant, width, color='#1D4ED8', label='Tolerant (linear decay, 90d)', edgecolor='white')

    for bars in (b1, b2):
        for bar in bars:
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.4,
                    f'{bar.get_height():.1f}', ha='center', va='bottom', fontsize=8)

    ax.set_xticks(x)
    ax.set_xticklabels(sources, fontsize=9)
    ax.set_ylabel('Avg Date Match %', fontsize=10)
    ax.set_title('Date Match % by Source: Strict vs Tolerant', fontsize=13, fontweight='bold', pad=10)
    ax.set_ylim(85, 102)
    ax.legend(fontsize=9, loc='lower right')
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    return save_fig(fig, 'strict_vs_tolerant')


# ── Chart 2: 0%-rows eliminated by the tolerant metric ─────────────────────

def chart_zero_rows_reduction():
    sources = ['Science News', 'Business']
    before  = [28, 1]
    after   = [9, 0]

    x = np.arange(len(sources))
    width = 0.36

    fig, ax = plt.subplots(figsize=(6, 3.5))
    ax.bar(x - width/2, before, width, color='#DC2626', label='Strict', edgecolor='white')
    ax.bar(x + width/2, after,  width, color='#16A34A', label='Tolerant', edgecolor='white')

    for i, (b, a) in enumerate(zip(before, after)):
        ax.text(i - width/2, b + 0.6, str(b), ha='center', va='bottom', fontsize=10, fontweight='bold')
        ax.text(i + width/2, a + 0.6, str(a), ha='center', va='bottom', fontsize=10, fontweight='bold')

    ax.set_xticks(x)
    ax.set_xticklabels(sources, fontsize=10)
    ax.set_ylabel('Rows with 0% date match', fontsize=10)
    ax.set_title('0%-Match Rows: eliminated by tolerant metric', fontsize=12, fontweight='bold', pad=10)
    ax.legend(fontsize=9, loc='upper right')
    ax.set_ylim(0, 34)
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    return save_fig(fig, 'zero_rows_reduction')


# ── Chart 3: Coverage breakdown — reference vs raw chunks vs final ─────────

def chart_coverage_breakdown():
    sources    = ['Government', 'Sci. News', 'Sci. Research', 'Bus. Press',
                  'Business', 'Futurists', 'Newspapers']
    reference  = [926, 1407, 1102, 494, 4157, 926, 3762]
    raw_chunks = [926, 1235, 1102, 494, 4157, 1065, 3469]  # Newspapers post-fix
    final      = [906, 1201, 1079, 491, 4154, 926, 3449]   # Newspapers post-fix

    x = np.arange(len(sources))
    width = 0.28

    fig, ax = plt.subplots(figsize=(10, 4))
    ax.bar(x - width, reference,  width, color='#94A3B8', label='Reference',  edgecolor='white')
    ax.bar(x,         raw_chunks, width, color='#F59E0B', label='Raw chunks', edgecolor='white')
    ax.bar(x + width, final,      width, color='#1D4ED8', label='Final',      edgecolor='white')

    ax.set_xticks(x)
    ax.set_xticklabels(sources, fontsize=9, rotation=20, ha='right')
    ax.set_ylabel('Article count', fontsize=10)
    ax.set_title('Coverage Breakdown: Reference -> Raw Chunks -> Final',
                 fontsize=13, fontweight='bold', pad=10)
    ax.legend(fontsize=9, loc='upper right')
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)

    # Annotate the two sources where raw_chunks < reference (the separator gap)
    for i, (ref, rc) in enumerate(zip(reference, raw_chunks)):
        if rc < ref:
            gap = ref - rc
            ax.annotate(f'-{gap}', xy=(x[i], rc), xytext=(x[i], rc - 280),
                        ha='center', fontsize=9, color='#DC2626', fontweight='bold',
                        arrowprops=dict(arrowstyle='->', color='#DC2626', lw=1.2))
    return save_fig(fig, 'coverage_breakdown')


# ── Chart 4: Coverage Report detail table ──────────────────────────────────

def chart_coverage_table():
    rows = [
        ['Source', 'Ref', 'Raw', 'Filter', 'Dedup', 'Final', 'Delta'],
        ['Government',       '926',  '926',  '0',  '20', '906',  '-20'],
        ['Science News',     '1407', '1235', '23', '11', '1201', '-206'],
        ['Science Research', '1102', '1102', '19', '4',  '1079', '-23'],
        ['Business Press',   '494',  '494',  '0',  '3',  '491',  '-3'],
        ['Business',         '4157', '4157', '0',  '3',  '4154', '-3'],
        ['Futurists*',       '926',  '1065', '38', '101','926',  '0'],
        ['Newspapers',       '3762', '3469', '5',  '15', '3449', '-313'],
        ['TOTAL',            '12774','12448','85', '157','12206','-568'],
    ]

    fig, ax = plt.subplots(figsize=(7.5, 3.5))
    ax.axis('off')

    table = ax.table(cellText=rows, loc='center', cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.35)

    # Header row
    for j in range(len(rows[0])):
        table[0, j].set_facecolor('#1D4ED8')
        table[0, j].set_text_props(color='white', fontweight='bold')

    # Body rows
    for i in range(1, len(rows)):
        is_total = (i == len(rows) - 1)
        for j in range(len(rows[0])):
            if is_total:
                table[i, j].set_facecolor('#E2E8F0')
                table[i, j].set_text_props(fontweight='bold')
            else:
                table[i, j].set_facecolor('#F8FAFC')

        # Highlight sources where raw < ref (the separator-coverage issue)
        delta = int(rows[i][6]) if not is_total else None
        if not is_total:
            ref = int(rows[i][1])
            raw = int(rows[i][2])
            if raw < ref:
                table[i, 2].set_facecolor('#FEE2E2')  # Raw column in pale red

    ax.set_title('Coverage Report (as written to output.xlsx)',
                 fontsize=12, fontweight='bold', pad=10)
    return save_fig(fig, 'coverage_table')


# ── Chart 5b: Newspaper separator fix — before/after ──────────────────────

def chart_separator_recovery():
    labels = ['Before fix', 'After fix']
    chunks = [3374, 3469]
    final  = [3354, 3449]

    x = np.arange(len(labels))
    width = 0.36

    fig, ax = plt.subplots(figsize=(6.5, 3.5))
    b1 = ax.bar(x - width/2, chunks, width, color='#F59E0B', label='Raw chunks', edgecolor='white')
    b2 = ax.bar(x + width/2, final,  width, color='#1D4ED8', label='Final (post-dedup)', edgecolor='white')

    for bar, val in zip(b1, chunks):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 15,
                str(val), ha='center', va='bottom', fontsize=10, fontweight='bold')
    for bar, val in zip(b2, final):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 15,
                str(val), ha='center', va='bottom', fontsize=10, fontweight='bold')

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10)
    ax.set_ylabel('Newspaper article count', fontsize=10)
    ax.set_title('Newspaper separator fix: +95 articles recovered',
                 fontsize=12, fontweight='bold', pad=10)
    ax.set_ylim(3200, 3600)
    ax.legend(fontsize=9, loc='lower right')
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    return save_fig(fig, 'separator_recovery')


# ── Chart 5: PDF export output sizes ───────────────────────────────────────

def chart_pdf_output():
    labels = ['Summary\n(12 figures)', 'Full archive\n(218 figures)']
    sizes_kb = [89, 905]
    colors = ['#16A34A', '#1D4ED8']

    fig, ax = plt.subplots(figsize=(5.5, 3.3))
    bars = ax.bar(labels, sizes_kb, color=colors, width=0.5, edgecolor='white', linewidth=1.2)
    for bar, size in zip(bars, sizes_kb):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 15,
                f'{size} KB', ha='center', va='bottom', fontsize=11, fontweight='bold')
    ax.set_ylabel('File size (KB)', fontsize=10)
    ax.set_title('Notebook Visualization PDFs', fontsize=12, fontweight='bold', pad=10)
    ax.set_ylim(0, 1100)
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    return save_fig(fig, 'pdf_output')


# ── PDF class ──────────────────────────────────────────────────────────────

class Report(FPDF):
    def header(self):
        if self.page_no() > 1:
            self.set_font('Helvetica', 'I', 8)
            self.set_text_color(120, 120, 120)
            self.cell(0, 8, 'Nanotech Research Pipeline - Round 2 Update', align='R')
            self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', align='C')

    def section_title(self, title):
        self.set_font('Helvetica', 'B', 14)
        self.set_text_color(29, 78, 216)
        self.cell(0, 10, title)
        self.ln(8)
        self.set_draw_color(29, 78, 216)
        self.set_line_width(0.4)
        self.line(self.get_x(), self.get_y(), self.get_x() + 190, self.get_y())
        self.ln(6)

    def subheading(self, label):
        self.set_font('Helvetica', 'B', 12)
        self.set_text_color(60, 60, 60)
        self.cell(0, 8, label)
        self.ln(8)

    def body_text(self, text):
        self.set_font('Helvetica', '', 11)
        self.set_text_color(30, 30, 30)
        self.multi_cell(0, 6, text)
        self.ln(2)

    def bullet(self, text):
        self.set_font('Helvetica', '', 11)
        self.set_text_color(30, 30, 30)
        self.cell(6, 6, '-')
        self.multi_cell(0, 6, text)
        self.ln(1)

    def mono(self, text):
        self.set_font('Courier', '', 10)
        self.set_text_color(30, 30, 30)
        self.set_fill_color(243, 244, 246)
        self.multi_cell(0, 5.5, text, fill=True)
        self.ln(2)


def build_report():
    img_strict_tolerant = chart_strict_vs_tolerant()
    img_zero_rows       = chart_zero_rows_reduction()
    img_coverage_bars   = chart_coverage_breakdown()
    img_coverage_tbl    = chart_coverage_table()
    img_separator_fix   = chart_separator_recovery()
    img_pdf_output      = chart_pdf_output()

    pdf = Report()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)

    # ── TITLE PAGE ─────────────────────────────────────────────────────────
    pdf.add_page()
    pdf.ln(38)
    pdf.set_font('Helvetica', 'B', 28)
    pdf.set_text_color(29, 78, 216)
    pdf.cell(0, 15, 'Pipeline Update Report', align='C')
    pdf.ln(14)
    pdf.set_font('Helvetica', 'B', 18)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 10, 'Round 2', align='C')
    pdf.ln(12)
    pdf.set_font('Helvetica', '', 16)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 10, 'Nanotech Media Analysis Project', align='C')
    pdf.ln(20)
    pdf.set_font('Helvetica', '', 12)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 8, 'April 22, 2026', align='C')
    pdf.ln(6)
    pdf.cell(0, 8, '5 next_steps items + Newspaper separator fix (+95 articles)', align='C')

    # ── OVERVIEW PAGE ──────────────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('Overview')
    pdf.body_text(
        'This round addresses five bullets from next_steps.txt, plus a follow-up '
        'fix the Coverage Report itself surfaced. The work spans the scraper '
        '(create_excelV5.py), the analysis notebook (nanotech_graphs.ipynb), and '
        'three new Python packages: sampling/, nanotech_plots/, and '
        'qualitative_coding/. The "what other fixes are needed?" bullet was '
        'answered by running the diagnostic once, finding the Newspaper separator '
        'gap, and fixing it -- recovering +95 articles and pushing total coverage '
        'from 94.8% to 95.6%.'
    )
    pdf.ln(2)

    pdf.subheading('What was done:')
    items = [
        '1. Added "Days Between" column and a tolerant Date Match % (linear decay to 90 days).',
        '2. Qualitative Samples and AI Qualitative Samples sheets now populate automatically when create_excelV5.py runs.',
        '3. Added a Coverage Report sheet that breaks down the per-source gap between the reference worksheet and output.xlsx.',
        '4. Built nanotech_plots/ and export_figures.py to produce a full archive PDF (218 figures) and a summary PDF (12 figures) from the notebook visualizations.',
        '5. Built qualitative_coding/: opt-in module that uses Claude (sonnet-4-6) to code sampled articles on four dimensions. Requires ANTHROPIC_API_KEY.',
        '6. Fixed SEP_NEWSPAPER to catch four Factiva boundary variants (pipe separator, Previous/Next Article suffixes, bare Article N). +95 articles recovered.',
    ]
    for item in items:
        pdf.bullet(item)

    pdf.ln(4)
    pdf.subheading('Key finding from the Coverage Report (and the fix):')
    pdf.body_text(
        'The initial -663 gap (12,774 reference -> 12,111 output) was not primarily '
        'caused by filters. Science News and Newspapers together lost 560 articles at '
        'the separator-split step -- the regex was missing whole article boundaries '
        'that exist in the reference. For Newspapers specifically, the raw file uses '
        'four boundary variants (including a pipe-separated form and bare Article-N '
        'markers) that the original regex did not handle. Fixing the regex (item 6) '
        'recovered 95 Newspaper articles and closed the total gap to -568 (95.6% '
        'coverage). The remaining Science News -172 and Newspapers -313 are '
        'articles absent from the raw text files -- not recoverable via parsing.'
    )

    # ── ITEM 1: TOLERANT DATE MATCHING ─────────────────────────────────────
    pdf.add_page()
    pdf.section_title('1. Days Between + tolerant Date Match %')

    pdf.body_text(
        'Two new columns in Sheet1:\n'
        '  - Days Between: absolute day gap between Original Date and Scraped Date.\n'
        '  - Date Match % (tolerant): 100 * max(0, 1 - days/90). Linear decay.\n'
        'The original Date Match % column is unchanged, so existing plots still work.'
    )
    pdf.ln(1)

    pdf.body_text(
        'The motivation: the strict metric snaps to 0 whenever the year differs, '
        'so Jan 5 1995 vs Dec 15 1994 (a 21-day gap across a year boundary) was '
        'scored as 0% -- the same as Dec 2004 vs Dec 1994. The tolerant metric '
        'treats year boundaries the same as any other 3-week gap.'
    )

    pdf.image(img_strict_tolerant, x=15, w=180)
    pdf.ln(2)

    pdf.body_text(
        'The two metrics disagree on Science News (94.2% strict vs 92.6% tolerant). '
        'They are measuring different things: strict gives 90% to any same-year match '
        'regardless of how many months apart the dates are, while tolerant penalizes '
        '3+ month gaps even within the same year. Use whichever matches the question '
        'being asked -- both columns are in Sheet1.'
    )

    pdf.ln(2)
    pdf.image(img_zero_rows, x=50, w=110)
    pdf.ln(2)

    pdf.body_text(
        'The cross-year-boundary rows the user flagged (0% strict despite being within '
        'a few weeks) drop from 28 to 9 for Science News and 1 to 0 for Business under '
        'the tolerant metric.'
    )

    # ── ITEM 2: AUTO QUAL SHEETS ───────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('2. Auto-populated Qualitative Samples sheets')

    pdf.body_text(
        'Previously, the Qualitative Samples (Space/Electronics/Biology) and AI '
        'Qualitative Samples (Nanotech+AI) sheets in output.xlsx only existed if the '
        'user manually opened nanotech_graphs.ipynb and ran cells 24 and 26. That is, '
        'output.xlsx straight out of the parser was incomplete.'
    )

    pdf.body_text(
        'The sampling logic has been extracted into a new sampling/ package (quota.py, '
        'qualitative.py, ai_qualitative.py). create_excelV5.py now calls '
        'write_qualitative_sheets() at the end of main(), so a fresh run produces all '
        'three sheets unattended. Notebook cells 24 and 26 have been rewritten to '
        'import from the same module -- one source of truth, no double-writes.'
    )

    pdf.subheading('Output of a fresh run:')
    pdf.mono(
        "$ python create_excelV5.py\n"
        "...\n"
        "Done! 12111 total rows written to 'output.xlsx'\n"
        "Wrote 264 qualitative + 75 AI-qualitative samples.\n"
        "Wrote Coverage Report (12111/12774 articles, 94.8% coverage)\n"
    )

    pdf.body_text(
        'Sheets in output.xlsx after a clean run: Sheet1, Qualitative Samples, '
        'AI Qualitative Samples, Coverage Report.'
    )

    # ── ITEM 3: COVERAGE REPORT ────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('3. Coverage Report: explaining the -663 gap')

    pdf.body_text(
        "Each parser now accepts a stats dict and tracks, per source: raw_chunks (how "
        "many pieces the separator split the file into), dropped_year, "
        "dropped_wordcount, dropped_empty_title, dropped_empty_chunk, "
        "dropped_boilerplate (Futurists only). Post-dedup counts and the reference "
        "counts are computed in main(). A new Coverage Report sheet summarizes all of "
        "it."
    )

    pdf.image(img_coverage_bars, x=10, w=195)
    pdf.ln(2)

    pdf.image(img_coverage_tbl, x=30, w=150)
    pdf.ln(2)

    pdf.body_text(
        'The gray -> orange step is where chunks go missing (separator coverage); '
        'the orange -> blue step is filters + dedup. Newspapers was -388 short in '
        'the separator step before the item-6 fix; it is now -293 (still short, '
        'but 95 articles were recovered). Science News is still -172 in the '
        'separator step -- but inspection of the raw file shows no alternative '
        'boundary patterns exist, so those articles are not present in the source '
        'text (a data-collection issue, not a parser issue).'
    )

    pdf.subheading('Follow-up surfaced by this diagnostic:')
    pdf.bullet(
        'Science News -172: articles are simply not in Science_news.txt. Either '
        're-scrape those articles from the original source or document the gap. '
        'No regex change will recover them -- only 1234 boundary markers exist '
        'in the file.'
    )
    pdf.bullet(
        'Newspapers -313 (post-fix): same situation. The raw text file contains '
        '3469 article markers; the reference lists 3762. The 293-article residue '
        'is absent from the source text.'
    )
    pdf.bullet(
        'Futurists is marked "(approx)" in the Coverage Report because it uses an '
        'ad-hoc article_idx counter rather than positional matching. Its -0 delta '
        'is misleading -- the raw chunk count (1065) is actually higher than the '
        'reference (926), indicating the Futurists separator is over-splitting. '
        'Worth a follow-up pass to tighten combined_sep in parse_futurist().'
    )

    # ── ITEM 4: PDF EXPORT ─────────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('4. PDF export of notebook visualizations')

    pdf.body_text(
        'Notebook visualizations were disappearing into plt.show() -- they rendered '
        'in Jupyter but were not persisted anywhere. Fixing that required pulling the '
        'plotting code out of the notebook so it could be driven headlessly.'
    )

    pdf.subheading('What was built:')
    pdf.bullet(
        'nanotech_plots/ package: config.py (constants), data.py (load_data, '
        'pct_per_year, moving_avg, smart_yticks helpers), plots.py (one function per '
        'chart type, each returns a matplotlib Figure), registry.py (list of '
        'every plot tagged full/summary).'
    )
    pdf.bullet(
        'export_figures.py: uses matplotlib.backends.backend_pdf.PdfPages to '
        'iterate the registry and write one figure per page. Flags: --full, --summary, '
        'or no flag (writes both).'
    )

    pdf.image(img_pdf_output, x=45, w=120)
    pdf.ln(4)

    pdf.subheading('Run it:')
    pdf.mono(
        "$ python export_figures.py --summary   # ~12 figures, ~90 KB\n"
        "$ python export_figures.py --full      # 218 figures, ~905 KB\n"
        "$ python export_figures.py             # both\n"
    )

    pdf.body_text(
        'Summary PDF curation: the macro co-occurrence plot (all sources), % articles '
        'mentioning nanotech for each source (7), and Top-8 keyword co-occurrence for '
        'the four most-interesting sources (Government, Business, Futurists, '
        'Newspapers). Tag a plot with is_summary=True in registry.py to add it.'
    )

    # ── ITEM 5: CLAUDE API CODING ──────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('5. Claude API qualitative coding (opt-in)')

    pdf.body_text(
        'New qualitative_coding/ package. Reads the Qualitative Samples and AI '
        'Qualitative Samples sheets, asks Claude (default: claude-sonnet-4-6) to '
        'code each article on four dimensions, and writes two new sheets: '
        '"Qualitative Samples (Coded)" and "AI Qualitative Samples (Coded)".'
    )

    pdf.subheading('Coding dimensions (9 output fields):')
    pdf.bullet('Attitude toward nanotech: positive / neutral / negative / mixed, plus a one-sentence rationale.')
    pdf.bullet('Analogy: present? type (short phrase)? temporality (past/present/future/timeless/NA)?')
    pdf.bullet('Sustaining vs disrupting: sustaining / disrupting / both / neither.')
    pdf.bullet('Funding argument: present? stance (pro/anti/descriptive/NA)? one-sentence summary.')

    pdf.subheading('Design decisions:')
    pdf.bullet(
        'Forced tool-use with strict=True on the record_coding tool gives guaranteed-valid '
        'JSON back. No regex parsing of free text.'
    )
    pdf.bullet(
        'System prompt (rubric + 6 few-shot examples, ~2100 tokens) is placed in a single '
        'cache_control: {type: "ephemeral"} block. Every call after the first reads the '
        'prefix for ~0.1x the cost. At 339 total sample rows (264 + 75) that saves the '
        'bulk of the input-token bill.'
    )
    pdf.bullet(
        'On-disk sidecar cache keyed by sha256(title + body). Reruns after a crash do not '
        're-pay the API. The cache also doubles as an audit trail for the methods section.'
    )
    pdf.bullet(
        'Exponential backoff on 429/5xx. Per-row errors are caught and written as '
        'ERROR: <exception> into that row\'s coding columns -- one bad row does not '
        'abort the whole run.'
    )

    pdf.subheading('Setup + run:')
    pdf.mono(
        "$ pip install anthropic\n"
        "$ $env:ANTHROPIC_API_KEY = 'sk-ant-...'      # PowerShell\n"
        "# or\n"
        "$ export ANTHROPIC_API_KEY=sk-ant-...        # bash\n"
        "\n"
        "$ python create_excelV5.py                   # produces samples\n"
        "$ python -m qualitative_coding               # codes them\n"
    )

    pdf.body_text(
        'The module is opt-in: create_excelV5.py never imports or calls it, so day-to-day '
        'parser runs stay offline and free. The API key is only read by '
        'qualitative_coding.client.CodingClient at instantiation, and a missing key '
        'raises a clear RuntimeError with PowerShell/bash setup instructions.'
    )

    pdf.subheading('Cost sketch:')
    pdf.body_text(
        'At ~2100 tokens of cached system prompt + ~500 tokens of per-article user text, '
        '~339 rows, claude-sonnet-4-6 ($3/M input, $15/M output), and ~300 output tokens '
        'per call, the run costs roughly $1.50-$3.00 after the first call writes the '
        'cache. The sidecar cache means reruns cost nothing.'
    )

    # ── ITEM 6: NEWSPAPER SEPARATOR FIX ────────────────────────────────────
    pdf.add_page()
    pdf.section_title('6. Newspaper separator-regex fix')

    pdf.body_text(
        'The Coverage Report (item 3) identified SEP_NEWSPAPER as the single '
        'biggest source of missing articles: -388 chunks vs the reference '
        'worksheet. Inspecting Newspapers_1984-2005.txt revealed four Factiva '
        'boundary variants, only the first of which the regex was catching:'
    )

    pdf.mono(
        '  "Article N ***..."                    3361 lines (caught)\n'
        '  "Article N Previous Article|***..."     87 lines (ORPHAN - pipe separator)\n'
        '  "Article N Previous Article ***..."      8 lines (ORPHAN)\n'
        '  "Article N Next Article"                 5 lines (ORPHAN - no asterisks)\n'
        '  "Article N"  (bare)                      2 lines (ORPHAN)\n'
    )

    pdf.body_text(
        'The pipe-separated variant was the big one. It looks like HTML-export '
        'artifacts from Factiva/LexisNexis where the "Previous Article" navigation '
        'link was separated from the asterisk line by a pipe character.'
    )

    pdf.subheading('Updated SEP_NEWSPAPER (create_excelV5.py line 45):')
    pdf.mono(
        'SEP_NEWSPAPER = re.compile(\n'
        '    r"(?:"\n'
        '    r"^\\s*Article\\s+\\d+"\n'
        '    r"(?:(?:\\s+(?:Previous|Next)\\s+Article)?\\s*\\|?\\s*\\*{10,})?"\n'
        '    r"\\s*$"\n'
        '    r"|^\\s*\\*\\s+\\*\\s+\\*\\s*$"\n'
        '    r")",\n'
        '    re.MULTILINE | re.IGNORECASE,\n'
        ')\n'
    )

    pdf.image(img_separator_fix, x=35, w=130)
    pdf.ln(2)

    pdf.subheading('Validation performed before applying the fix:')
    pdf.bullet('Zero old matches dropped -- existing 3373 boundaries still hit.')
    pdf.bullet('95 new matches added -- all confirmed to be on lines starting with "Article N".')
    pdf.bullet('Min chunk size post-fix: 233 chars. No micro-chunks indicating over-splitting.')
    pdf.bullet('No false positives: no "Article N <body text>" matches, because $ anchor requires end-of-line.')

    pdf.subheading('Impact:')
    pdf.bullet('Newspapers final count: 3354 -> 3449 (+95 articles)')
    pdf.bullet('Newspapers coverage: 89.2% -> 91.7%')
    pdf.bullet('Overall output.xlsx: 12,111 -> 12,206 rows')
    pdf.bullet('Overall coverage: 94.8% -> 95.6%')

    # ── NEXT STEPS ─────────────────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('Next steps')

    pdf.subheading('Remaining concrete follow-ups:')
    pdf.bullet(
        'Science News -172 and Newspapers -313 are articles absent from the raw text '
        'files -- inspect Science_news.txt and Newspapers_1984-2005.txt against the '
        'reference worksheet to confirm which specific articles are missing, then '
        're-scrape from the original source. Not a parser fix.'
    )
    pdf.bullet(
        'Futurists is over-splitting (raw=1065 vs ref=926). Its combined_sep in '
        'parse_futurist() matches three separators, including a dash-run "----------" '
        'that may be firing inside article bodies used for formatting. The '
        'overshoot is currently absorbed by boilerplate and dedup filters -- which '
        'means dates are aligned approximately, not exactly. Worth tightening.'
    )
    pdf.bullet(
        'Calibrate the qualitative-coding output: hand-code 10 articles and compare against '
        "Claude's output. Target >=80% agreement on Attitude and Sustaining/Disrupting "
        '(the more subjective dimensions). If lower, add more few-shots to '
        'qualitative_coding/prompts.py.'
    )
    pdf.bullet(
        "Consider widening the tolerant metric's decay window from 90 days to 180 days if "
        'the 92.6% tolerant score for Science News proves to penalize too many '
        'legitimate same-year-different-month matches. Adjustable in '
        'create_excelV5.py:_compute_date_match_tolerant.'
    )

    pdf.ln(2)
    pdf.subheading('Deferred:')
    pdf.bullet(
        'Community dimension graphs over time (next_steps.txt bullet 2, marked IGNORE). '
        'The Claude coding module produces exactly the per-article codings needed to drive '
        'these graphs -- once it has been run, a new notebook section can read the coded '
        'sheets and plot the requested series.'
    )

    # ── SAVE ───────────────────────────────────────────────────────────────
    output_path = OUT_DIR / 'Pipeline_Update_Report_v2.pdf'
    pdf.output(str(output_path))
    print(f"Report saved to: {output_path}")

    # Clean up temp images
    for f in IMG_DIR.glob('*.png'):
        f.unlink()
    IMG_DIR.rmdir()


if __name__ == '__main__':
    build_report()
