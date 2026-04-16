"""Generate a PDF report summarizing the pipeline improvements."""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
from fpdf import FPDF
from pathlib import Path
import os

OUT_DIR = Path(__file__).parent
IMG_DIR = OUT_DIR / "_report_imgs"
IMG_DIR.mkdir(exist_ok=True)


# ── helper: save a matplotlib figure as a PNG ──────────────────────────────

def save_fig(fig, name):
    path = IMG_DIR / f"{name}.png"
    fig.savefig(path, dpi=180, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return str(path)


# ── Chart 1: Science News article recovery ─────────────────────────────────

def chart_science_news():
    labels = ['Before', 'After']
    values = [858, 1212]
    colors = ['#D97706', '#16A34A']

    fig, ax = plt.subplots(figsize=(5, 3))
    bars = ax.bar(labels, values, color=colors, width=0.5, edgecolor='white', linewidth=1.2)
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 20,
                str(val), ha='center', va='bottom', fontsize=14, fontweight='bold')
    ax.set_ylabel('Articles Parsed', fontsize=11)
    ax.set_title('Science News: Articles Recovered', fontsize=13, fontweight='bold', pad=10)
    ax.set_ylim(0, 1400)
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    return save_fig(fig, 'science_news')


# ── Chart 2: Newspaper date match ──────────────────────────────────────────

def chart_newspaper_dates():
    labels = ['Before', 'After']
    values = [30.0, 100.0]
    colors = ['#DC2626', '#16A34A']

    fig, ax = plt.subplots(figsize=(5, 3))
    bars = ax.bar(labels, values, color=colors, width=0.5, edgecolor='white', linewidth=1.2)
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1.5,
                f'{val:.0f}%', ha='center', va='bottom', fontsize=14, fontweight='bold')
    ax.set_ylabel('Avg Date Match %', fontsize=11)
    ax.set_title('Newspapers: Date Alignment Accuracy', fontsize=13, fontweight='bold', pad=10)
    ax.set_ylim(0, 115)
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    return save_fig(fig, 'newspaper_dates')


# ── Chart 3: Date match table (updated with new weighting) ────────────────

def chart_date_match_table():
    sources = ['Government', 'Science News', 'Sci. Research', 'Business Press',
               'Business', 'Futurists', 'Newspapers']
    current = ['N/A', '94.5%', '99.8%', '100.0%', '99.1%', 'N/A', '100.0%']

    fig, ax = plt.subplots(figsize=(5.5, 3))
    ax.axis('off')

    table_data = [['Source', 'Date Match %']]
    for s, c in zip(sources, current):
        table_data.append([s, c])

    table = ax.table(cellText=table_data, loc='center', cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.5)

    # Style header row
    for j in range(2):
        table[0, j].set_facecolor('#1D4ED8')
        table[0, j].set_text_props(color='white', fontweight='bold')

    # Color cells
    for i in range(1, len(table_data)):
        for j in range(2):
            table[i, j].set_facecolor('#F8FAFC')
        c = current[i - 1]
        if c == 'N/A':
            table[i, 1].set_facecolor('#DBEAFE')
        elif c == '100.0%':
            table[i, 1].set_facecolor('#DCFCE7')
        elif float(c.replace('%', '')) >= 94:
            table[i, 1].set_facecolor('#F0FDF4')

    ax.set_title('Date Match % by Source', fontsize=13, fontweight='bold', pad=15)
    return save_fig(fig, 'date_match_table')


# ── Chart 4: Article count comparison ─────────────────────────────────────

def chart_article_counts():
    sources =  ['Government', 'Sci. News', 'Sci. Research', 'Bus. Press',
                'Business', 'Futurists', 'Newspapers']
    reference = [926, 1407, 1102, 494, 4157, 926, 3762]
    output =    [905, 1197, 1079, 491, 4154, 926, 3354]

    fig, ax = plt.subplots(figsize=(8, 4))
    x = np.arange(len(sources))
    width = 0.35

    bars1 = ax.bar(x - width/2, reference, width, color='#94A3B8', label='Reference', edgecolor='white')
    bars2 = ax.bar(x + width/2, output, width, color='#1D4ED8', label='Output (parsed)', edgecolor='white')

    # Add coverage % labels
    for i, (r, o) in enumerate(zip(reference, output)):
        pct = o / r * 100
        ax.text(x[i] + width/2, o + 40, f'{pct:.0f}%', ha='center', va='bottom',
                fontsize=8, fontweight='bold', color='#1D4ED8')

    ax.set_xticks(x)
    ax.set_xticklabels(sources, fontsize=9, rotation=30, ha='right')
    ax.set_ylabel('Article Count', fontsize=11)
    ax.set_title('Article Coverage: Output vs Reference', fontsize=13, fontweight='bold', pad=10)
    ax.legend(fontsize=9, loc='upper right')
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    return save_fig(fig, 'article_counts')


# ── Chart 5: Article count detail table ───────────────────────────────────

def chart_article_count_table():
    sources =   ['Government', 'Science News', 'Sci. Research', 'Business Press',
                 'Business', 'Futurists', 'Newspapers', 'TOTAL']
    reference = [926, 1407, 1102, 494, 4157, 926, 3762, 12774]
    output =    [905, 1197, 1079, 491, 4154, 926, 3354, 12106]

    fig, ax = plt.subplots(figsize=(7, 3.5))
    ax.axis('off')

    table_data = [['Source', 'Reference', 'Output', 'Diff', 'Coverage']]
    for s, r, o in zip(sources, reference, output):
        diff = o - r
        pct = f'{o/r*100:.1f}%'
        diff_str = f'{diff:+d}'
        table_data.append([s, str(r), str(o), diff_str, pct])

    table = ax.table(cellText=table_data, loc='center', cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.4)

    # Style header row
    for j in range(5):
        table[0, j].set_facecolor('#1D4ED8')
        table[0, j].set_text_props(color='white', fontweight='bold')

    # Style data rows
    for i in range(1, len(table_data)):
        is_total = (i == len(table_data) - 1)
        for j in range(5):
            if is_total:
                table[i, j].set_facecolor('#E2E8F0')
                table[i, j].set_text_props(fontweight='bold')
            else:
                table[i, j].set_facecolor('#F8FAFC')
        # Color coverage green if >= 97%
        pct_val = output[i-1] / reference[i-1] * 100
        if pct_val >= 99:
            table[i, 4].set_facecolor('#DCFCE7')
        elif pct_val >= 95:
            table[i, 4].set_facecolor('#F0FDF4')

    ax.set_title('Article Count: Output vs Reference Worksheet', fontsize=13, fontweight='bold', pad=15)
    return save_fig(fig, 'article_count_table')


# ── Chart 6: Qualitative sampling comparison ───────────────────────────────

def chart_sampling():
    topics = ['Space', 'Electronics', 'Biology']
    before_counts = [42, 40, 40]
    after_counts = [88, 88, 88]

    x = np.arange(len(topics))
    width = 0.3

    fig, ax = plt.subplots(figsize=(5, 3))
    bars1 = ax.bar(x - width/2, before_counts, width, color='#D97706', label='Before (90-day window)')
    bars2 = ax.bar(x + width/2, after_counts, width, color='#16A34A', label='After (4/year quota)')

    for bars in [bars1, bars2]:
        for bar in bars:
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                    str(int(bar.get_height())), ha='center', va='bottom', fontsize=10, fontweight='bold')

    ax.set_xticks(x)
    ax.set_xticklabels(topics, fontsize=11)
    ax.set_ylabel('Articles Sampled', fontsize=11)
    ax.set_title('Qualitative Sampling Improvement', fontsize=13, fontweight='bold', pad=10)
    ax.legend(fontsize=8, loc='upper left')
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    ax.set_ylim(0, 110)
    return save_fig(fig, 'sampling')


# ── Build PDF ──────────────────────────────────────────────────────────────

class Report(FPDF):
    def header(self):
        if self.page_no() > 1:
            self.set_font('Helvetica', 'I', 8)
            self.set_text_color(120, 120, 120)
            self.cell(0, 8, 'Nanotech Research Pipeline - Update Report', align='R')
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

    def bold_text(self, label, text):
        self.set_font('Helvetica', 'B', 11)
        self.set_text_color(30, 30, 30)
        self.write(6, label)
        self.set_font('Helvetica', '', 11)
        self.write(6, text)
        self.ln(7)


def build_report():
    # Generate all charts
    img_science = chart_science_news()
    img_newspaper = chart_newspaper_dates()
    img_date_table = chart_date_match_table()
    img_article_counts = chart_article_counts()
    img_article_table = chart_article_count_table()
    img_sampling = chart_sampling()

    pdf = Report()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)

    # ── TITLE PAGE ──────────────────────────────────────────────────────────
    pdf.add_page()
    pdf.ln(40)
    pdf.set_font('Helvetica', 'B', 28)
    pdf.set_text_color(29, 78, 216)
    pdf.cell(0, 15, 'Pipeline Update Report', align='C')
    pdf.ln(14)
    pdf.set_font('Helvetica', '', 16)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 10, 'Nanotech Media Analysis Project', align='C')
    pdf.ln(20)
    pdf.set_font('Helvetica', '', 12)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 8, 'April 2026', align='C')
    pdf.ln(6)
    pdf.cell(0, 8, '7 improvements across scraper and analysis notebook', align='C')

    # ── OVERVIEW PAGE ───────────────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('Overview')
    pdf.body_text(
        'This report summarizes 7 improvements made to the nanotech media analysis '
        'pipeline. Changes span two files: the article scraper (create_excelV5.py) '
        'and the analysis notebook (nanotech_graphs.ipynb).'
    )
    pdf.ln(2)

    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(29, 78, 216)
    pdf.cell(0, 8, 'What was done:')
    pdf.ln(8)

    items = [
        ('Phase 1 - Data Fixes (Scraper)', [
            'Recovered 354 missing Science News articles',
            'Fixed newspaper date alignment (30% to 100% accuracy)',
            'Switched Year column to use reference worksheet values',
            'Updated date match weighting: year match = 90% minimum',
        ]),
        ('Phase 2 - Visualization Fixes (Notebook)', [
            'Removed misleading date match % for Government & Futurists',
            'Simplified co-occurrence graphs (top 8 keywords + "Other")',
        ]),
        ('Phase 3 - Sampling Improvements (Notebook)', [
            'Improved qualitative sampling with year-based quotas',
            'Added new AI keyword qualitative search sheet',
        ]),
    ]

    for phase_title, bullets in items:
        pdf.set_font('Helvetica', 'B', 11)
        pdf.set_text_color(60, 60, 60)
        pdf.cell(0, 7, phase_title)
        pdf.ln(7)
        for b in bullets:
            pdf.bullet(b)
        pdf.ln(3)

    # ── DATA QUALITY PAGE ─────────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('Data Quality: Article Coverage')

    pdf.body_text(
        'The chart below compares the number of articles in the parsed output '
        '(output.xlsx) against the reference worksheet (Combined community '
        'worksheet.xlsx). Overall coverage is 94.8% (12,106 of 12,774 articles). '
        'Futurists has perfect coverage. The main gaps are in Science News and '
        'Newspapers, where some articles in the reference could not be matched to '
        'entries in the raw text files.'
    )

    pdf.image(img_article_counts, x=15, w=175)
    pdf.ln(4)
    pdf.image(img_article_table, x=15, w=175)

    # ── PHASE 1 DETAIL ─────────────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('Phase 1: Data Fixes')

    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 8, '1. Science News - Recovering Missing Articles')
    pdf.ln(8)

    pdf.body_text(
        'The scraper was discarding articles where the article content was incorrectly '
        'identified as the title, leaving the body empty. This mostly affected short '
        '"This Week in Science" summaries. The fix detects when a title is unusually '
        'long (>20 words) with an empty body, and recovers the content.'
    )

    pdf.image(img_science, x=30, w=140)
    pdf.ln(8)

    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 8, '2. Newspaper Date Alignment')
    pdf.ln(8)

    pdf.body_text(
        'Newspaper articles were being matched to reference dates using a simple '
        'positional index (1st parsed article = 1st reference row, etc.). But the '
        'parser produced 3,354 articles while the reference had 3,762 rows, causing '
        'dates to drift out of alignment. The fix groups articles by date and matches '
        'within each date group, achieving 100% accuracy.'
    )

    pdf.image(img_newspaper, x=30, w=140)

    # ── YEAR + DATE ACCURACY ──────────────────────────────────────────────
    pdf.add_page()

    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 8, '3. Year Column Now Uses Reference Worksheet')
    pdf.ln(8)

    pdf.body_text(
        'Previously, the Year column was derived from the scraped date. Now it uses '
        'the Year value from the reference worksheet (Combined community worksheet.xlsx), '
        'which is more reliable since it was manually verified. Falls back to scraped '
        'date year when no reference is available.'
    )
    pdf.ln(2)

    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 8, '4. Date Match Weighting Updated')
    pdf.ln(8)

    pdf.body_text(
        'The accuracy scoring now gives heavier weight to year-level matches, since '
        'correct year is the most important factor for the temporal analysis. '
        'Scoring: exact date = 100%, same year+month = 95%, same year = 90%, '
        'wrong year = 0%. Government and Futurists are marked N/A since their '
        'source files have no inline dates.'
    )
    pdf.ln(2)

    pdf.image(img_date_table, x=35, w=130)
    pdf.ln(8)

    # ── PHASE 2 DETAIL ─────────────────────────────────────────────────────
    pdf.section_title('Phase 2: Visualization Fixes')

    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 8, '5. Gov/Futurist Date Match % Marked as N/A')
    pdf.ln(8)

    pdf.body_text(
        'The notebook was showing misleading date match percentages for Government '
        'and Futurists. These numbers were low simply because those source files lack '
        'inline dates, not because of scraper errors. They are now displayed as "N/A" '
        'while keeping the article counts visible.'
    )
    pdf.ln(2)

    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 8, '6. Co-occurrence Graphs Simplified')
    pdf.ln(8)

    pdf.body_text(
        'The co-occurrence graphs were plotting all 21 keywords per source, each with '
        'a different color. With so many overlapping lines, the graphs were unreadable. '
        'Now only the top 8 keywords (ranked by total co-occurrence) are shown with '
        'distinct colors. The remaining keywords are aggregated into a single gray '
        '"Other" line.'
    )

    # ── PHASE 3 DETAIL ─────────────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('Phase 3: Sampling Improvements')

    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 8, '7. Qualitative Sampling - Year-Based Quotas')
    pdf.ln(8)

    pdf.body_text(
        'The old sampling method picked one article every 90 days, which meant some '
        'years could end up with fewer than 4 samples if articles were clustered. '
        'The new method targets exactly 4 articles per year per topic. If a year has '
        'fewer than 4 qualifying articles, the shortfall carries forward to the next '
        'year, ensuring no data is lost.'
    )

    pdf.image(img_sampling, x=30, w=140)
    pdf.ln(8)

    pdf.set_font('Helvetica', 'B', 12)
    pdf.set_text_color(60, 60, 60)
    pdf.cell(0, 8, '8. New AI Qualitative Search Sheet')
    pdf.ln(8)

    pdf.body_text(
        'A new "AI Qualitative Samples" sheet has been added to output.xlsx. It works '
        'the same way as the existing qualitative sampling (Space, Electronics, Biology) '
        'but filters for articles mentioning both Nanotech and Artificial Intelligence. '
        'By default it samples from Futurists, but the source can be easily changed to '
        'any single source or all sources by adjusting a single variable in the notebook.'
    )

    # ── NEXT STEPS ─────────────────────────────────────────────────────────
    pdf.ln(10)
    pdf.section_title('Deferred / Future Work')

    pdf.bullet(
        'Claude API analysis: Use Claude to automatically code sampled articles on '
        'qualitative dimensions (attitude toward nanotech, analogy usage, etc.). '
        'Requires setting up an Anthropic API key.'
    )
    pdf.bullet(
        'Community dimension graphs: Once articles are coded on the qualitative '
        'dimensions, generate temporal graphs showing how different communities '
        '(Government, Business, Futurists, etc.) differ over time.'
    )

    # ── Save ───────────────────────────────────────────────────────────────
    output_path = OUT_DIR / 'Pipeline_Update_Report.pdf'
    pdf.output(str(output_path))
    print(f"Report saved to: {output_path}")

    # Clean up temp images
    for f in IMG_DIR.glob('*.png'):
        f.unlink()
    IMG_DIR.rmdir()


if __name__ == '__main__':
    build_report()
