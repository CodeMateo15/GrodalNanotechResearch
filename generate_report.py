"""Generate a PDF report for Round 3 pipeline work (April 30, 2026).

Covers:
  1. Imported Claude analysis from old integration file (10 columns, ~4,300 rows)
  2. Normalized qualitative coding columns for graphing
  3. Added 13 qualitative dimension graphs over time
  4. Combined all-sources overlay charts for summary PDF
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
from fpdf import FPDF
from pathlib import Path

from nanotech_plots.data import load_data
from nanotech_plots.registry import ALL_PLOTS

OUT_DIR = Path(__file__).parent
IMG_DIR = OUT_DIR / "_report_imgs"
IMG_DIR.mkdir(exist_ok=True)


def save_fig(fig, name):
    path = IMG_DIR / f"{name}.png"
    fig.savefig(path, dpi=180, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return str(path)


# ── Chart helpers ────────────────────────────────────────────────────────

def chart_claude_import_coverage(df):
    """Per-source bar chart showing how many rows have Claude analysis."""
    source_names = {1: 'Government', 2: 'Sci. News', 3: 'Sci. Research',
                    4: 'Bus. Press', 5: 'Business', 6: 'Futurists', 7: 'Newspapers'}
    srcs = sorted(source_names.keys())
    labels = [source_names[s] for s in srcs]
    totals = [len(df[df['Sources'] == s]) for s in srcs]
    coded = [len(df[(df['Sources'] == s) & (df['Attitude'].notna())]) for s in srcs]
    pcts = [100 * c / t if t > 0 else 0 for c, t in zip(coded, totals)]

    x = np.arange(len(labels))
    width = 0.36

    fig, ax = plt.subplots(figsize=(10, 4))
    ax.bar(x - width/2, totals, width, color='#94A3B8', label='Total articles', edgecolor='white')
    ax.bar(x + width/2, coded,  width, color='#7C3AED', label='With Claude analysis', edgecolor='white')

    for i in range(len(labels)):
        ax.text(x[i] + width/2, coded[i] + 30, f'{pcts[i]:.0f}%',
                ha='center', va='bottom', fontsize=8, fontweight='bold', color='#7C3AED')

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=9, rotation=20, ha='right')
    ax.set_ylabel('Article count', fontsize=10)
    ax.set_title('Claude Analysis Coverage by Source (imported from old file)',
                 fontsize=13, fontweight='bold', pad=10)
    ax.legend(fontsize=9, loc='upper right')
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    return save_fig(fig, 'claude_import_coverage')


def chart_attitude_distribution(df):
    """Stacked bar: attitude breakdown per source."""
    source_names = {1: 'Government', 2: 'Sci. News', 4: 'Bus. Press',
                    5: 'Business', 6: 'Futurists'}
    srcs = [s for s in sorted(source_names.keys())
            if len(df[(df['Sources'] == s) & (df['Attitude'].notna())]) > 0]
    labels = [source_names[s] for s in srcs]

    attitudes = ['positive', 'neutral', 'negative', 'mixed']
    colors = ['#16A34A', '#94A3B8', '#DC2626', '#F59E0B']

    x = np.arange(len(labels))
    width = 0.6

    fig, ax = plt.subplots(figsize=(9, 4))
    bottom = np.zeros(len(labels))

    for att, color in zip(attitudes, colors):
        vals = []
        for s in srcs:
            coded = df[(df['Sources'] == s) & (df['Attitude'].notna())]
            total = len(coded)
            count = len(coded[coded['Attitude'] == att])
            vals.append(100 * count / total if total > 0 else 0)
        ax.bar(x, vals, width, bottom=bottom, color=color, label=att.capitalize(), edgecolor='white', linewidth=0.5)
        bottom += vals

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10)
    ax.set_ylabel('% of coded articles', fontsize=10)
    ax.set_title('Attitude Toward Nanotechnology by Source', fontsize=13, fontweight='bold', pad=10)
    ax.legend(fontsize=9, loc='upper right')
    ax.set_ylim(0, 105)
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    return save_fig(fig, 'attitude_distribution')


def chart_dimension_summary(df):
    """Grouped bar: % analogy present, % disrupting, % funding argument per source."""
    source_names = {1: 'Government', 2: 'Sci. News', 4: 'Bus. Press',
                    5: 'Business', 6: 'Futurists'}
    srcs = [s for s in sorted(source_names.keys())
            if len(df[(df['Sources'] == s) & (df['Attitude'].notna())]) > 0]
    labels = [source_names[s] for s in srcs]

    analogy_pcts, disrupt_pcts, funding_pcts = [], [], []
    for s in srcs:
        coded = df[(df['Sources'] == s) & (df['Attitude'].notna())]
        total = len(coded)
        if total == 0:
            analogy_pcts.append(0); disrupt_pcts.append(0); funding_pcts.append(0)
            continue
        analogy_pcts.append(100 * (coded['Analogy Present'] == 1).sum() / total)

        sd = coded[coded['Sustaining vs Disrupting'].notna()]
        disrupt_pcts.append(100 * (sd['Sustaining vs Disrupting'] == 'disrupting').sum() / len(sd) if len(sd) > 0 else 0)

        fa = coded[coded['Funding Argument Present'].notna()]
        funding_pcts.append(100 * (fa['Funding Argument Present'] == 1).sum() / len(fa) if len(fa) > 0 else 0)

    x = np.arange(len(labels))
    width = 0.25

    fig, ax = plt.subplots(figsize=(10, 4))
    ax.bar(x - width, analogy_pcts, width, color='#0891B2', label='Analogy present', edgecolor='white')
    ax.bar(x,         disrupt_pcts, width, color='#DC2626', label='Disrupting status quo', edgecolor='white')
    ax.bar(x + width, funding_pcts, width, color='#16A34A', label='Funding argument', edgecolor='white')

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10, rotation=20, ha='right')
    ax.set_ylabel('% of coded articles', fontsize=10)
    ax.set_title('Qualitative Dimensions by Source', fontsize=13, fontweight='bold', pad=10)
    ax.legend(fontsize=9, loc='upper right')
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    return save_fig(fig, 'dimension_summary')


def save_registry_plot(df, plot_name):
    """Find a plot by name in the registry, generate it, and save as image."""
    for name, fn, _ in ALL_PLOTS:
        if plot_name in name:
            fig = fn(df)
            if fig is not None:
                safe = plot_name[:40].replace(' ', '_').replace('"', '').replace('%', 'pct')
                return save_fig(fig, safe)
    return None


# ── PDF class ────────────────────────────────────────────────────────────

class Report(FPDF):
    def header(self):
        if self.page_no() > 1:
            self.set_font('Helvetica', 'I', 8)
            self.set_text_color(120, 120, 120)
            self.cell(0, 8, 'Nanotech Research Pipeline - Round 3 Report', align='R')
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
    print("Loading output.xlsx ...")
    df = load_data()
    print(f"  ({len(df)} rows, Attitude non-null: {df['Attitude'].notna().sum()})")

    print("Generating charts ...")
    img_claude_coverage = chart_claude_import_coverage(df)
    img_attitude_dist   = chart_attitude_distribution(df)
    img_dimension_sum   = chart_dimension_summary(df)

    # Time-series plots from the registry
    img_positive     = save_registry_plot(df, '% "positive" coded')
    img_negative     = save_registry_plot(df, '% "negative" coded')
    img_analogy      = save_registry_plot(df, 'Analogy Presence')
    img_disrupting   = save_registry_plot(df, '% "disrupting"')
    img_sustaining   = save_registry_plot(df, '% "sustaining"')
    img_funding      = save_registry_plot(df, 'Funding Argument')
    img_temp_future  = save_registry_plot(df, 'Temporality — % "future"')
    img_temp_present = save_registry_plot(df, 'Temporality — % "present"')
    img_nanotech_all = save_registry_plot(df, '% of articles mentioning nanotech (3yr MA)')
    img_article_count = save_registry_plot(df, 'Article count per year')

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
    pdf.cell(0, 10, 'Round 3', align='C')
    pdf.ln(12)
    pdf.set_font('Helvetica', '', 16)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(0, 10, 'Nanotech Media Analysis Project', align='C')
    pdf.ln(20)
    pdf.set_font('Helvetica', '', 12)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 8, 'April 30, 2026', align='C')
    pdf.ln(6)
    pdf.cell(0, 8, '12,206 articles | 7 sources | 1983-2005', align='C')

    # ── OVERVIEW PAGE ──────────────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('Overview')
    pdf.body_text(
        'This round imported existing Claude analysis from a prior integration '
        'file, normalized the qualitative coding columns for graphing, and built '
        '13 new time-series charts showing how sources/communities differ across '
        'five qualitative dimensions: attitude toward nanotechnology, analogy '
        'presence, analogy temporality, sustaining vs disrupting framing, and '
        'funding arguments.'
    )
    pdf.ln(2)

    pdf.subheading('What was done:')
    items = [
        '1. Imported Claude analysis from old integration file: 10 columns, ~4,341 rows matched by Title + Sources key.',
        '2. Normalized 5 graphing columns: Attitude, Analogy Present, Analogy Temporality (Coded), Sustaining vs Disrupting, Funding Argument Present.',
        '3. Added 13 qualitative dimension time-series charts (3yr moving average, all sources overlaid).',
        '4. Replaced per-source summary charts with combined all-sources overlay charts. Summary PDF now has 21 figures.',
    ]
    for item in items:
        pdf.bullet(item)

    # ── 1. CLAUDE IMPORT ──────────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('1. Imported Claude analysis from old file')
    pdf.body_text(
        'The file "Old Python Files/Claude integration of output file.xlsx" '
        'contained prior Claude analysis for ~5,000 articles across its "Big dataset" '
        'sheet (10,156 rows, 42 columns). import_claude_columns.py matched rows by '
        'Title + Sources key and imported 10 Claude analysis columns into the current '
        'output.xlsx.'
    )
    pdf.subheading('Imported columns:')
    pdf.bullet('Sentiment Toward Nanotechnology (4,341 rows matched)')
    pdf.bullet('Analogies, Purpose of Analogies (997 rows with text)')
    pdf.bullet('Temporality of Analogies (162 unique free-text values)')
    pdf.bullet('Analogy: Status Quo (Reinforcing/Disrupting/Both)')
    pdf.bullet('Analogy: Funding Argument (Yes/No)')
    pdf.bullet('Source Analyzed, Perspective on Technology, Mentions Nanotechnology, Actors Mentioned')

    pdf.ln(2)
    pdf.image(img_claude_coverage, x=10, w=190)

    pdf.ln(4)
    pdf.body_text(
        'Science Research and Newspapers have 0% coverage because they were not '
        'present in the old integration file. Science News has low coverage (18%) '
        'because most of its articles in the old file did not receive Claude analysis.'
    )

    # ── 2. NORMALIZATION ──────────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('2. Normalized qualitative coding columns')
    pdf.body_text(
        'The old file used different column names and value formats than the current '
        'coding schema. import_claude_columns.py normalizes these into 5 graph-ready '
        'columns:'
    )
    pdf.ln(1)

    norm_rows = [
        ['New Column', 'Source', 'Mapping'],
        ['Attitude', 'Sentiment Toward Nanotech', 'lowercase (Positive -> positive)'],
        ['Analogy Present', 'Analogies', '1 if text exists, 0 if null but coded'],
        ['Analogy Temporality (Coded)', 'Temporality of Analogies', 'Free text -> past/present/future/mixed'],
        ['Sustaining vs Disrupting', 'Analogy: Status Quo', 'Reinforcing -> sustaining'],
        ['Funding Argument Present', 'Analogy: Funding Argument', 'Yes -> 1, No -> 0'],
    ]

    fig, ax = plt.subplots(figsize=(9, 3))
    ax.axis('off')
    table = ax.table(cellText=norm_rows, loc='center', cellLoc='left')
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1, 1.4)
    for j in range(len(norm_rows[0])):
        table[0, j].set_facecolor('#1D4ED8')
        table[0, j].set_text_props(color='white', fontweight='bold')
    for i in range(1, len(norm_rows)):
        for j in range(len(norm_rows[0])):
            table[i, j].set_facecolor('#F8FAFC')
    img_norm = save_fig(fig, 'normalization_table')
    pdf.image(img_norm, x=10, w=190)

    pdf.ln(4)
    pdf.subheading('Value distributions after normalization:')
    pdf.mono(
        f"Attitude:      positive={len(df[df['Attitude']=='positive']):,}  "
        f"neutral={len(df[df['Attitude']=='neutral']):,}  "
        f"negative={len(df[df['Attitude']=='negative']):,}  "
        f"mixed={len(df[df['Attitude']=='mixed']):,}\n"
        f"Analogy:       present={int((df['Analogy Present']==1).sum()):,}  "
        f"absent={int((df['Analogy Present']==0).sum()):,}\n"
        f"Status Quo:    sustaining={len(df[df['Sustaining vs Disrupting']=='sustaining']):,}  "
        f"disrupting={len(df[df['Sustaining vs Disrupting']=='disrupting']):,}  "
        f"both={len(df[df['Sustaining vs Disrupting']=='both']):,}\n"
        f"Funding:       yes={int((df['Funding Argument Present']==1).sum()):,}  "
        f"no={int((df['Funding Argument Present']==0).sum()):,}"
    )

    # ── 3. QUALITATIVE GRAPHS ────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('3. Qualitative dimension graphs over time')
    pdf.body_text(
        '13 new time-series charts showing how sources/communities differ across '
        '5 qualitative dimensions. Each chart overlays all sources with available '
        'data, using 3-year centered moving averages. Only sources with coded '
        'articles appear (Science Research and Newspapers are absent).'
    )

    pdf.image(img_attitude_dist, x=10, w=190)
    pdf.ln(2)
    pdf.image(img_dimension_sum, x=10, w=190)

    # Attitude over time
    pdf.add_page()
    pdf.subheading('Attitude over time')
    pdf.body_text(
        'Percentage of coded articles with positive vs negative attitude toward '
        'nanotechnology, per source, over time.'
    )
    if img_positive:
        pdf.image(img_positive, x=10, w=190)
        pdf.ln(2)
    if img_negative:
        pdf.image(img_negative, x=10, w=190)

    # Analogy
    pdf.add_page()
    pdf.subheading('Analogy presence over time')
    pdf.body_text(
        'Percentage of coded articles that use an analogy or metaphor to '
        'explain/frame nanotechnology.'
    )
    if img_analogy:
        pdf.image(img_analogy, x=10, w=190)

    pdf.ln(4)
    pdf.subheading('Analogy temporality over time')
    pdf.body_text(
        'Among articles with analogies, the temporal orientation of the comparison.'
    )
    if img_temp_future:
        pdf.image(img_temp_future, x=10, w=190)
        pdf.ln(2)
    if img_temp_present:
        pdf.image(img_temp_present, x=10, w=190)

    # Sustaining/Disrupting
    pdf.add_page()
    pdf.subheading('Sustaining vs Disrupting over time')
    pdf.body_text(
        'Whether nanotechnology is framed as sustaining existing industries '
        'or disrupting the status quo.'
    )
    if img_sustaining:
        pdf.image(img_sustaining, x=10, w=190)
        pdf.ln(2)
    if img_disrupting:
        pdf.image(img_disrupting, x=10, w=190)

    # Funding
    pdf.add_page()
    pdf.subheading('Funding argument over time')
    pdf.body_text(
        'Percentage of coded articles that contain an argument for more '
        'nanotechnology funding.'
    )
    if img_funding:
        pdf.image(img_funding, x=10, w=190)

    # ── 4. COMBINED OVERLAY CHARTS ───────────────────────────────────────
    pdf.add_page()
    pdf.section_title('4. Combined all-sources overlay charts')
    pdf.body_text(
        'The summary PDF now uses combined all-sources overlay charts instead of '
        'per-source individual charts. All 7 sources appear on one plot with 3-year '
        'moving averages, making cross-source comparison much easier.'
    )
    pdf.body_text(
        'Summary PDF contents (21 figures): 3 combined keyword charts (article '
        'count, nanotech %, nanotech total), 1 macro co-occurrence, 4 top-8 '
        'co-occurrence (key sources), and 13 qualitative dimension charts.'
    )
    if img_nanotech_all:
        pdf.image(img_nanotech_all, x=10, w=190)
        pdf.ln(2)
    if img_article_count:
        pdf.image(img_article_count, x=10, w=190)

    # ── NEXT STEPS ───────────────────────────────────────────────────────
    pdf.add_page()
    pdf.section_title('Next steps')

    pdf.subheading('Data gaps:')
    pdf.bullet(
        'Science Research and Newspapers have 0% Claude analysis coverage. '
        'Run python -m qualitative_coding --sheets "Sheet1" to code all uncoded '
        'articles via the API (~$30-60, ~7,800 articles).'
    )
    pdf.bullet(
        'Science News has only 18% coverage from the old file import.'
    )

    pdf.subheading('Quality:')
    pdf.bullet(
        'Calibrate qualitative coding: hand-code 10 articles and compare against '
        "Claude's output. Target >=80% agreement on Attitude and Sustaining/Disrupting."
    )

    pdf.subheading('Analysis:')
    pdf.bullet(
        'The qualitative dimension graphs currently cover ~4,300 of 12,200 articles. '
        'Running the Claude API on the remaining articles would give full coverage '
        'and enable meaningful graphs for all 7 sources.'
    )
    pdf.bullet(
        'Stars vs reference count validation per next_steps.txt points 1 and 2 '
        '(being done manually).'
    )

    # ── SAVE ─────────────────────────────────────────────────────────────
    output_path = OUT_DIR / 'Pipeline_Update_Report.pdf'
    pdf.output(str(output_path))
    print(f"Report saved to: {output_path}")

    # Clean up temp images
    for f in IMG_DIR.glob('*.png'):
        f.unlink()
    IMG_DIR.rmdir()


if __name__ == '__main__':
    build_report()
