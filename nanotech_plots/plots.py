"""Plot functions extracted from nanotech_graphs.ipynb cells 4, 6, 8, 10, 12, 14, 16, 18, 20.

Every function returns a matplotlib Figure; callers decide whether to show,
save, or close.
"""

import numpy as np
import matplotlib.pyplot as plt

from .config import SOURCE_INFO, KEYWORD_COLS, COOCCURRENCE_COLS, COLORS
from .data import (
    ALL_YEARS, get_source_df, pct_per_year, pct_per_year_from_col,
    moving_avg, smart_yticks, make_bar,
)


# Cell 4 ─────────────────────────────────────────────────────────────────
def make_diagnostics_fig(df_all, src):
    df = get_source_df(df_all, src)
    name, color = SOURCE_INFO.get(src, (str(src), '#888888'))
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 7), sharex=True)
    fig.suptitle(f'{name} — Diagnostics', fontsize=14, fontweight='bold')

    word_avg = df.groupby('Year')['Word count'].mean().reindex(ALL_YEARS, fill_value=0)
    make_bar(ax1, word_avg.values, color,
             'Average word count per article per year', is_percent=False)
    ax1.set_ylabel('Avg word count')

    art_count = df.groupby('Year').size().reindex(ALL_YEARS, fill_value=0)
    make_bar(ax2, art_count.values, color,
             'Number of articles per year', is_percent=False)
    ax2.set_ylabel('# of articles')
    plt.tight_layout()
    return fig


# Cell 6 ─────────────────────────────────────────────────────────────────
def make_nanotech_total_fig(df_all, src):
    df = get_source_df(df_all, src)
    name, color = SOURCE_INFO.get(src, (str(src), '#888888'))
    nano_sum = df.groupby('Year')['Nanotech'].sum().reindex(ALL_YEARS, fill_value=0)
    fig, ax = plt.subplots(figsize=(14, 4))
    make_bar(ax, nano_sum.values, color,
             f'{name} — Total Nanotech mentions per year', is_percent=False)
    ax.set_ylabel('Total mentions')
    plt.tight_layout()
    return fig


# Cell 8 ─────────────────────────────────────────────────────────────────
def make_nanotech_pct_fig(df_all, src):
    df = get_source_df(df_all, src)
    name, color = SOURCE_INFO.get(src, (str(src), '#888888'))
    pct = pct_per_year(df, df['Nanotech'] >= 1)
    fig, ax = plt.subplots(figsize=(14, 4))
    make_bar(ax, pct.values, color,
             f'{name} — % of articles mentioning nanotech >= 1 time')
    plt.tight_layout()
    return fig


# Cell 10 ────────────────────────────────────────────────────────────────
def make_zero_nanotech_fig(df_all, src):
    df = get_source_df(df_all, src)
    name, _ = SOURCE_INFO.get(src, (str(src), '#888888'))
    pct = pct_per_year(df, df['Nanotech'] == 0)
    fig, ax = plt.subplots(figsize=(14, 4))
    make_bar(ax, pct.values, '#DC2626',
             f'{name} — % of articles with zero nanotech mentions')
    plt.tight_layout()
    return fig


# Cell 12 ────────────────────────────────────────────────────────────────
def make_keyword_single_fig(df_all, src, col, color_idx=0):
    df = get_source_df(df_all, src)
    name, _ = SOURCE_INFO.get(src, (str(src), '#888888'))
    if col not in df.columns:
        return None
    pct = pct_per_year_from_col(df, col)
    fig, ax = plt.subplots(figsize=(14, 4))
    make_bar(ax, pct.values, COLORS[color_idx % len(COLORS)],
             f'{name} — % of articles mentioning "{col}"')
    plt.tight_layout()
    return fig


# Cell 14 ────────────────────────────────────────────────────────────────
def make_keyword_overview_fig(df_all, src):
    df = get_source_df(df_all, src)
    name, _ = SOURCE_INFO.get(src, (str(src), '#888888'))
    valid_cols = [c for c in KEYWORD_COLS if c in df.columns]
    ncols = 3
    nrows = -(-len(valid_cols) // ncols)

    fig, axes = plt.subplots(nrows, ncols, figsize=(18, nrows * 3.5))
    axes = axes.flatten()

    for i, col in enumerate(valid_cols):
        pct = pct_per_year_from_col(df, col)
        make_bar(axes[i], pct.values, COLORS[i % len(COLORS)], col)
        axes[i].set_xlabel('')
        axes[i].tick_params(axis='x', labelsize=6)

    for j in range(len(valid_cols), len(axes)):
        axes[j].set_visible(False)

    fig.suptitle(f'{name} — % of articles per keyword per year',
                 fontsize=15, fontweight='bold', y=1.01)
    plt.tight_layout()
    return fig


# Cell 16 ────────────────────────────────────────────────────────────────
def make_cooccurrence_per_keyword_fig(df_all, sources, col):
    """All sources overlaid on one plot, 3yr MA of `col` co-occurrence with Nanotech."""
    if not any(col in get_source_df(df_all, s).columns for s in sources):
        return None
    fig, ax = plt.subplots(figsize=(14, 4))
    all_vals = []
    for src in sources:
        df = get_source_df(df_all, src)
        if col not in df.columns:
            continue
        nano_df = df[df['Nanotech'] >= 1]
        name, color = SOURCE_INFO.get(src, (str(src), '#888888'))
        raw = pct_per_year_from_col(nano_df, col).values
        smoothed = moving_avg(raw)
        all_vals.extend([v for v in smoothed if not np.isnan(v)])
        ax.plot(ALL_YEARS, smoothed, marker='o', markersize=3,
                linewidth=1.8, color=color, label=name)
    ax.set_title(
        f'Among Nanotech >= 1 — % also mentioning "{col}" (3yr MA) — All Sources',
        fontsize=12, fontweight='bold', pad=8,
    )
    ax.set_xlabel('Year', fontsize=10)
    ax.set_ylabel('% of articles', fontsize=10)
    ax.set_xticks(ALL_YEARS)
    ax.set_xticklabels(ALL_YEARS, rotation=45, ha='right', fontsize=8)
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.4)
    smart_yticks(ax, all_vals, is_percent=True)
    ax.legend(fontsize=8, loc='upper left', framealpha=0.7)
    plt.tight_layout()
    return fig


# Cell 18 ────────────────────────────────────────────────────────────────
def make_top8_per_source_fig(df_all, src, top_n=8):
    df = get_source_df(df_all, src)
    nano_df = df[df['Nanotech'] >= 1]
    name, _ = SOURCE_INFO.get(src, (str(src), '#888888'))

    valid_cols = [c for c in COOCCURRENCE_COLS if c in nano_df.columns]
    totals = {col: nano_df[col].clip(lower=0).sum() for col in valid_cols}
    ranked = sorted(totals, key=totals.get, reverse=True)
    top_cols = ranked[:top_n]
    other_cols = ranked[top_n:]

    fig, ax = plt.subplots(figsize=(14, 5))
    all_vals = []

    for i, col in enumerate(top_cols):
        raw = pct_per_year_from_col(nano_df, col).values
        smoothed = moving_avg(raw)
        all_vals.extend([v for v in smoothed if not np.isnan(v)])
        ax.plot(ALL_YEARS, smoothed, linewidth=1.6, marker='o', markersize=2,
                color=COLORS[i % len(COLORS)], label=col)

    if other_cols:
        other_pcts = []
        for yr in ALL_YEARS:
            yr_nano = nano_df[nano_df['Year'] == yr]
            if len(yr_nano) == 0:
                other_pcts.append(0)
            else:
                has_any = (yr_nano[other_cols].sum(axis=1) >= 1).sum()
                other_pcts.append(100 * has_any / len(yr_nano))
        smoothed = moving_avg(other_pcts)
        all_vals.extend([v for v in smoothed if not np.isnan(v)])
        ax.plot(ALL_YEARS, smoothed, linewidth=1.4, linestyle='--',
                color='#999999', label=f'Other ({len(other_cols)} keywords)')

    ax.set_title(
        f'{name} — Co-occurrence with Nanotech >= 1, top {top_n} keywords (3yr MA)',
        fontsize=12, fontweight='bold', pad=8,
    )
    ax.set_xlabel('Year', fontsize=10)
    ax.set_ylabel('% of articles', fontsize=10)
    ax.set_xticks(ALL_YEARS)
    ax.set_xticklabels(ALL_YEARS, rotation=45, ha='right', fontsize=8)
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.4)
    smart_yticks(ax, all_vals, is_percent=True)
    ax.legend(fontsize=7, loc='upper left', framealpha=0.7,
              ncol=1, bbox_to_anchor=(1.01, 1), borderaxespad=0)
    plt.tight_layout()
    return fig


# Cell 20 ────────────────────────────────────────────────────────────────
def make_any_keyword_macro_fig(df_all, sources):
    fig, ax = plt.subplots(figsize=(14, 5))
    all_vals = []
    for src in sources:
        df = get_source_df(df_all, src)
        nano_df = df[df['Nanotech'] >= 1].copy()
        name, color = SOURCE_INFO.get(src, (str(src), '#888888'))
        valid = [c for c in COOCCURRENCE_COLS if c in nano_df.columns]
        nano_df['any_keyword'] = (nano_df[valid] >= 1).any(axis=1).astype(int)
        raw = pct_per_year(nano_df, nano_df['any_keyword'] == 1).values
        smoothed = moving_avg(raw)
        all_vals.extend([v for v in smoothed if not np.isnan(v)])
        ax.plot(ALL_YEARS, smoothed, linewidth=2, marker='o', markersize=3,
                color=color, label=name)
    ax.set_title(
        'Macro — % of Nanotech >= 1 articles mentioning any keyword (3yr MA) — All Sources',
        fontsize=12, fontweight='bold', pad=8,
    )
    ax.set_xlabel('Year', fontsize=10)
    ax.set_ylabel('% of articles', fontsize=10)
    ax.set_xticks(ALL_YEARS)
    ax.set_xticklabels(ALL_YEARS, rotation=45, ha='right', fontsize=8)
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.4)
    smart_yticks(ax, all_vals, is_percent=True)
    ax.legend(fontsize=9, loc='upper left', framealpha=0.7)
    plt.tight_layout()
    return fig
