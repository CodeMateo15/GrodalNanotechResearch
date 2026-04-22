"""Registry of every plot, tagged for full or summary PDF export.

ALL_PLOTS is a list of (name, callable(df_all) -> Figure, is_summary) tuples.
The callable closes over its arguments (source, keyword, etc.) so callers
don't need to know plot-specific parameters.
"""

from functools import partial

from . import plots
from .config import SOURCE_INFO, KEYWORD_COLS, COOCCURRENCE_COLS


def _all_sources():
    return sorted(SOURCE_INFO.keys())


def build_registry():
    """Build the full plot list. Returns [(name, fn_of_df, is_summary), ...]."""
    srcs = _all_sources()
    items = []

    # Cell 20 — macro (summary)
    items.append((
        'Macro – any keyword co-occurrence, all sources',
        partial(plots.make_any_keyword_macro_fig, sources=srcs),
        True,
    ))

    # Cell 8 — % of articles mentioning nanotech, per source (summary)
    for src in srcs:
        name = SOURCE_INFO[src][0]
        items.append((
            f'{name} — % of articles mentioning nanotech',
            partial(plots.make_nanotech_pct_fig, src=src),
            True,
        ))

    # Cell 18 — Top 8 co-occurrence per source (summary for a few key sources)
    summary_sources = {1, 5, 6, 7}  # Government, Business, Futurists, Newspapers
    for src in srcs:
        name = SOURCE_INFO[src][0]
        items.append((
            f'{name} — Top 8 co-occurrence (3yr MA)',
            partial(plots.make_top8_per_source_fig, src=src),
            src in summary_sources,
        ))

    # Cell 4 — diagnostics per source (full only)
    for src in srcs:
        name = SOURCE_INFO[src][0]
        items.append((
            f'{name} — Diagnostics (word count + article count)',
            partial(plots.make_diagnostics_fig, src=src),
            False,
        ))

    # Cell 6 — total nanotech mentions per source (full only)
    for src in srcs:
        name = SOURCE_INFO[src][0]
        items.append((
            f'{name} — Total nanotech mentions per year',
            partial(plots.make_nanotech_total_fig, src=src),
            False,
        ))

    # Cell 10 — % articles with zero nanotech mentions (full only)
    for src in srcs:
        name = SOURCE_INFO[src][0]
        items.append((
            f'{name} — % articles with zero nanotech mentions',
            partial(plots.make_zero_nanotech_fig, src=src),
            False,
        ))

    # Cell 14 — keyword-overview grid per source (full only)
    for src in srcs:
        name = SOURCE_INFO[src][0]
        items.append((
            f'{name} — Keyword overview (all keywords)',
            partial(plots.make_keyword_overview_fig, src=src),
            False,
        ))

    # Cell 12 — % of articles mentioning each keyword, per source (full only, dense)
    for src in srcs:
        src_name = SOURCE_INFO[src][0]
        for i, col in enumerate(KEYWORD_COLS):
            items.append((
                f'{src_name} — % of articles mentioning "{col}"',
                partial(plots.make_keyword_single_fig, src=src, col=col, color_idx=i),
                False,
            ))

    # Cell 16 — per-keyword co-occurrence overlay (full only)
    for col in COOCCURRENCE_COLS:
        items.append((
            f'Co-occurrence — % Nanotech articles also mentioning "{col}" (all sources)',
            partial(plots.make_cooccurrence_per_keyword_fig, sources=srcs, col=col),
            False,
        ))

    return items


ALL_PLOTS = build_registry()
