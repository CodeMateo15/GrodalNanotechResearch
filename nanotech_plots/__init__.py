"""Plotting functions extracted from nanotech_graphs.ipynb.

Each plot function returns a matplotlib Figure (no plt.show), so the same
implementation can drive the interactive notebook and the PDF exporter
(export_figures.py).
"""

from .config import (
    YEAR_MIN, YEAR_MAX, SOURCE_INFO, KEYWORD_COLS,
    COOCCURRENCE_COLS, COLORS,
)
from .data import load_data, get_source_df, pct_per_year, pct_per_year_from_col, moving_avg
from . import plots
from . import registry

__all__ = [
    'YEAR_MIN', 'YEAR_MAX', 'SOURCE_INFO', 'KEYWORD_COLS',
    'COOCCURRENCE_COLS', 'COLORS',
    'load_data', 'get_source_df', 'pct_per_year', 'pct_per_year_from_col', 'moving_avg',
    'plots', 'registry',
]
