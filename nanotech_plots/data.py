"""Data loading + chart-axis helpers."""

import warnings

import numpy as np
import pandas as pd
import matplotlib.ticker as mticker

# The notebook calls warnings.filterwarnings('ignore') at import time; match
# that so pandas FutureWarnings from .fillna don't drown out export output.
warnings.filterwarnings('ignore', category=FutureWarning)

from .config import YEAR_MIN, YEAR_MAX, STANDARDIZE_Y


ALL_YEARS = list(range(YEAR_MIN, YEAR_MAX + 1))


def load_data(path='output.xlsx'):
    """Load Sheet1 and apply the same year-filter the notebook does."""
    df = pd.read_excel(path)
    df = df[df['Year'].notna()].copy()
    df['Year'] = df['Year'].astype(int)
    df = df[(df['Year'] >= YEAR_MIN) & (df['Year'] <= YEAR_MAX)]
    return df


def get_source_df(df_all, source_num):
    return df_all[df_all['Sources'] == source_num].copy()


def pct_per_year(df, mask_series):
    total = df.groupby('Year').size().reindex(ALL_YEARS, fill_value=0)
    hit = df[mask_series].groupby('Year').size().reindex(ALL_YEARS, fill_value=0)
    return (hit / total.replace(0, pd.NA) * 100).fillna(0)


def pct_per_year_from_col(df, col):
    return pct_per_year(df, df[col] >= 1)


def moving_avg(values, window=3):
    """3-year centered moving average; leading zeros become NaN so lines
    start at the first real data point instead of at the x-axis."""
    s = pd.Series(values, dtype=float)
    smoothed = s.rolling(window, center=True, min_periods=1).mean().values
    first_nonzero = None
    for i, v in enumerate(values):
        if v > 0:
            first_nonzero = i
            break
    if first_nonzero is not None and first_nonzero > 0:
        smoothed[:first_nonzero] = np.nan
    return smoothed


def smart_yticks(ax, values, is_percent=True):
    if is_percent and STANDARDIZE_Y:
        ax.set_ylim(0, 100)
        ax.set_yticks(range(0, 101, 10))
        ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
        return
    vmax = max(values) if len(values) > 0 and max(values) > 0 else 1
    if is_percent:
        ax.set_ylim(0, min(vmax * 1.2 + 0.5, 100))
        if vmax <= 5:
            step = 0.5
        elif vmax <= 15:
            step = 2
        elif vmax <= 40:
            step = 5
        else:
            step = 10
        ticks = np.arange(0, ax.get_ylim()[1] + step, step)
        ax.set_yticks(ticks)
        fmt = '%.1f%%' if step < 2 else '%.0f%%'
        ax.yaxis.set_major_formatter(mticker.FormatStrFormatter(fmt))
    else:
        ax.set_ylim(0, vmax * 1.15 + 1)
        ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f'))


def pct_per_year_categorical(df, col, value):
    """% of articles per year where col == value, among rows where col is non-null."""
    coded = df[df[col].notna()].copy()
    return pct_per_year(coded, coded[col] == value)


def pct_per_year_binary(df, col, threshold=1):
    """% of articles per year where col >= threshold, among rows where col is non-null."""
    coded = df[df[col].notna()].copy()
    return pct_per_year(coded, coded[col] >= threshold)


def make_bar(ax, values, color, title, is_percent=True):
    ax.bar(ALL_YEARS, values, color=color, edgecolor='white', linewidth=0.5)
    ax.set_title(title, fontsize=12, fontweight='bold', pad=8)
    ax.set_xlabel('Year', fontsize=10)
    ax.set_ylabel('% of articles' if is_percent else 'Count', fontsize=10)
    ax.set_xticks(ALL_YEARS)
    ax.set_xticklabels(ALL_YEARS, rotation=45, ha='right', fontsize=8)
    ax.spines[['top', 'right']].set_visible(False)
    ax.grid(axis='y', linestyle='--', alpha=0.4)
    smart_yticks(ax, values, is_percent=is_percent)
