"""Year-quota sampler with deficit carry-forward."""

import numpy as np
import pandas as pd


def sample_with_quota(articles_df, target_per_year, year_min, year_max, date_col='Date'):
    """Sample articles using a year-based quota with deficit carry-forward.

    For each year in [year_min, year_max]:
      - target `target_per_year` articles plus any deficit carried from earlier years
      - if fewer qualifying articles exist, take them all and carry the shortfall forward
      - sampling is evenly spaced across the sorted date column (np.linspace of integer indices)

    Returns a single concatenated DataFrame preserving all input columns.
    """
    samples = []
    deficit = 0
    for year in range(year_min, year_max + 1):
        year_articles = articles_df[articles_df['Year'] == year].sort_values(date_col)
        quota = target_per_year + deficit
        available = len(year_articles)
        if available == 0:
            deficit = quota
            continue
        n_sample = min(quota, available)
        if n_sample >= available:
            selected = year_articles
        else:
            indices = np.linspace(0, available - 1, n_sample, dtype=int)
            selected = year_articles.iloc[indices]
        samples.append(selected)
        deficit = quota - n_sample
    if samples:
        return pd.concat(samples)
    return pd.DataFrame()
