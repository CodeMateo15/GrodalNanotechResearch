"""Sample articles that mention both nanotech and AI."""

import pandas as pd

from .quota import sample_with_quota

AI_ARTICLES_PER_YEAR = 4
AI_COL = 'Artificial Intelligence'
BODY_CHAR_LIMIT = 2000


def build_ai_qualitative_samples(df_all, source_filter=6, year_min=1983, year_max=2005):
    """Build the 'AI Qualitative Samples' dataframe.

    source_filter: Sources numeric code to restrict to (default 6 = Futurists),
                   or None to sample across all sources.
    """
    if source_filter is not None:
        ai_base = df_all[df_all['Sources'] == source_filter].copy()
    else:
        ai_base = df_all.copy()

    # Falls back to Scraped Date where Original Date is missing — some Futurist
    # rows only have a scraped date.
    ai_base['_Date'] = pd.to_datetime(
        ai_base['Original Date'].fillna(ai_base['Scraped Date']),
        format='%d %B %Y', errors='coerce'
    )
    ai_base = ai_base.dropna(subset=['_Date'])

    if AI_COL not in ai_base.columns:
        print(f"Column '{AI_COL}' not found in dataset, skipping AI qualitative samples.")
        return pd.DataFrame()

    ai_matched = ai_base[
        (ai_base['Nanotech'] >= 1) & (ai_base[AI_COL] >= 1)
    ].copy().sort_values('_Date')

    if len(ai_matched) == 0:
        return pd.DataFrame()

    sampled = sample_with_quota(
        ai_matched, AI_ARTICLES_PER_YEAR, year_min, year_max, date_col='_Date'
    )

    rows = []
    for _, row in sampled.iterrows():
        rows.append({
            'Year': row['Year'],
            'Month': row['_Date'].strftime('%B %Y'),
            'Date': row['_Date'].strftime('%d %B %Y'),
            'Source': row['Name'],
            'Title': row['Title'],
            'Body': str(row['Body'])[:BODY_CHAR_LIMIT],
            'Word Count': row['Word count'],
            'Nanotech Mentions': row['Nanotech'],
            'AI Mentions': row[AI_COL],
        })
    return pd.DataFrame(rows)
