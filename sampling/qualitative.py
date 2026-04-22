"""Sample Futurist articles that mention both nanotech and one of Space/Electronics/Biology."""

import pandas as pd

from .quota import sample_with_quota

QUAL_TOPICS = {
    'Space': 'Space',
    'Electronics': 'Electronics',
    'Biology': 'Biotech/Biology',
}

ARTICLES_PER_YEAR = 4
BODY_CHAR_LIMIT = 2000


def build_qualitative_samples(df_all, year_min=1983, year_max=2005):
    """Build the 'Qualitative Samples' dataframe from the full article dataset.

    Filters Futurist rows (Sources == 6) that have Nanotech >= 1 AND topic >= 1,
    then year-quota-samples ~ARTICLES_PER_YEAR per year per topic with deficit
    carry-forward.
    """
    futurist = df_all[df_all['Sources'] == 6].copy()
    futurist['Date'] = pd.to_datetime(
        futurist['Original Date'], format='%d %B %Y', errors='coerce'
    )
    futurist = futurist.dropna(subset=['Date'])

    qual_samples = []
    for topic_label, topic_col in QUAL_TOPICS.items():
        if topic_col not in futurist.columns:
            print(f"Column '{topic_col}' not found, skipping {topic_label}")
            continue

        matched = futurist[
            (futurist['Nanotech'] >= 1) & (futurist[topic_col] >= 1)
        ].copy().sort_values('Date')

        if len(matched) == 0:
            continue

        sampled = sample_with_quota(matched, ARTICLES_PER_YEAR, year_min, year_max)
        for _, row in sampled.iterrows():
            qual_samples.append({
                'Topic': topic_label,
                'Year': row['Year'],
                'Month': row['Date'].strftime('%B %Y'),
                'Date': row['Date'].strftime('%d %B %Y'),
                'Title': row['Title'],
                'Body': str(row['Body'])[:BODY_CHAR_LIMIT],
                'Word Count': row['Word count'],
                'Nanotech Mentions': row['Nanotech'],
                f'{topic_label} Mentions': row[topic_col],
            })

    return pd.DataFrame(qual_samples)
