"""Import Claude analysis columns from the old integration file into output.xlsx.

Reads 'Big dataset' from 'Old Python Files/Claude integration of output file.xlsx',
matches rows to output.xlsx Sheet1 by Title+Sources key, and imports the 10 Claude
analysis columns. Also creates normalized columns for graphing.

Creates a backup of output.xlsx before modifying it.
"""

import re
import shutil
from pathlib import Path

import pandas as pd

OLD_FILE = Path('Old Python Files/Claude integration of output file.xlsx')
OUTPUT_FILE = Path('output.xlsx')
BACKUP_FILE = Path('output_backup.xlsx')

# The 10 Claude analysis columns in the old file (cols 32-41)
CLAUDE_COLS = [
    'Source Analyzed',
    'Perspective on Technology',
    'Mentions Nanotechnology',
    'Sentiment Toward Nanotechnology',
    'Analogies',
    'Purpose of Analogies',
    'Temporality of Analogies',
    'Analogy: Status Quo',
    'Analogy: Funding Argument',
    'Actors Mentioned',
]


def make_key(df):
    """Build a composite match key from Title + Sources."""
    title = df['Title'].astype(str).str.strip()
    sources = df['Sources'].astype(str).str.strip()
    return title + '|' + sources


def normalize_temporality(raw):
    """Map free-text temporality to coded categories.

    Values like 'Already happening' -> 'present',
    'Near future' / 'Distant future' -> 'future',
    combos with multiple -> 'mixed',
    null -> NaN.
    """
    if pd.isna(raw):
        return pd.NA

    text = str(raw).lower().strip()
    if not text:
        return pd.NA

    has_past = bool(re.search(r'\bpast\b|\bhistor', text))
    has_present = bool(re.search(r'\balready happening\b|\bcurrent\b|\bpresent\b|\bongoing\b', text))
    has_future = bool(re.search(r'\bfuture\b|\bnear future\b|\bdistant future\b', text))

    categories = []
    if has_past:
        categories.append('past')
    if has_present:
        categories.append('present')
    if has_future:
        categories.append('future')

    if len(categories) == 0:
        # Try harder: check for numbered sections indicating multiple temporalities
        if ';' in text or re.search(r'\d\)', text):
            return 'mixed'
        # Default to present if text exists but doesn't match patterns
        return 'present'
    elif len(categories) == 1:
        return categories[0]
    else:
        return 'mixed'


def normalize_columns(df):
    """Add normalized columns for graphing from the imported Claude columns."""

    # Attitude: lowercase version of Sentiment
    if 'Sentiment Toward Nanotechnology' in df.columns:
        df['Attitude'] = df['Sentiment Toward Nanotechnology'].apply(
            lambda x: str(x).lower().strip() if pd.notna(x) else pd.NA
        )

    # Analogy Present: 1 if Analogies has text, 0 if Sentiment exists but no analogy, NaN otherwise
    if 'Analogies' in df.columns and 'Sentiment Toward Nanotechnology' in df.columns:
        has_sentiment = df['Sentiment Toward Nanotechnology'].notna()
        has_analogy = df['Analogies'].notna()
        df['Analogy Present'] = pd.NA
        df.loc[has_sentiment, 'Analogy Present'] = 0
        df.loc[has_analogy, 'Analogy Present'] = 1

    # Analogy Temporality (Coded): mapped from free text
    if 'Temporality of Analogies' in df.columns:
        df['Analogy Temporality (Coded)'] = df['Temporality of Analogies'].apply(
            normalize_temporality
        )

    # Sustaining vs Disrupting: map old values
    if 'Analogy: Status Quo' in df.columns:
        status_map = {
            'Reinforcing': 'sustaining',
            'Disrupting': 'disrupting',
            'Both': 'both',
        }
        df['Sustaining vs Disrupting'] = df['Analogy: Status Quo'].map(status_map)

    # Funding Argument Present: 1/0 from Yes/No
    if 'Analogy: Funding Argument' in df.columns:
        funding_map = {'Yes': 1, 'No': 0}
        df['Funding Argument Present'] = df['Analogy: Funding Argument'].map(funding_map)

    return df


def main():
    # Step 1: Backup
    print(f'Backing up {OUTPUT_FILE} -> {BACKUP_FILE}')
    shutil.copy2(OUTPUT_FILE, BACKUP_FILE)

    # Step 2: Load both files
    print(f'Loading old file: {OLD_FILE}')
    old_df = pd.read_excel(OLD_FILE, sheet_name='Big dataset')
    print(f'  Old file: {len(old_df)} rows, {len(old_df.columns)} columns')

    print(f'Loading current file: {OUTPUT_FILE}')
    new_df = pd.read_excel(OUTPUT_FILE, sheet_name='Sheet1')
    print(f'  Current file: {len(new_df)} rows, {len(new_df.columns)} columns')

    # Step 3: Build match keys
    old_df['_key'] = make_key(old_df)
    new_df['_key'] = make_key(new_df)

    # Dedup old file on key (keep first occurrence)
    old_deduped = old_df.drop_duplicates(subset='_key', keep='first')
    print(f'  Old file unique keys: {len(old_deduped)} / {len(old_df)}')

    # Step 4: Build lookup from old file
    # Only keep Claude columns + key
    available_cols = [c for c in CLAUDE_COLS if c in old_deduped.columns]
    print(f'  Claude columns found: {available_cols}')
    old_lookup = old_deduped.set_index('_key')[available_cols]

    # Step 5: Merge - left join on new_df
    merged = new_df.merge(old_lookup, left_on='_key', right_index=True, how='left')

    # Drop the temp key column
    merged.drop(columns=['_key'], inplace=True)

    # Step 6: Normalize columns for graphing
    merged = normalize_columns(merged)

    # Step 7: Report statistics
    matched = merged['Sentiment Toward Nanotechnology'].notna().sum()
    analogies = merged['Analogies'].notna().sum()
    attitude = merged['Attitude'].notna().sum()
    analogy_present = (merged.get('Analogy Present') == 1).sum()
    funding = merged['Funding Argument Present'].notna().sum()
    sustaining = merged['Sustaining vs Disrupting'].notna().sum()

    print(f'\n--- Import Statistics ---')
    print(f'Total rows in output: {len(merged)}')
    print(f'Rows with Sentiment:          {matched}')
    print(f'Rows with Analogies (text):    {analogies}')
    print(f'Rows with Attitude (norm):     {attitude}')
    print(f'Rows with Analogy Present:     {analogy_present} yes, {(merged.get("Analogy Present") == 0).sum()} no')
    print(f'Rows with Funding Argument:    {funding}')
    print(f'Rows with Sustaining/Disrupt:  {sustaining}')

    # Per-source breakdown
    print(f'\n--- Per-Source Sentiment Coverage ---')
    source_names = {1: 'Government', 2: 'Science News', 3: 'Science Research',
                    4: 'Business Press', 5: 'Business', 6: 'Futurists', 7: 'Newspapers'}
    for src in sorted(source_names):
        mask = merged['Sources'] == src
        total = mask.sum()
        coded = (mask & merged['Sentiment Toward Nanotechnology'].notna()).sum()
        print(f'  {source_names[src]:20s}: {coded:5d} / {total:5d} ({100*coded/total:.1f}%)')

    # Step 8: Write back to output.xlsx
    # Read all other sheets first to preserve them
    print(f'\nWriting enriched data back to {OUTPUT_FILE}...')
    other_sheets = {}
    xls = pd.ExcelFile(BACKUP_FILE)
    for sheet_name in xls.sheet_names:
        if sheet_name != 'Sheet1':
            other_sheets[sheet_name] = pd.read_excel(BACKUP_FILE, sheet_name=sheet_name)

    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        merged.to_excel(writer, sheet_name='Sheet1', index=False)
        for sheet_name, sheet_df in other_sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f'Done. {OUTPUT_FILE} updated with {len(merged.columns)} columns.')
    print(f'Backup saved at {BACKUP_FILE}.')


if __name__ == '__main__':
    main()
