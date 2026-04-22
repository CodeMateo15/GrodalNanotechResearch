"""Drive the coding loop: read a sample sheet, code each row, write the coded sheet."""

import pandas as pd

from .client import CodingClient, DEFAULT_MODEL
from .schema import CODING_COLUMNS, coding_to_row


def code_sheet(xlsx_path, sheet_name, client):
    """Read `sheet_name`, code each row, return a DataFrame with coding columns appended."""
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    if len(df) == 0:
        print(f"  {sheet_name}: empty, skipping")
        return df

    print(f"  {sheet_name}: {len(df)} rows")
    coded_rows = []
    total_in = total_out = cache_hits = 0
    errors = 0

    for i, row in df.iterrows():
        title = str(row.get('Title', ''))
        body = str(row.get('Body', ''))
        try:
            coding, usage = client.code_article(title, body)
            coded_rows.append(coding_to_row(coding, client.model, usage))
            total_in += usage.input_tokens
            total_out += usage.output_tokens
            if getattr(usage, 'cache_read_input_tokens', 0) or 0:
                cache_hits += 1
            if (i + 1) % 10 == 0 or i == len(df) - 1:
                print(f"    {i + 1}/{len(df)}  in={total_in} out={total_out} cache_hits={cache_hits}")
        except Exception as e:
            errors += 1
            coded_rows.append({c: f'ERROR: {e.__class__.__name__}' for c in CODING_COLUMNS})
            print(f"    {i + 1}/{len(df)}  ERROR: {e}")

    coded_df = pd.DataFrame(coded_rows, index=df.index)
    combined = pd.concat([df, coded_df], axis=1)
    print(f"    done. total in={total_in} out={total_out} cache_hits={cache_hits}/{len(df)} errors={errors}")
    return combined


def run(xlsx_path='output.xlsx',
        sheets=('Qualitative Samples', 'AI Qualitative Samples'),
        model=DEFAULT_MODEL,
        output_suffix=' (Coded)'):
    client = CodingClient(model=model)
    results = {}
    for sheet in sheets:
        print(f"\nCoding sheet: {sheet}")
        results[sheet + output_suffix] = code_sheet(xlsx_path, sheet, client)

    with pd.ExcelWriter(xlsx_path, engine='openpyxl',
                        mode='a', if_sheet_exists='replace') as writer:
        for name, df in results.items():
            df.to_excel(writer, sheet_name=name, index=False)
            print(f"Wrote sheet '{name}' ({len(df)} rows)")
