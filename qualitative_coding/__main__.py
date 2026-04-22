"""CLI entry point.

Setup:
    1. pip install anthropic
    2. Set ANTHROPIC_API_KEY in your environment:
         PowerShell:  $env:ANTHROPIC_API_KEY = "sk-ant-..."
         bash:        export ANTHROPIC_API_KEY=sk-ant-...
    3. Run create_excelV5.py first to populate 'Qualitative Samples' and
       'AI Qualitative Samples' sheets in output.xlsx.

Usage:
    python -m qualitative_coding
    python -m qualitative_coding --input output.xlsx
    python -m qualitative_coding --sheets "Qualitative Samples,AI Qualitative Samples"
    python -m qualitative_coding --model claude-sonnet-4-6
"""

import argparse

from .client import DEFAULT_MODEL
from .code_samples import run


def main():
    ap = argparse.ArgumentParser(description='Qualitative code sampled articles via Claude.')
    ap.add_argument('--input', default='output.xlsx')
    ap.add_argument('--sheets',
                    default='Qualitative Samples,AI Qualitative Samples',
                    help='Comma-separated sheet names to code.')
    ap.add_argument('--model', default=DEFAULT_MODEL)
    ap.add_argument('--output-suffix', default=' (Coded)')
    args = ap.parse_args()

    sheets = tuple(s.strip() for s in args.sheets.split(',') if s.strip())
    run(xlsx_path=args.input, sheets=sheets, model=args.model,
        output_suffix=args.output_suffix)


if __name__ == '__main__':
    main()
