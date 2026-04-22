"""Export the notebook's visualizations to multi-page PDFs.

Usage:
    python export_figures.py --full       # ~200+ page archive
    python export_figures.py --summary    # ~12 key charts
    python export_figures.py              # writes both
"""

import argparse
from pathlib import Path

import matplotlib
matplotlib.use('Agg')  # headless
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

from nanotech_plots.data import load_data
from nanotech_plots.registry import ALL_PLOTS


def export(output_pdf, df_all, summary_only=False):
    written = 0
    skipped = 0
    with PdfPages(output_pdf) as pdf:
        for name, fn, is_summary in ALL_PLOTS:
            if summary_only and not is_summary:
                continue
            fig = fn(df_all)
            if fig is None:
                skipped += 1
                continue
            pdf.savefig(fig, bbox_inches='tight')
            plt.close(fig)
            written += 1
    print(f"  -> {output_pdf}: {written} figures written"
          + (f", {skipped} skipped" if skipped else ""))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--full', action='store_true', help='Write the full archive PDF')
    ap.add_argument('--summary', action='store_true', help='Write the summary PDF')
    ap.add_argument('--input', default='output.xlsx')
    ap.add_argument('--full-out', default='nanotech_graphs_full.pdf')
    ap.add_argument('--summary-out', default='nanotech_graphs_summary.pdf')
    args = ap.parse_args()

    # If neither flag passed, do both.
    if not args.full and not args.summary:
        args.full = args.summary = True

    script_dir = Path(__file__).parent
    input_path = script_dir / args.input
    print(f"Loading {input_path} ...")
    df_all = load_data(str(input_path))
    print(f"  ({len(df_all)} rows)")

    if args.summary:
        export(script_dir / args.summary_out, df_all, summary_only=True)
    if args.full:
        export(script_dir / args.full_out, df_all, summary_only=False)


if __name__ == '__main__':
    main()
