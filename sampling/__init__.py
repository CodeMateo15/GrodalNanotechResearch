"""Year-quota sampling for qualitative article selection.

Shared by create_excelV5.py (to auto-populate the 'Qualitative Samples' and
'AI Qualitative Samples' sheets in output.xlsx) and by nanotech_graphs.ipynb
(for interactive exploration). Keeping the implementation in one place keeps
notebook and pipeline outputs in sync.
"""

from .quota import sample_with_quota
from .qualitative import build_qualitative_samples
from .ai_qualitative import build_ai_qualitative_samples

__all__ = [
    'sample_with_quota',
    'build_qualitative_samples',
    'build_ai_qualitative_samples',
]
