"""Qualitative coding of sampled articles via the Anthropic API.

Reads 'Qualitative Samples' and 'AI Qualitative Samples' sheets from
output.xlsx, asks Claude to code each article on four dimensions
(Attitude, Analogy, Sustaining vs Disrupting, Funding), and writes the
results back to '<sheet> (Coded)' sheets.

Opt-in: requires ANTHROPIC_API_KEY in the environment. Not invoked
automatically by create_excelV5.py.
"""

def run(*args, **kwargs):
    """Lazy shim — delegates to code_samples.run without importing anthropic
    at package-load time. Lets you `import qualitative_coding` for schema
    inspection without having installed the SDK.
    """
    from .code_samples import run as _run
    return _run(*args, **kwargs)


__all__ = ['run']
