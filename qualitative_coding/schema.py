"""Tool-input schema for the forced `record_coding` tool.

Strict JSON schema — the Anthropic API validates tool inputs against this
before returning, so we get back a guaranteed-valid dict.
"""

CODING_TOOL = {
    'name': 'record_coding',
    'description': (
        'Record your qualitative coding of the article on four dimensions. '
        'You MUST call this tool exactly once per article with all fields populated.'
    ),
    'strict': True,
    'input_schema': {
        'type': 'object',
        'properties': {
            'attitude': {
                'type': 'string',
                'enum': ['positive', 'neutral', 'negative', 'mixed'],
                'description': (
                    "The article's overall attitude toward nanotechnology. "
                    "positive: enthusiastic/optimistic; negative: skeptical/concerned/critical; "
                    "mixed: both positive and negative framings present; neutral: descriptive only."
                ),
            },
            'attitude_rationale': {
                'type': 'string',
                'description': 'One short sentence quoting or paraphrasing the text that drove the attitude judgment.',
            },
            'analogy_present': {
                'type': 'boolean',
                'description': 'True if the article uses an analogy, metaphor, or comparison to explain or frame nanotechnology.',
            },
            'analogy_type': {
                'type': 'string',
                'description': (
                    "Short phrase describing the analogy (e.g. 'biological/cellular machinery', "
                    "'industrial revolution', 'science fiction', 'computer chip miniaturization'). "
                    "Use 'none' if analogy_present is false."
                ),
            },
            'analogy_temporality': {
                'type': 'string',
                'enum': ['past', 'present', 'future', 'timeless', 'NA'],
                'description': (
                    "When the analogy's referent sits in time. past: historical event/era; "
                    "future: speculative/anticipated; timeless: ahistorical (biology, physics); "
                    "NA: no analogy present."
                ),
            },
            'sustaining_vs_disrupting': {
                'type': 'string',
                'enum': ['sustaining', 'disrupting', 'both', 'neither'],
                'description': (
                    'Is nanotech framed as sustaining existing industries/practices, '
                    'or disrupting them? both: article argues both; neither: not framed this way.'
                ),
            },
            'funding_argument_present': {
                'type': 'boolean',
                'description': 'True if the article makes any argument about nanotech funding (public or private).',
            },
            'funding_argument_stance': {
                'type': 'string',
                'enum': ['pro-funding', 'anti-funding', 'descriptive', 'NA'],
                'description': (
                    'pro-funding: argues for more funding; anti-funding: argues against; '
                    'descriptive: reports on funding without taking a stance; NA: no funding argument.'
                ),
            },
            'funding_argument_summary': {
                'type': 'string',
                'description': (
                    "One short sentence summarizing the funding argument. "
                    "Empty string if funding_argument_present is false."
                ),
            },
        },
        'required': [
            'attitude', 'attitude_rationale',
            'analogy_present', 'analogy_type', 'analogy_temporality',
            'sustaining_vs_disrupting',
            'funding_argument_present', 'funding_argument_stance', 'funding_argument_summary',
        ],
        'additionalProperties': False,
    },
}

CODING_COLUMNS = [
    'Attitude', 'Attitude Rationale',
    'Analogy Present', 'Analogy Type', 'Analogy Temporality',
    'Sustaining vs Disrupting',
    'Funding Argument Present', 'Funding Argument Stance', 'Funding Argument Summary',
    'Model Used', 'Tokens In', 'Tokens Out', 'Cache Hit',
]


def coding_to_row(coding, model, usage):
    """Flatten a `record_coding` tool input + usage block into row columns."""
    cache_read = getattr(usage, 'cache_read_input_tokens', 0) or 0
    return {
        'Attitude': coding['attitude'],
        'Attitude Rationale': coding['attitude_rationale'],
        'Analogy Present': bool(coding['analogy_present']),
        'Analogy Type': coding['analogy_type'],
        'Analogy Temporality': coding['analogy_temporality'],
        'Sustaining vs Disrupting': coding['sustaining_vs_disrupting'],
        'Funding Argument Present': bool(coding['funding_argument_present']),
        'Funding Argument Stance': coding['funding_argument_stance'],
        'Funding Argument Summary': coding['funding_argument_summary'],
        'Model Used': model,
        'Tokens In': usage.input_tokens,
        'Tokens Out': usage.output_tokens,
        'Cache Hit': cache_read > 0,
    }
