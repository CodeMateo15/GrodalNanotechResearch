"""Plotting configuration — mirrors notebook cell 1 constants."""

YEAR_MIN = 1983
YEAR_MAX = 2005

STANDARDIZE_Y = True  # True = lock all % graphs to 0-100%

SOURCE_INFO = {
    1: ('Government',        '#1D4ED8'),
    2: ('Science News',      '#0891B2'),
    3: ('Science Research',  '#16A34A'),
    4: ('Business Press',    '#D97706'),
    5: ('Business',          '#DC2626'),
    6: ('Futurists',         '#7C3AED'),
    7: ('Newspapers',        '#EAFF2F'),
}

KEYWORD_COLS = [
    'Space', 'Electronics', 'Artificial Intelligence', 'Photonics',
    'Biotech/Biology', 'Semiconductors', 'Robotics', 'Computers/Computing',
    'Material Science', 'Cleantech', 'Hypertext', 'Internet',
    'Chemistry', 'Physics', 'Engineering', 'Nanotech', 'Molecular Manufacturing',
    'Financial', 'Commerce', 'Application', 'Revolution',
    'Total Electronics/Computing',
]

# For co-occurrence analysis, exclude Nanotech/Nano to avoid self-referential 100%.
COOCCURRENCE_COLS = [c for c in KEYWORD_COLS if c not in ('Nanotech', 'Nano')]

# Qualitative coding columns (from Claude analysis import)
QUALITATIVE_COLS = [
    'Attitude', 'Analogy Present', 'Analogy Temporality (Coded)',
    'Sustaining vs Disrupting', 'Funding Argument Present',
]

ATTITUDE_VALUES = ['positive', 'negative', 'neutral', 'mixed']
TEMPORALITY_VALUES = ['past', 'present', 'future', 'mixed']
SUSTAINING_VALUES = ['sustaining', 'disrupting', 'both']

COLORS = [
    '#7C3AED', '#0891B2', '#D97706', '#059669', '#BE185D',
    '#1D4ED8', '#B45309', '#0F766E', '#9333EA', '#C2410C',
    '#0369A1', '#15803D', '#A21CAF', '#B91C1C', '#047857',
    '#9D174D', '#1E40AF', '#B44C00', '#6D28D9', '#EA580C',
    '#0284C7', '#4F46E5',
]
