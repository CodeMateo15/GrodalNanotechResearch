"""Anthropic SDK wrapper: forced-tool-use call with prompt caching + retry."""

import hashlib
import json
import os
import random
import time
from pathlib import Path

import anthropic

from .prompts import build_system_blocks
from .schema import CODING_TOOL


DEFAULT_MODEL = 'claude-sonnet-4-6'
MAX_TOKENS = 1024  # record_coding output is ~200-400 tokens; buffer
MAX_RETRIES = 4
BASE_BACKOFF_SEC = 2.0


class CodingClient:
    """Thin wrapper around anthropic.Anthropic that forces a `record_coding`
    tool call, caches the rubric prefix, and caches results on disk by
    (model, sha256(body)) so reruns don't re-pay.
    """

    def __init__(self, model=DEFAULT_MODEL, cache_dir=None):
        # Anthropic() reads ANTHROPIC_API_KEY from env. Raise a clear error if missing.
        if not os.environ.get('ANTHROPIC_API_KEY'):
            raise RuntimeError(
                "ANTHROPIC_API_KEY is not set. Export it before running, e.g.:\n"
                "  (PowerShell) $env:ANTHROPIC_API_KEY = 'sk-ant-...'\n"
                "  (bash)       export ANTHROPIC_API_KEY=sk-ant-..."
            )
        self.client = anthropic.Anthropic()
        self.model = model
        self.system_blocks = build_system_blocks()
        self.cache_dir = Path(cache_dir) if cache_dir else (
            Path(__file__).parent / 'cache'
        )
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def _cache_path(self, title, body):
        key = hashlib.sha256(
            (self.model + '\n' + title + '\n' + body).encode('utf-8')
        ).hexdigest()
        return self.cache_dir / f'{key}.json'

    def code_article(self, title, body):
        """Return (coding_dict, usage_like_obj). Cached on disk by body hash."""
        cache_path = self._cache_path(title, body)
        if cache_path.exists():
            with open(cache_path, 'r', encoding='utf-8') as f:
                blob = json.load(f)
            return blob['coding'], _UsageReplay(blob['usage'])

        user_text = (
            f"Article to code.\n\nTitle: {title}\n\nBody excerpt:\n{body}\n\n"
            "Call record_coding with your codings now."
        )

        coding, usage = self._call_with_retry(user_text)

        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump({
                'coding': coding,
                'usage': {
                    'input_tokens': usage.input_tokens,
                    'output_tokens': usage.output_tokens,
                    'cache_creation_input_tokens': getattr(usage, 'cache_creation_input_tokens', 0) or 0,
                    'cache_read_input_tokens': getattr(usage, 'cache_read_input_tokens', 0) or 0,
                },
            }, f, indent=2)
        return coding, usage

    def _call_with_retry(self, user_text):
        last_exc = None
        for attempt in range(MAX_RETRIES):
            try:
                response = self.client.messages.create(
                    model=self.model,
                    max_tokens=MAX_TOKENS,
                    system=self.system_blocks,
                    tools=[CODING_TOOL],
                    tool_choice={'type': 'tool', 'name': 'record_coding'},
                    messages=[{'role': 'user', 'content': user_text}],
                )
                for block in response.content:
                    if block.type == 'tool_use' and block.name == 'record_coding':
                        return block.input, response.usage
                raise RuntimeError(
                    f"Model did not call record_coding. stop_reason={response.stop_reason}"
                )
            except (anthropic.RateLimitError, anthropic.APIStatusError,
                    anthropic.APIConnectionError) as e:
                last_exc = e
                if isinstance(e, anthropic.APIStatusError) and e.status_code < 500:
                    raise
                sleep_for = BASE_BACKOFF_SEC * (2 ** attempt) + random.uniform(0, 1)
                print(f"    retry {attempt + 1}/{MAX_RETRIES} in {sleep_for:.1f}s: {e}")
                time.sleep(sleep_for)
        raise last_exc


class _UsageReplay:
    """Shape-compatible stand-in for response.usage when replaying from disk."""
    def __init__(self, d):
        self.input_tokens = d.get('input_tokens', 0)
        self.output_tokens = d.get('output_tokens', 0)
        self.cache_creation_input_tokens = d.get('cache_creation_input_tokens', 0)
        self.cache_read_input_tokens = d.get('cache_read_input_tokens', 0)
