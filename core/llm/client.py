"""Thin wrapper around the local LM Studio OpenAI-compatible chat endpoint.

Never raises out to callers that can't afford to fail loudly (bot email sends,
chat replies) — see call_chat's `raise_on_error` param. Every caller must decide
explicitly whether a failure should propagate or degrade silently.

LM Studio serves one generation at a time — a second concurrent request just
sits queued behind the first rather than erroring, which meant a second caller
could silently burn its whole timeout waiting on someone else's request. A
process-wide lock (_LLM_LOCK) below serializes calls so that's an explicit,
visible condition (LLMBusy) instead of an unexplained timeout.
"""
from __future__ import annotations

import logging
import re
import threading

import requests
from django.conf import settings

logger = logging.getLogger(__name__)

_THINK_TAG_RE = re.compile(r'<think>.*?</think>', re.DOTALL | re.IGNORECASE)
_LLM_LOCK = threading.Lock()


class LLMUnavailable(Exception):
    """Raised by call_chat(..., raise_on_error=True) on any failure: network,
    timeout, non-200, or malformed response body."""


class LLMBusy(Exception):
    """Raised by call_chat(..., lock_wait=...) when another call is already
    in flight and lock_wait seconds passed without it finishing. Raised
    regardless of raise_on_error, since it's a distinct condition callers
    that pass lock_wait have explicitly opted to detect."""


def _strip_think_tags(text: str) -> str:
    return _THINK_TAG_RE.sub('', text).strip()


def _ai_enabled() -> bool:
    """AISettings DB row (editable from Settings, no restart needed) wins;
    the LLM_ENABLED env var is the fallback for servers set up before that
    UI existed — same pattern as EmailSettings/DynamicGmailEmailBackend."""
    try:
        from core.models import AISettings
        row = AISettings.objects.first()
        return row.ai_enabled if row is not None else settings.LLM_ENABLED
    except Exception:
        return settings.LLM_ENABLED


def call_chat(
    messages: list[dict],
    *,
    max_tokens: int = 4000,
    temperature: float = 0.2,
    timeout: float | None = None,
    raise_on_error: bool = False,
    lock_wait: float | None = None,
    model: str | None = None,
    return_usage: bool = False,
) -> str | None | tuple[str | None, dict | None]:
    """POST to {LLM_ENDPOINT_URL}/v1/chat/completions. Returns the assistant
    message content, or None on failure (unless raise_on_error=True).

    lock_wait: max seconds to wait for another in-flight call_chat() to
    finish before giving up with LLMBusy. None (default) waits as long as it
    takes — right for background/scheduled callers where no one is watching
    a clock. Pass a short value (e.g. 5) for interactive callers that want to
    fail fast and say "busy" rather than silently eat their own timeout
    waiting behind someone else's request.

    model: override settings.LLM_MODEL_ID for this call only — for
    benchmarking candidate models (see llm_benchmark management command)
    without touching the configured default.

    return_usage: if True, returns (content, usage_dict) instead of just
    content — usage_dict is the API's raw `usage` object (prompt_tokens,
    completion_tokens, completion_tokens_details.reasoning_tokens), or None
    on failure. Only for diagnostics; production callers should leave this
    False and keep the plain string-or-None contract.
    """
    if not _ai_enabled():
        if raise_on_error:
            raise LLMUnavailable('AI is disabled (Settings > AI, or LLM_ENABLED env var)')
        return (None, None) if return_usage else None

    acquired = _LLM_LOCK.acquire(timeout=-1 if lock_wait is None else lock_wait)
    if not acquired:
        raise LLMBusy('Another LLM call is already in progress')

    try:
        url = f'{settings.LLM_ENDPOINT_URL.rstrip("/")}/v1/chat/completions'
        body = {
            'model': model or settings.LLM_MODEL_ID,
            'messages': messages,
            'max_tokens': max_tokens,
            'temperature': temperature,
            'stream': False,
        }
        try:
            response = requests.post(url, json=body, timeout=timeout or settings.LLM_TIMEOUT_SECONDS)
            response.raise_for_status()
            data = response.json()
            content = _strip_think_tags(data['choices'][0]['message']['content'])
            return (content, data.get('usage')) if return_usage else content
        except Exception:
            logger.warning('LLM call failed', exc_info=True)
            if raise_on_error:
                raise LLMUnavailable('LLM call failed') from None
            return (None, None) if return_usage else None
    finally:
        _LLM_LOCK.release()
