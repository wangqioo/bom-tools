"""Process-local Web session state.

Keeping state here lets route modules be split out of ``pstx_web.py`` without
duplicating caches or making the desktop shell depend on the monolithic module.
The data remains in-process only.
"""

from __future__ import annotations

from collections import OrderedDict

from pstx_agent_runtime import AgentBackgroundRunner, AgentDurableRunStore, AgentTraceStore


MAX_RUNS = 12
MAX_AGENT_RUNS = 50

RUN_CACHE: "OrderedDict[str, dict]" = OrderedDict()
AGENT_RUN_CACHE = AgentTraceStore(max_items=MAX_AGENT_RUNS)
AGENT_CONTEXT_CACHE: "OrderedDict[str, dict]" = OrderedDict()
AGENT_DURABLE_STORE = AgentDurableRunStore()
AGENT_BACKGROUND_RUNNER = AgentBackgroundRunner(AGENT_DURABLE_STORE)


def clear_web_session_state() -> None:
    """Clear in-process Web caches for tests or local reset flows."""
    RUN_CACHE.clear()
    AGENT_RUN_CACHE.clear()
    AGENT_CONTEXT_CACHE.clear()
