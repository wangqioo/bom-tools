# -*- coding: utf-8 -*-
"""In-process background worker queue for durable PSTX agent runs."""

from __future__ import annotations

from dataclasses import dataclass
import os
import queue
import threading
import time
from typing import Callable, Mapping

from .durable_store import AgentDurableRunStore


BACKGROUND_RUNNER_VERSION = "pstx-agent-background/v1"


def _now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S")


@dataclass(frozen=True)
class AgentBackgroundJob:
    agent_run_id: str
    scope_id: str
    kind: str
    run: Callable[[str], Mapping[str, object]]


class AgentBackgroundRunner:
    """Small process-local queue. It persists run state before execution."""

    def __init__(self,
                 store: AgentDurableRunStore,
                 *,
                 worker_count: int | None = None,
                 queue_limit: int | None = None):
        self.store = store
        self.worker_count = max(1, int(worker_count or os.environ.get("PSTX_AGENT_WORKER_COUNT") or 2))
        self.queue_limit = max(1, int(queue_limit or os.environ.get("PSTX_AGENT_QUEUE_LIMIT") or 16))
        self._queue: "queue.Queue[AgentBackgroundJob]" = queue.Queue(maxsize=self.queue_limit)
        self._started = False
        self._lock = threading.Lock()

    def start(self) -> None:
        with self._lock:
            if self._started:
                return
            self._started = True
            for index in range(self.worker_count):
                thread = threading.Thread(target=self._worker, name=f"pstx-agent-worker-{index + 1}", daemon=True)
                thread.start()

    def submit(self, job: AgentBackgroundJob) -> dict:
        self.start()
        try:
            self._queue.put_nowait(job)
        except queue.Full as exc:
            self.store.fail_record(job.agent_run_id, f"后台 Agent 队列已满，上限 {self.queue_limit}。")
            raise RuntimeError(f"后台 Agent 队列已满，上限 {self.queue_limit}。") from exc
        return self.store.update_record(
            job.agent_run_id,
            status="queued",
            current_phase="queued",
            heartbeat_at=_now(),
            checkpoint={"phase": "queued", "queue_size": self._queue.qsize(), "heartbeat_at": _now()},
            progress={"step_index": 0, "tool_call_count": 0, "evidence_count": 0},
        )

    def cancel(self, agent_run_id: object) -> dict:
        return self.store.mark_cancel_requested(agent_run_id)

    def status(self) -> dict:
        return {
            "version": BACKGROUND_RUNNER_VERSION,
            "started": self._started,
            "worker_count": self.worker_count,
            "queue_limit": self.queue_limit,
            "queue_size": self._queue.qsize(),
        }

    def _worker(self) -> None:
        while True:
            job = self._queue.get()
            try:
                record = self.store.read_record(job.agent_run_id)
                if not record or record.get("cancel_requested") or record.get("status") == "cancelled":
                    self.store.update_record(
                        job.agent_run_id,
                        status="cancelled",
                        current_phase="cancelled",
                        heartbeat_at=_now(),
                        checkpoint={"phase": "cancelled_before_run", "heartbeat_at": _now()},
                    )
                    continue
                self.store.update_record(
                    job.agent_run_id,
                    status="running",
                    current_phase="running",
                    heartbeat_at=_now(),
                    checkpoint={"phase": "running", "kind": job.kind, "heartbeat_at": _now()},
                )
                result = job.run(job.agent_run_id)
                self.store.finish_record(job.agent_run_id, result)
            except Exception as exc:  # pragma: no cover - guarded by route tests.
                self.store.fail_record(job.agent_run_id, exc)
            finally:
                self._queue.task_done()
