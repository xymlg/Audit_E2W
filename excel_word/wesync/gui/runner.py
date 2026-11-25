"""后台任务执行器：在子线程里跑命令，把日志与结果投递到界面队列。"""

from __future__ import annotations

import queue
import threading
import time
import traceback
from typing import Callable, Iterable

from ..console import reset_log_state, set_log_sink
from ..errors import SyncError

Step = tuple[str, Callable[[], int]]


class TaskRunner:
    """串行执行若干步骤；同一时刻只允许一个任务。"""

    def __init__(self) -> None:
        self.queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.busy = False

    def submit(self, steps: Iterable[Step]) -> bool:
        steps = list(steps)
        if self.busy or not steps:
            return False
        self.busy = True

        def worker() -> None:
            started = time.perf_counter()
            done: list[tuple[str, int]] = []
            reset_log_state()
            set_log_sink(lambda message: self.queue.put(("log", message)))
            try:
                for title, func in steps:
                    self.queue.put(("step", title))
                    code = func()
                    if not isinstance(code, int):
                        code = 0  # 步骤返回摘要等非整数视为成功，避免误报执行失败
                    done.append((title, code))
                    if code not in (0, None):
                        self.queue.put(("log", f"步骤「{title}」返回码 {code}，已停止后续步骤"))
                        break
                self.queue.put(("done", {"steps": done, "elapsed": time.perf_counter() - started}))
            except SyncError as exc:
                self.queue.put(
                    ("error", {"title": steps[-1][0], "message": str(exc), "elapsed": time.perf_counter() - started})
                )
            except Exception as exc:  # 未预期异常：把堆栈写进日志便于排查
                self.queue.put(
                    (
                        "error",
                        {
                            "title": steps[-1][0],
                            "message": f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}",
                            "elapsed": time.perf_counter() - started,
                        },
                    )
                )
            finally:
                set_log_sink(None)
                self.busy = False

        threading.Thread(target=worker, name="wesync-gui-task", daemon=True).start()
        return True

    def drain(self) -> list[tuple[str, object]]:
        items: list[tuple[str, object]] = []
        while True:
            try:
                items.append(self.queue.get_nowait())
            except queue.Empty:
                return items
