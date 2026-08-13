"""桌面GUI的后台任务支持组件。"""

from .worker import GuiBatchWorker, WorkerEvent

__all__ = ["GuiBatchWorker", "WorkerEvent"]
