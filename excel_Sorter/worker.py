"""Background worker for batch Excel processing."""
import threading
from typing import Callable, List

class BatchWorker(threading.Thread):
    """Threaded worker for processing a list of file paths.
    callback signature:
        progress_cb(idx:int, total:int, path:str, state:str)
    states: "started", "locked", "loaded", "sorted", "saved", "error", "done" """
    def __init__(self, paths: List[str], handler_cls, callback: Callable):
        super().__init__(daemon=True)
        self.paths = list(paths)
        self.handler_cls = handler_cls
        self.callback = callback
        self._stop = False

    def stop(self):
        """Request stop (best effort)."""
        self._stop = True

    def run(self):
        total = len(self.paths)
        for idx, path in enumerate(self.paths, start=1):
            if self._stop:
                break
            self.callback(idx, total, path, "started")
            try:
                handler = self.handler_cls(path)
                loaded = handler.load_workbook()
            except Exception as exc:  # pragma: no cover - top-level safety
                self.callback(idx, total, path, f"error:{exc}")
                continue

            if getattr(handler, "file_open_locked", False):
                self.callback(idx, total, path, "locked")
                continue

            if not loaded:
                self.callback(idx, total, path, "error:load_failed")
                continue
            self.callback(idx, total, path, "loaded")

            try:
                ok = handler.sort_sheets_alphabetically()
                if not ok:
                    self.callback(idx, total, path, "error:sort_failed")
                    continue
                self.callback(idx, total, path, "sorted")
            except Exception as exc:  # pragma: no cover
                self.callback(idx, total, path, f"error:{exc}")
                continue
            # saving left to caller (UI) via handler methods if needed
            self.callback(idx, total, path, "done")
        # finished
        self.callback(total, total, "", "finished")
