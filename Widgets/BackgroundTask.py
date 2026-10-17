import traceback

from PySide6.QtCore import QObject, QRunnable, Signal


class TaskSignals(QObject):
    finished = Signal(object)
    failed = Signal(object)


class BackgroundTask(QRunnable):
    """
    Runs fn in the Qt thread pool, so the window stays responsive. Results and errors are delivered through
    signals, which Qt hands over to the main thread.
    """

    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = TaskSignals()

    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
        except Exception as e:
            e.formatted_traceback = traceback.format_exc()
            self.signals.failed.emit(e)
            return
        self.signals.finished.emit(result)
