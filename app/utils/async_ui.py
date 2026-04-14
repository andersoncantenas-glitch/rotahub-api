import logging
import threading


def run_async_ui(widget, work, on_success, on_error=None):
    """Executa trabalho em background e devolve resultado no thread da UI."""

    def _deliver(callback, *args):
        if not callback:
            return
        try:
            if widget.winfo_exists():
                widget.after(0, lambda: widget.winfo_exists() and callback(*args))
        except Exception:
            logging.debug("Falha ao entregar resultado async para a UI", exc_info=True)

    def _worker():
        try:
            result = work()
        except Exception as exc:
            _deliver(on_error, exc)
            return
        _deliver(on_success, result)

    threading.Thread(target=_worker, daemon=True).start()
