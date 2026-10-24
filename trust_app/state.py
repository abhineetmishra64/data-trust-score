"""The single record of what is loaded. Loads run in a thread, the page polls this."""
import threading

_lock = threading.Lock()
_state = {"status": "idle", "message": "Choose where to read the Trust Score from.",
          "rows": [], "raw": None, "loaded_at": None, "source": None}


def update(**values):
    with _lock:
        _state.update(values)


def get(*keys):
    with _lock:
        return {k: _state[k] for k in keys} if keys else dict(_state)


def is_loading():
    return get("status")["status"] == "loading"
