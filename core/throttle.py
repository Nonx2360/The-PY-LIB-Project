# core/throttle.py — in-process sliding-window login throttling
#
# Shared by the web app (key = (ip, username)) and the desktop app
# (key = ("local", username)). After MAX_FAILED_AUTH failed logins within
# WINDOW_SECONDS for one key, further attempts are locked until that window
# expires. Keys are dropped after the window to avoid unbounded memory growth.
import time
from collections import defaultdict, deque

MAX_FAILED_AUTH = 5
WINDOW_SECONDS = 15 * 60

_attempts = defaultdict(deque)  # key -> deque of failure timestamps


def _prune(key, now):
    dq = _attempts.get(key)
    if dq is None:
        return 0
    while dq and now - dq[0] > WINDOW_SECONDS:
        dq.popleft()
    if not dq:
        _attempts.pop(key, None)
    return len(dq)


def is_locked(key):
    """True if the key has hit MAX_FAILED_AUTH recent failures."""
    return _prune(key, time.time()) >= MAX_FAILED_AUTH


def lock_remaining_seconds(key):
    _prune(key, time.time())
    dq = _attempts.get(key)
    if dq and len(dq) >= MAX_FAILED_AUTH:
        return max(0, int(WINDOW_SECONDS - (time.time() - dq[0])))
    return 0


def lock_remaining_minutes(key):
    """Whole minutes left in the lockout, minimum 1 (for user-facing text)."""
    return max(1, (lock_remaining_seconds(key) + 59) // 60)


def record_failure(key):
    _prune(key, time.time())
    _attempts[key].append(time.time())


def record_success(key):
    _attempts.pop(key, None)
