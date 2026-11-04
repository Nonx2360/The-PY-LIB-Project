# web/live.py - realtime WebSocket push hub
"""Live data channel for the web UI and the member portal.

One WebSocket per open browser tab carries only two tiny message types:
a member snapshot (own counts + the three module lists) or an admin
topics ping (which table feeds changed). The heavy lifting stays where
it belongs - the browser either re-renders its DOM locally or asks the
existing JSON feeds again; the server never streams large payloads.

Architecture (deliberately event-free):
  * Mutations happen in two processes: the web app itself and the
    desktop app. Cross-process plumbing can miss writes, so instead a
    single background ticker re-reads cheap DB markers (counts, max id)
    every TICK_INTERVAL seconds and diffs them against the last check.
    Any change fans out to the live sockets - members get a fresh
    snapshot, admins get the changed topic list. Worst-case staleness
    is one tick, which reads as "instant" at school scale.
  * Data sources are the same repositories the page handlers use, so
    tests can repoint DB_PATH and connected sockets keep working.
"""
import json
import threading

from core.logger import logger


TICK_INTERVAL = 2.0        # seconds between DB marker checks
PING_INTERVAL = 30.0       # server-side read timeout; client pings at 20s


class LiveHub:
    """Registry of open sockets + serialized per-socket send."""

    def __init__(self):
        self._lock = threading.Lock()
        self._conns = {}       # ws object -> entry dict

    def register(self, kind, ws, key=None, payload_fn=None):
        entry = {
            "kind": kind, "key": key, "payload_fn": payload_fn,
            "send_lock": threading.Lock(),
        }
        with self._lock:
            self._conns[ws] = entry
        return entry

    def unregister(self, ws):
        with self._lock:
            self._conns.pop(ws, None)

    def counts(self):
        with self._lock:
            members = sum(1 for e in self._conns.values() if e["kind"] == "member")
            admins = len(self._conns) - members
        return members, admins

    def send_json(self, ws, payload):
        with self._lock:
            entry = self._conns.get(ws)
            if entry is None:
                return False
            entry = entry
        try:
            with entry["send_lock"]:
                ws.send(json.dumps(payload, ensure_ascii=False))
            return True
        except Exception:
            self.unregister(ws)
            return False

    def broadcast_members(self, payload_fn):
        """Push a fresh snapshot to every open member socket."""
        with self._lock:
            targets = [(ws, e["key"]) for ws, e in self._conns.items()
                       if e["kind"] == "member"]
        for ws, key in targets:
            with self._lock:
                e = self._conns.get(ws)
                fn = e["payload_fn"] if e else payload_fn
            try:
                data = (fn or payload_fn)(key)
            except Exception:
                # repo hiccup (e.g. transient lock): skip this beat; the
                # next tick simply re-pushes the same state
                logger.debug("web live ws: member fan-out skipped (transient db)")
                continue
            self.send_json(ws, data)

    def broadcast_admins(self, payload):
        with self._lock:
            targets = [ws for ws, e in self._conns.items()
                       if e["kind"] == "admin"]
        for ws in targets:
            self.send_json(ws, payload)


class LiveTicker:
    """Diff cheap DB markers and fan changes out to the hub.

    `topics` maps topic name -> callable returning a JSON-able tuple of
    marker values; a changed tuple means that topic changed. The commit
    of the diff + pushes happens through `tick_once` so tests can drive
    it synchronously without any thread.
    """

    def __init__(self, hub, topics, interval=TICK_INTERVAL,
                 member_payload_fn=None):
        self.hub = hub
        self.topics = topics
        self.interval = interval
        self.member_payload_fn = member_payload_fn
        self._stop = threading.Event()
        self._thread = None
        self._prev = {}
        self._prev = self._read_markers()

    # -- marker reading -------------------------------------------------
    def _read_markers(self):
        markers = {}
        for name, fn in self.topics.items():
            try:
                markers[name] = tuple(fn() or ())
            except Exception:
                # transient read failure (busy DB): report the topic
                # unchanged so one hiccup never triggers a bogus broadcast
                markers[name] = self._prev.get(name)
        return markers

    def tick_once(self):
        """One diff beat; returns the set of changed topic names."""
        current = self._read_markers()
        changed = sorted(
            name for name in self.topics
            if current.get(name) != self._prev.get(name))
        self._prev = current
        if not changed:
            return changed
        members, admins = self.hub.counts()
        if admins:
            self.hub.broadcast_admins(
                {"type": "topics", "topics": changed})
        if members and self.member_payload_fn is not None:
            self.hub.broadcast_members(self.member_payload_fn)
        return changed

    # -- lifecycle --------------------------------------------------------
    def ensure_running(self):
        if self._thread is not None and self._thread.is_alive():
            return
        self._stop.clear()

        def _loop():
            while not self._stop.wait(self.interval):
                try:
                    self.tick_once()
                except Exception:
                    # never let the ticker die silently (no member ids in
                    # logs, the tick itself carries none)
                    pass

        self._thread = threading.Thread(target=_loop, name="pylib-live-tick",
                                        daemon=True)
        self._thread.start()

    def stop(self):
        self._stop.set()
