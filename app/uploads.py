"""Bounded, private temporary copies of workbooks already inspected by the UI."""

from dataclasses import dataclass
import re
import secrets
import tempfile
import threading
import time
from typing import BinaryIO


@dataclass(frozen=True)
class InspectedUpload:
    filename: str
    data: bytes


@dataclass
class _Entry:
    filename: str
    file: BinaryIO
    size: int
    deadline: float


class InspectionCache:
    def __init__(self, max_bytes: int = 64 * 1024 * 1024, max_files: int = 20,
                 ttl: float = 15 * 60):
        self.max_bytes, self.max_files, self.ttl = max_bytes, max_files, ttl
        self.lock = threading.Lock()
        self.entries: dict[str, _Entry] = {}

    def _discard(self, token):
        entry = self.entries.pop(token, None)
        if entry:
            entry.file.close()

    def _expire(self):
        now = time.monotonic()
        for token, entry in list(self.entries.items()):
            if entry.deadline <= now:
                self._discard(token)

    def put(self, filename: str, data: bytes) -> str | None:
        with self.lock:
            self._expire()
            if len(data) > self.max_bytes or self.max_files < 1:
                return None
            used = sum(entry.size for entry in self.entries.values())
            while self.entries and (len(self.entries) >= self.max_files or used + len(data) > self.max_bytes):
                oldest = next(iter(self.entries))
                used -= self.entries[oldest].size
                self._discard(oldest)
            upload = None
            try:
                # TemporaryFile is private and unlinked; process exit also deletes it.
                upload = tempfile.TemporaryFile(prefix="qcp-inspected-")
                upload.write(data)
                upload.flush()
            except OSError:
                if upload:
                    upload.close()
                return None  # A full disk must not prevent ordinary uploads.
            token = secrets.token_urlsafe(32)
            self.entries[token] = _Entry(filename, upload, len(data), time.monotonic() + self.ttl)
            return token

    def get(self, token: str) -> InspectedUpload | None:
        if not re.fullmatch(r"[A-Za-z0-9_-]{43}", token):
            return None
        with self.lock:
            self._expire()
            entry = self.entries.get(token)
            if entry:
                try:
                    entry.file.seek(0)
                    return InspectedUpload(entry.filename, entry.file.read())
                except OSError:
                    self._discard(token)
            return None

    def discard(self, token: str):
        with self.lock:
            self._discard(token)

    def expire(self):
        with self.lock:
            self._expire()

    def close(self):
        with self.lock:
            for token in list(self.entries):
                self._discard(token)
