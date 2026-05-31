"""
TempCleanupMiddleware
---------------------
Periodically removes upload and pickle directories for sessions that have
expired or whose session file no longer exists.

The cleanup runs AT MOST once every CLEANUP_INTERVAL_SECONDS (default 3600 s /
1 hour) so it never adds meaningful overhead to regular requests.  A small
sentinel file (temp_data/.last_cleanup) records the last run time.
"""

import shutil
import time
from pathlib import Path

from django.conf import settings

CLEANUP_INTERVAL_SECONDS = getattr(settings, 'TEMP_CLEANUP_INTERVAL', 3600)


class TempCleanupMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response
        self._sentinel = Path(settings.TEMP_DATA_DIR) / '.last_cleanup'

    def __call__(self, request):
        self._maybe_cleanup()
        return self.get_response(request)

    def _maybe_cleanup(self):
        now = time.time()
        try:
            if self._sentinel.exists():
                last_run = self._sentinel.stat().st_mtime
                if now - last_run < CLEANUP_INTERVAL_SECONDS:
                    return  # Too soon — skip
        except OSError:
            pass

        # Update sentinel first so concurrent workers don't all pile in.
        try:
            self._sentinel.touch()
        except OSError:
            return

        self._cleanup()

    def _cleanup(self):
        max_age = getattr(settings, 'SESSION_COOKIE_AGE', 86400)
        now = time.time()
        sessions_dir = Path(settings.SESSION_FILE_PATH)

        for sub in ('uploads', 'pickles'):
            base = Path(settings.TEMP_DATA_DIR) / sub
            if not base.exists():
                continue
            for session_dir in base.iterdir():
                if not session_dir.is_dir():
                    continue
                session_key = session_dir.name
                session_file = sessions_dir / session_key
                try:
                    if not session_file.exists():
                        shutil.rmtree(session_dir, ignore_errors=True)
                    elif now - session_file.stat().st_mtime > max_age:
                        shutil.rmtree(session_dir, ignore_errors=True)
                        session_file.unlink(missing_ok=True)
                except OSError:
                    pass
