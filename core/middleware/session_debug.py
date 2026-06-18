from django.utils.deprecation import MiddlewareMixin
from django.conf import settings
from pathlib import Path
import json


class SessionDebugMiddleware(MiddlewareMixin):
    """Temporary middleware to log authentication and session info for each request.

    Writes a small JSON line to `session_debug.log` in the project root. This
    avoids relying on stdout capture and makes it easy to read from the test
    harness.
    """

    LOG_PATH = Path(settings.BASE_DIR) / 'session_debug.log'

    def process_request(self, request):
        user = getattr(request, 'user', None)
        username = getattr(user, 'username', None) if user else None
        is_auth = bool(getattr(user, 'is_authenticated', False)) if user else False
        try:
            session_key = request.session.session_key
        except Exception:
            session_key = None

        entry = {
            'method': request.method,
            'path': request.path,
            'user': username,
            'authenticated': is_auth,
            'session_key': session_key,
            'cookie': (request.META.get('HTTP_COOKIE') or '')[:200],
        }
        try:
            with open(self.LOG_PATH, 'a', encoding='utf-8') as fh:
                fh.write(json.dumps(entry) + '\n')
        except Exception:
            # Never raise from middleware during debugging
            pass
