# web/csrf.py — lightweight per-session CSRF tokens
import hmac
import secrets
from markupsafe import Markup

_TOKEN_KEY = "csrf_token"


def _get_token(session):
    token = session.get(_TOKEN_KEY)
    if not token:
        token = secrets.token_urlsafe(32)
        session[_TOKEN_KEY] = token
    return token


def csrf_field_markup(session):
    """Hidden input HTML to embed in every POST <form>."""
    return Markup(
        f'<input type="hidden" name="csrf_token" value="{_get_token(session)}">')


def validate(session, form):
    """Constant-time comparison of submitted token vs session token."""
    expected = session.get(_TOKEN_KEY)
    submitted = form.get(_TOKEN_KEY, "")
    if not expected:
        return False
    return hmac.compare_digest(expected, submitted)
