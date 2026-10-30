# config/env_store.py — single place that owns the project's secret file (.env)
#
# Rules enforced here:
#   .env must NEVER be hardcoded defaults (see BANNED/high-risk values).
#   Missing weak secrets are replaced with freshly generated random values.
#   The file is rewritten ONLY when something needs to change, preserving
#   every key we don't manage.
#
# WARNING (data migration): if SECRET_PASSPHRASE changes, QR payloads that
# were encrypted with the previous key can no longer be decrypted. Existing
# installs must run `tools/rotate_qr_key.py` — see docs/SECURITY.md.
import os
import secrets

ENV_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env")

# Values that must never be accepted: the shipped placeholders and the
# historical hardcoded default from before the refactor.
BANNED_VALUES = {
    "default_secret_py_lib_key_change_me_in_dotenv",
    "pylibsalt123",
    "your_secret_school_library_passphrase_here",
    "your_secret_salt_here",
}

# key -> (minimum length, generator)
_KEYS = {
    "SECRET_PASSPHRASE": (24, lambda: secrets.token_urlsafe(32)),
    "SECRET_SALT": (10, lambda: secrets.token_urlsafe(12)),
    "FLASK_SECRET_KEY": (32, lambda: secrets.token_hex(32)),
}


def _read_env_file(path):
    """Parse a .env-style file into a dict, preserving unknown keys."""
    values = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                values[key.strip()] = val.strip()
    return values


def _write_env_file(path, values):
    lines = [
        "# THE PY-LIB secrets — DO NOT COMMIT THIS FILE (kept in .gitignore).",
        "# If SECRET_PASSPHRASE ever changes, run tools/rotate_qr_key.py afterwards:",
        "# stored QR payloads must be re-encrypted, see docs/SECURITY.md.",
        "",
    ]
    for key, val in values.items():
        lines.append(f"{key}={val}")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    try:  # best effort on POSIX; no-op on Windows
        os.chmod(path, 0o600)
    except OSError:
        pass


def _resolve_path(path):
    return path or os.environ.get("PYLIB_ENV_FILE") or ENV_FILE


def update_secrets(updates=None, path=None):
    """Merge `updates` into the .env file, preserving every other key."""
    path = _resolve_path(path)
    values = _read_env_file(path)
    for key, val in (updates or {}).items():
        values[key] = val
    _write_env_file(path, values)
    return values


def ensure_secrets(path=None, announce=None):
    """Make sure every managed secret exists and is not a known-bad value.

    Returns a dict with all managed keys. Generates + persists new random
    values when a key is missing, too short, or matches a known placeholder.
    `announce` is an optional callable(log_line) used instead of printing.
    Path resolution: explicit arg > PYLIB_ENV_FILE env var > project .env
    (tests/CI set PYLIB_ENV_FILE to a temp file so no real .env is touched).
    """
    path = _resolve_path(path)
    values = _read_env_file(path)

    generated = []
    for key, (min_len, gen) in _KEYS.items():
        current = values.get(key, "")
        if (not current or len(current) < min_len
                or current in BANNED_VALUES):
            values[key] = gen()
            generated.append(key)

    if generated:
        _write_env_file(path, values)
        if announce:
            announce(
                f"Generated new secret(s) {', '.join(generated)} and saved them to {os.path.basename(path)}"
            )
            if any(k != "FLASK_SECRET_KEY" for k in generated):
                announce(
                    "QR encryption material changed: payloads enciphered with the previous key "
                    "will no longer decrypt. Run tools/rotate_qr_key.py BEFORE opening existing "
                    "member data — see docs/SECURITY.md."
                )

    return values
