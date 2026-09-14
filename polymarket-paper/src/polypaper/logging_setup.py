"""Logging with a redaction filter that runs before anything reaches a handler.

The filter is applied at the root logger, so it also covers log records emitted
by the SDK and any third-party library, not just our own calls.
"""

from __future__ import annotations

import logging
import re
import sys
from typing import Any, Final

__all__ = ["RedactingFilter", "configure_logging"]

_REDACTED: Final[str] = "[REDACTED]"

# Ordered most-specific first. Each pattern targets a shape that should never
# be printed, rather than a variable name, because the leak we care about is
# the value showing up somewhere we did not anticipate.
_PATTERNS: Final[tuple[re.Pattern[str], ...]] = (
    # PEM blocks, including the body.
    re.compile(r"-----BEGIN[^-]*PRIVATE KEY-----.*?-----END[^-]*PRIVATE KEY-----", re.DOTALL),
    # 0x-prefixed 64-hex secp256k1 private keys.
    re.compile(r"\b0x[a-fA-F0-9]{64}\b"),
    # Bare 64-hex, which is the same key without the prefix.
    re.compile(r"\b[a-fA-F0-9]{64}\b"),
    # BIP-39 style mnemonics: 12 or 24 lowercase words in a row.
    re.compile(r"\b(?:[a-z]{3,8}\s+){11}[a-z]{3,8}\b"),
    # key=value / "key": "value" for anything that smells like a credential.
    re.compile(
        r"(?i)\b(private_?key|secret|passphrase|api_?secret|mnemonic|seed|token)"
        r"\b(\s*[:=]\s*)(\"[^\"]*\"|'[^']*'|\S+)"
    ),
)


def _scrub(text: str) -> str:
    for index, pattern in enumerate(_PATTERNS):
        # The last pattern is a key/value pair: keep the key, drop the value.
        if index == len(_PATTERNS) - 1:
            text = pattern.sub(lambda m: f"{m.group(1)}{m.group(2)}{_REDACTED}", text)
        else:
            text = pattern.sub(_REDACTED, text)
    return text


class RedactingFilter(logging.Filter):
    """Scrub anything key-shaped out of the message and its arguments."""

    def filter(self, record: logging.LogRecord) -> bool:
        if isinstance(record.msg, str):
            record.msg = _scrub(record.msg)

        if record.args:
            if isinstance(record.args, dict):
                scrubbed_map: dict[str, Any] = {
                    key: _scrub(value) if isinstance(value, str) else value
                    for key, value in record.args.items()
                }
                record.args = scrubbed_map
            else:
                record.args = tuple(
                    _scrub(value) if isinstance(value, str) else value for value in record.args
                )

        if record.exc_text:
            record.exc_text = _scrub(record.exc_text)

        return True


def configure_logging(level: int = logging.INFO) -> logging.Logger:
    """Install the redacting filter on the root logger and return our logger."""
    handler = logging.StreamHandler(stream=sys.stderr)
    handler.setFormatter(
        logging.Formatter(
            fmt="%(asctime)s %(levelname)-7s %(name)s | %(message)s",
            datefmt="%Y-%m-%dT%H:%M:%S%z",
        )
    )
    handler.addFilter(RedactingFilter())

    root = logging.getLogger()
    root.handlers.clear()
    root.addHandler(handler)
    root.setLevel(level)

    return logging.getLogger("polypaper")
