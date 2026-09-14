"""Hard interlock against real-money trading.

Phase constraint: this system must not be *able* to send a live order, not
merely refrain from doing so. Intent is not a safety mechanism; absence of a
reachable code path is.

The interlock works in two directions:

1.  Import-time poisoning of the authenticated SDK clients. ``SecureClient``
    and ``AsyncSecureClient`` are the only classes in ``polymarket-client``
    that can sign or submit an order. We replace their ``__init__`` so that
    constructing one raises. Nothing downstream can "forget" to check a flag.

2.  Credential refusal. If a signing key is present in the environment we
    abort rather than run, because a key that exists is a key that can leak
    into a log, a traceback, or a journal row.

Lifting this is a deliberate, reviewable act: delete this module's call site
and the tests that assert it. That should never happen by accident.
"""

from __future__ import annotations

import os
from typing import Any, Final, NoReturn

__all__ = ["LiveTradingDisabled", "engage_read_only_interlock"]


class LiveTradingDisabled(RuntimeError):
    """Raised when code attempts to reach a real-money path."""


# Environment variables that indicate a signing-capable credential is present.
# Read-only work needs none of these.
_SIGNING_ENV_VARS: Final[tuple[str, ...]] = (
    "POLYMARKET_PRIVATE_KEY",
    "PRIVATE_KEY",
    "POLYMARKET_API_SECRET",
    "POLYMARKET_PASSPHRASE",
    "WALLET_PRIVATE_KEY",
    "MNEMONIC",
)

_MESSAGE: Final[str] = (
    "Live trading is disabled in this build. The authenticated Polymarket "
    "clients are deliberately unconstructable until the simulator has been "
    "validated and a human has removed this interlock on purpose."
)


def _refuse(*_args: Any, **_kwargs: Any) -> NoReturn:
    raise LiveTradingDisabled(_MESSAGE)


def engage_read_only_interlock() -> None:
    """Make real-money trading unreachable. Call before touching the SDK.

    Idempotent, so it is safe to call from every entry point.
    """
    present = [name for name in _SIGNING_ENV_VARS if os.environ.get(name)]
    if present:
        raise LiveTradingDisabled(
            f"{_MESSAGE} Refusing to start because signing credentials are present "
            f"in the environment: {', '.join(sorted(present))}. Remove them from "
            f"your .env and shell before running paper-trading code."
        )

    import polymarket

    for class_name in ("SecureClient", "AsyncSecureClient"):
        client_class = getattr(polymarket, class_name, None)
        if client_class is None:
            # The SDK dropped or renamed the class. Fail loud: an interlock
            # that silently protects nothing is worse than none at all.
            raise LiveTradingDisabled(
                f"Cannot engage interlock: polymarket.{class_name} not found. "
                f"The SDK surface changed; re-verify before running anything."
            )
        client_class.__init__ = _refuse
