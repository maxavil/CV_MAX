#!/usr/bin/env python3
"""Thin wrapper so the Phase 0 deliverable runs without installing the package."""

from __future__ import annotations

import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent / "src"))

from polypaper.scripts_phase0 import main

if __name__ == "__main__":
    raise SystemExit(main())
