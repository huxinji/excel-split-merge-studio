"""Ignore obsolete backport metadata when the standard-library module wins import resolution.

The host Python exposes global site packages to the project environment.  It contains
``pathlib-1.0.1.dist-info`` even though imports correctly resolve to Python's standard
library.  PyInstaller checks distribution metadata only and otherwise refuses to run.
This startup shim is used solely by the packaging command and leaves the host untouched.
"""

from __future__ import annotations

import importlib.metadata

_distribution = importlib.metadata.distribution


def _distribution_without_pathlib_backport(name: str) -> importlib.metadata.Distribution:
    if name.casefold() == "pathlib":
        raise importlib.metadata.PackageNotFoundError(name)
    return _distribution(name)


importlib.metadata.distribution = _distribution_without_pathlib_backport
