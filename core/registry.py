# core/registry.py
"""
Lightweight plugin registry for processors and checks.

Usage pattern:
- Each concrete processor/check module creates an instance and calls
  register_processor(...) / register_check(...) at import time.
- The orchestrator asks this registry for all processors and checks,
  and then wires them together based on file extensions.

Why this design (SOLID):
- OCP: New processors/checks register themselves; no central switch/case grows.
- DIP: Orchestrator depends on abstractions (interfaces) and this registry, not concretes.
- SRP: This module only manages registration & retrieval. Nothing else.
"""

from __future__ import annotations

from typing import List, Type

from .interfaces import FileProcessor, Check

# Internal stores (module-private): lists of SINGLETON-like instances.
# We keep instances stateless; multiple instances of the same class are unnecessary.
_PROCESSORS: List[FileProcessor] = []
_CHECKS: List[Check] = []


def register_processor(p: FileProcessor) -> None:
    """
    Register a processor instance if not already present.
    We consider processors unique by their concrete class to avoid duplicates
    when modules are imported more than once in some environments.
    """
    if not any(isinstance(existing, type(p)) for existing in _PROCESSORS):
        _PROCESSORS.append(p)


def register_check(c: Check) -> None:
    """
    Register a check instance if not already present.
    Checks are unique by concrete class (one instance per check class).
    """
    if not any(isinstance(existing, type(c)) for existing in _CHECKS):
        _CHECKS.append(c)


def processors() -> List[FileProcessor]:
    """
    Return a shallow copy of registered processors to prevent accidental mutation
    of the internal list by callers.
    """
    return list(_PROCESSORS)


def checks() -> List[Check]:
    """
    Return a shallow copy of registered checks.
    """
    return list(_CHECKS)


def clear_registry() -> None:
    """
    Testing helper: wipe current registrations.
    Not intended for use in the running app.
    """
    _PROCESSORS.clear()
    _CHECKS.clear()
