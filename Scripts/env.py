"""Import shim for 00_EnvironmentVariables (leading digit prevents normal import)."""

import importlib.util
from pathlib import Path

_ENV_PATH = Path(__file__).resolve().with_name("00_EnvironmentVariables.py")
_spec = importlib.util.spec_from_file_location("rm_environment", _ENV_PATH)
if _spec is None or _spec.loader is None:
    raise ImportError(f"Cannot load environment module from {_ENV_PATH}")

_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)

for _name, _value in vars(_mod).items():
    if not _name.startswith("_"):
        globals()[_name] = _value

__all__ = sorted(name for name in globals() if not name.startswith("_"))

del _ENV_PATH, _name, _spec, _value, importlib, Path
