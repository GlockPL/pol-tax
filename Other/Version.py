"""
The app version lives only in pyproject.toml. Builds bake it into Other/_version.py (generated, not committed),
when running from source it is read straight from pyproject.toml.
"""
import tomllib
from pathlib import Path


def app_version() -> str:
    try:
        from Other._version import VERSION
        return VERSION
    except ImportError:
        pass

    pyproject = Path(__file__).resolve().parent.parent / "pyproject.toml"
    try:
        with pyproject.open("rb") as f:
            return tomllib.load(f)["project"]["version"]
    except (OSError, KeyError, tomllib.TOMLDecodeError):
        return "dev"


def windows_version_tuple(version: str, build: int = 0) -> tuple:
    """
    Windows file versions need exactly four numbers: major.minor.patch.build
    """
    parts = [int(part) for part in version.split(".")[:3]]
    parts += [0] * (3 - len(parts))
    return tuple(parts) + (build,)
