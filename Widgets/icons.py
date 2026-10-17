"""
Icons are loaded straight from the resources folder through the Qt "icons:" search path,
e.g. QPixmap("icons:excel_icon.jpg").
"""
import sys
from pathlib import Path

from PySide6.QtCore import QDir


def resources_dir() -> Path:
    # PyInstaller unpacks bundled data next to the modules (pol-tax.spec puts it in "resources")
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "resources"
    return Path(__file__).resolve().parent.parent / "resources"


QDir.addSearchPath("icons", str(resources_dir()))
