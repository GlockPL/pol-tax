import os
import sys
import tempfile
from pathlib import Path

# Keep app data (settings, cache, Excel output) out of the user's Documents folder
os.environ.setdefault("POL_TAX_HOME", tempfile.mkdtemp(prefix="pol-tax-tests-"))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
