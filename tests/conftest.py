"""Let the tests import sound_tracker.py from the folder above."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
