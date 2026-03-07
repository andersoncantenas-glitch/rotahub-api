#!/usr/bin/env python3
from pathlib import Path
import runpy
import sys


ROOT = Path(__file__).resolve().parents[1]


if __name__ == "__main__":
    sys.path.insert(0, str(ROOT))
    runpy.run_path(str(ROOT / "bump_version.py"), run_name="__main__")
