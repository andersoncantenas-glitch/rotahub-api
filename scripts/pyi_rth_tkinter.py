import os
import sys
from pathlib import Path


def _set_tk_env() -> None:
    if not getattr(sys, "frozen", False):
        return

    base_dir = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    tcl_root = base_dir / "tcl"
    if not tcl_root.exists():
        return

    if not os.environ.get("TCL_LIBRARY"):
        for candidate in sorted(tcl_root.glob("tcl8.*")):
            if candidate.is_dir():
                os.environ["TCL_LIBRARY"] = str(candidate)
                break

    if not os.environ.get("TK_LIBRARY"):
        for candidate in sorted(tcl_root.glob("tk8.*")):
            if candidate.is_dir():
                os.environ["TK_LIBRARY"] = str(candidate)
                break


_set_tk_env()
