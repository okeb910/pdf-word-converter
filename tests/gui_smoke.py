"""Short-lived GUI smoke test for macOS GitHub-hosted runners."""

from __future__ import annotations

import os
from pathlib import Path
import sys
import threading


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def main() -> int:
    from tkinterdnd2 import TkinterDnD

    from pdf_word_converter import ConverterApp

    completed = threading.Event()

    def abort_hung_smoke() -> None:
        if not completed.is_set():
            print("GUI smoke test exceeded 8 seconds", file=sys.stderr, flush=True)
            os._exit(124)

    watchdog = threading.Timer(8.0, abort_hung_smoke)
    watchdog.daemon = True
    watchdog.start()

    root = TkinterDnD.Tk()
    root._pdf_converter_dnd_available = True
    root.withdraw()
    app = ConverterApp(root)

    def close_window() -> None:
        completed.set()
        root.destroy()

    root.after(2000, close_window)
    root.mainloop()
    completed.set()
    watchdog.cancel()

    assert app.root is root
    print("GUI smoke test created and destroyed ConverterApp")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
