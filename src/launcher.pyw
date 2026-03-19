import sys
import os
from pathlib import Path

if getattr(sys, "frozen", False):
    _app_dir = Path(sys.executable).parent.resolve()
    os.chdir(str(_app_dir))
    _src_in_bundle = str(Path(sys._MEIPASS))
    if _src_in_bundle not in sys.path:
        sys.path.insert(0, _src_in_bundle)
else:
    _app_dir = Path(__file__).parent.resolve()
    if str(_app_dir) not in sys.path:
        sys.path.insert(0, str(_app_dir))

from main import main
main()