import sys
from pathlib import Path

# Add app directory to path
app_dir = Path(__file__).parent.resolve()
sys.path.insert(0, str(app_dir))

from gui import main
main()