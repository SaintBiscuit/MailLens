import sys
from pathlib import Path
import os

sys.path.insert(0, str(Path(__file__).parent.absolute()))

os.system("streamlit run frontend/app.py")