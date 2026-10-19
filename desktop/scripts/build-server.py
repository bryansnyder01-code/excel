"""Build the backend into one folder with a `sheetcraft-server` program (PyInstaller), at backend/dist/.
Run from anywhere:  python desktop/scripts/build-server.py"""

import os
import sys
from pathlib import Path

import PyInstaller.__main__

backend = Path(__file__).resolve().parents[2] / "backend"
os.chdir(backend)
sys.path.insert(0, str(backend))
PyInstaller.__main__.run([
    str(backend / "desktop_server.py"),
    "--name", "sheetcraft-server",
    "--onedir",
    "--noconfirm",
    "--clean",
    "--distpath", str(backend / "dist"),
    "--workpath", str(backend / "build"),
    "--specpath", str(backend / "build"),
    "--add-data", f"{backend / 'app' / 'templates'}{os.pathsep}app/templates",
    "--collect-submodules", "app",
    "--collect-submodules", "uvicorn",
    "--collect-submodules", "websockets",
    "--collect-data", "anthropic",
    # Not used by the backend; leaving them out keeps the download smaller.
    "--exclude-module", "numpy",
    "--exclude-module", "pandas",
    "--exclude-module", "lxml",
    "--exclude-module", "tkinter",
    "--exclude-module", "pytest",
])
