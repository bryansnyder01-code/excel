"""The backend as one program for the desktop app: serves the API and the built screens on one port.

    python desktop_server.py --port 8765 --web ../frontend/out --data ~/.sheetcraft

The desktop app passes the Anthropic API key in the ANTHROPIC_API_KEY environment variable.
PyInstaller builds this file into the `sheetcraft-server` program (see desktop/README.md).
"""

import argparse
import os
import sys
import threading


def exit_with_parent() -> None:
    """The desktop app keeps our stdin open; when it closes (the app quit or crashed), stop too,
    so no engine is left running in the background holding the port."""
    def watch():
        try:
            while sys.stdin.buffer.read(1024):
                pass
        except (OSError, ValueError):
            pass
        os._exit(0)

    threading.Thread(target=watch, daemon=True).start()


def main() -> None:
    parser = argparse.ArgumentParser(description="SheetCraft AI desktop server")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--web", required=True, help="folder with the frontend's static export")
    parser.add_argument("--data", required=True, help="where accounts and saved workbooks live")
    parser.add_argument("--exit-with-parent", action="store_true", help="stop when stdin closes (set by the desktop app)")
    args = parser.parse_args()
    if args.exit_with_parent:
        exit_with_parent()

    origin = f"http://127.0.0.1:{args.port}"
    # Settings are read when app.main is imported, so set them first.
    os.environ.update({
        "SHEETCRAFT_STATIC_DIR": os.path.abspath(args.web),
        "SHEETCRAFT_DATA_DIR": os.path.abspath(args.data),
        "SHEETCRAFT_APP_URL": origin,
        "SHEETCRAFT_CORS_ORIGINS": f'["{origin}"]',
        # It's the user's own key on their own computer, so no daily allowance.
        "SHEETCRAFT_AI_LIMIT_ENABLED": "false",
    })
    os.makedirs(args.data, exist_ok=True)

    import uvicorn
    from app.main import app

    print(f"SheetCraft AI running at {origin}", flush=True)
    uvicorn.run(app, host="127.0.0.1", port=args.port, log_level="warning")


if __name__ == "__main__":
    sys.exit(main())
