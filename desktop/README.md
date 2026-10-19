# SheetCraft AI desktop app

The whole app in one program for Windows and Mac: no Docker, Python or Node to install.
The backend is bundled inside and serves the screens on this computer (http://127.0.0.1:47615).
On first launch the app asks for an Anthropic API key and keeps it on this computer,
encrypted with the system keychain where available. **File → Change API key** (on a Mac,
**SheetCraft AI → Change API key**) replaces it.

## Getting the installers

Easiest: let GitHub build them. Push the project to GitHub, open **Actions → Desktop app →
Run workflow** (or push a tag such as `v1.0.0`). When the run finishes, download
`SheetCraft-AI-Windows` (the `.exe` installer) and `SheetCraft-AI-macOS` (the `.dmg`) from the
run's **Artifacts**.

The installers aren't code-signed yet, so the first time:

* **Windows:** SmartScreen says "Windows protected your PC". Click **More info → Run anyway**.
* **Mac:** open the `.dmg`, drag SheetCraft AI to Applications, then open it. If macOS says it
  can't be opened, go to **System Settings → Privacy & Security** and click **Open Anyway**.

To remove those warnings, sign the apps: a Windows code-signing certificate, and an Apple
Developer ID with notarization. electron-builder picks them up from `CSC_LINK` /
`CSC_KEY_PASSWORD` (and `APPLE_ID`, `APPLE_APP_SPECIFIC_PASSWORD`, `APPLE_TEAM_ID` for
notarization) set as GitHub secrets; then drop `CSC_IDENTITY_AUTO_DISCOVERY` from
`.github/workflows/desktop.yml` and `"identity": "-"` from `package.json`.

The Mac build is for Apple Silicon (M1 and later). Intel Macs need a build on an Intel runner.

## Building it yourself

Each installer has to be built on its own system (Windows for `.exe`, a Mac for `.dmg`).
Needs Python 3.11+ and Node 20+.

```bash
pip install -r backend/requirements.txt pyinstaller pillow
python desktop/scripts/build-server.py      # backend → backend/dist/sheetcraft-server
cd frontend && npm install && npm run build:desktop && cd ..   # screens → frontend/out
cd desktop && npm install
npm start                                    # try it (uses Python and frontend/out directly)
npm run dist                                 # installer for this system → desktop/dist
```

## What works, and what doesn't

Everything that runs on one computer works: building and editing with AI, uploads, photo and
PDF import, templates, charts, dashboards, downloads (Excel, CSV; PDF when LibreOffice is
installed), version history, live data connections, and your own sign-in for saving workbooks.

Things that need a shared server don't apply here: live co-editing with other people, invite
and share links, comments with @mentions to others, and emailed reports. For those, host the
web version (see the main README).

Workbooks and the account database live in the app's data folder (**File → Open data folder**):
`%APPDATA%\SheetCraft AI` on Windows, `~/Library/Application Support/SheetCraft AI` on a Mac.
If the app won't start, `server.log` there says why.
