"""PDF export via LibreOffice (headless), which computes every formula before printing."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from app.schemas import WorkbookSpec
from app.services.workbook_builder import build_workbook


class PdfUnavailable(RuntimeError):
    pass


def _soffice() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def pdf_available() -> bool:
    return _soffice() is not None


def workbook_to_pdf(spec: WorkbookSpec, timeout: int = 90) -> bytes:
    exe = _soffice()
    if not exe:
        raise PdfUnavailable("PDF export needs LibreOffice installed on the server.")
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / "workbook.xlsx"
        build_workbook(spec).save(src)
        # A private profile dir keeps parallel conversions from fighting over one LibreOffice profile.
        profile = Path(tmp) / "profile"
        cmd = [exe, f"-env:UserInstallation=file://{profile}", "--headless", "--calc",
               "--convert-to", "pdf", "--outdir", tmp, str(src)]
        try:
            subprocess.run(cmd, check=True, capture_output=True, timeout=timeout)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as e:
            raise PdfUnavailable("The PDF couldn't be created. Please try again.") from e
        out = Path(tmp) / "workbook.pdf"
        if not out.exists():
            raise PdfUnavailable("The PDF couldn't be created. Please try again.")
        return out.read_bytes()
