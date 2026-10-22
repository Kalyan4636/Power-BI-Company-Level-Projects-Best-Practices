"""Recalculate an .xlsx with headless LibreOffice and return a path to the recalculated copy.

openpyxl writes formulas but no cached values. Opening the file in LibreOffice and re-saving
computes every formula, so tests can read real results with openpyxl(data_only=True).
"""
import os, shutil, subprocess, sys, tempfile


def recalc(src: str, timeout: int = 300) -> str:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice (soffice) not found on PATH")
    out_dir = tempfile.mkdtemp(prefix="recalc_")
    # isolated profile avoids lock/first-run problems in CI
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    cmd = [soffice, f"-env:UserInstallation=file://{profile}", "--headless",
           "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out_dir, src]
    subprocess.run(cmd, check=True, timeout=timeout, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(out_dir, os.path.basename(src))


if __name__ == "__main__":
    print(recalc(sys.argv[1]))
