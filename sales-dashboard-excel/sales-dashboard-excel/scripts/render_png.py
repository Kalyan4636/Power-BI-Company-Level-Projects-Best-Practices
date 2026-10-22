"""Export the Dashboard sheet as a single-page 4K PNG (3840 px wide).

Pipeline: build image-mode workbook -> hide other sheets -> LibreOffice PDF -> poppler 600 dpi
-> crop to content -> Lanczos downscale to 3840 px.

Requires: LibreOffice, poppler-utils (pdftoppm), Pillow, openpyxl.
"""
import argparse, os, shutil, subprocess, sys, tempfile
import openpyxl
from PIL import Image, ImageChops

Image.MAX_IMAGE_PIXELS = None
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "assets", "dashboard_4k.png"))
    ap.add_argument("--width", type=int, default=3840)
    ap.add_argument("--set", nargs="*", default=[], metavar="CELL=VALUE",
                    help="optional filter cells on Dashboard, e.g. J6=Camping D6=Germany")
    a = ap.parse_args()
    tmp = tempfile.mkdtemp()
    xlsx = os.path.join(tmp, "img.xlsx")
    subprocess.run([sys.executable, os.path.join(ROOT, "src", "build_dashboard.py"), "--image-mode", "--out", xlsx], check=True)
    wb = openpyxl.load_workbook(xlsx)
    for kv in a.set:
        k, v = kv.split("=", 1); wb["Dashboard"][k] = v
    for n in wb.sheetnames:
        if n != "Dashboard": wb[n].sheet_state = "hidden"
    wb.save(xlsx)
    profile = tempfile.mkdtemp()
    subprocess.run(["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf", "--outdir", tmp, xlsx],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    subprocess.run(["pdftoppm", "-r", "600", "-png", "-singlefile", os.path.join(tmp, "img.pdf"), os.path.join(tmp, "big")], check=True)
    im = Image.open(os.path.join(tmp, "big.png")).convert("RGB")
    box = ImageChops.difference(im, Image.new("RGB", im.size, (255, 255, 255))).getbbox()
    c = im.crop(box)
    out = c.resize((a.width, round(c.size[1] * a.width / c.size[0])), Image.LANCZOS)
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    out.save(a.out, optimize=True, dpi=(300, 300))
    print(f"Wrote {a.out} {out.size}")
    shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
