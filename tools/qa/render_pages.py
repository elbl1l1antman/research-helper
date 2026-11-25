"""Render HWPX pages to PNG through Hancom COM (local Windows only).

Opens a COPY of the document so a file held open elsewhere is never touched.
Usage: render_pages.py <doc.hwpx> <out_dir> [--pages 1,3,14] [--dpi 60]
"""
import argparse
import os
import shutil
import tempfile

import win32com.client
from PIL import Image

parser = argparse.ArgumentParser()
parser.add_argument("doc")
parser.add_argument("out_dir")
parser.add_argument("--pages", help="1-based, comma separated (default: all)")
parser.add_argument("--dpi", type=int, default=60)
args = parser.parse_args()
os.makedirs(args.out_dir, exist_ok=True)
copy = os.path.join(tempfile.mkdtemp(prefix="qa_render_"), "copy.hwpx")
shutil.copy2(args.doc, copy)

hwp = win32com.client.DispatchEx("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
try:
    if not hwp.Open(copy, "HWPX", "forceopen:true"):
        raise SystemExit("open failed")
    count = hwp.PageCount
    pages = [int(p) for p in args.pages.split(",")] if args.pages else range(1, count + 1)
    for page in pages:
        if page > count:
            continue
        bmp = os.path.join(args.out_dir, f"p{page:02d}.bmp")
        hwp.CreatePageImage(bmp, page - 1, args.dpi, 24, "bmp")
        Image.open(bmp).save(bmp[:-4] + ".png")
        os.remove(bmp)
    print("pages", count)
finally:
    hwp.Clear(1)
    hwp.Quit()
