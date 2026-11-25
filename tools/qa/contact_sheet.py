import sys, glob, os
from PIL import Image, ImageDraw
src, prefix, per, cols = sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4])
files = sorted(glob.glob(os.path.join(src, "*.png")))
tw = 560
for k in range(0, len(files), per):
    chunk = files[k:k + per]
    ims = [Image.open(f) for f in chunk]
    th = int(tw * ims[0].height / ims[0].width)
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * (tw + 10) + 10, rows * (th + 30) + 10), "gray")
    d = ImageDraw.Draw(sheet)
    for i, (f, im) in enumerate(zip(chunk, ims)):
        x, y = 10 + (i % cols) * (tw + 10), 10 + (i // cols) * (th + 30)
        sheet.paste(im.convert("RGB").resize((tw, th)), (x, y + 20))
        d.text((x, y + 2), os.path.basename(f), fill="yellow")
    out = f"{prefix}-{k // per + 1}.png"; sheet.save(out); print(out)
