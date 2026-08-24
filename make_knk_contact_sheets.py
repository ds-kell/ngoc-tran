from pathlib import Path
from PIL import Image, ImageOps, ImageDraw

root = Path(r"D:\NgocTran\phase1\_work_knk_2022\render")
for folder in sorted(p for p in root.iterdir() if p.is_dir()):
    pages = sorted(folder.glob("page-*.png"))
    for start in range(0, len(pages), 8):
        chunk = pages[start:start+8]
        thumbs = []
        for p in chunk:
            im = Image.open(p).convert("RGB")
            im.thumbnail((360, 510))
            canvas = Image.new("RGB", (380, 550), "white")
            canvas.paste(im, ((380-im.width)//2, 20))
            ImageDraw.Draw(canvas).text((12, 528), p.stem, fill="black")
            thumbs.append(canvas)
        sheet = Image.new("RGB", (380*4, 550*2), (210,210,210))
        for i, im in enumerate(thumbs):
            sheet.paste(im, ((i%4)*380, (i//4)*550))
        sheet.save(root / f"{folder.name}_pages_{start+1:02d}-{start+len(chunk):02d}.jpg", quality=90)
