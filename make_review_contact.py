from pathlib import Path
from PIL import Image, ImageDraw

source = Path(r"C:\tmp\bs_pure_review")
groups = {
    "C:/tmp/bs_pure_review_inputs.jpg": ["Control.png", "Setting_Variables.png", "Setting_Labels.png"],
    "C:/tmp/bs_pure_review_outputs.jpg": ["Summary.png", "SandP.png", "Correspondence_S_.png", "Correspondence_P_.png", "QC_Excluded.png"],
}
for output, names in groups.items():
    pieces = []
    for name in names:
        image = Image.open(source / name).convert("RGB")
        image.thumbnail((1100, 900))
        canvas = Image.new("RGB", (1120, image.height + 45), "white")
        ImageDraw.Draw(canvas).text((10, 10), name, fill="black")
        canvas.paste(image, (10, 35))
        pieces.append(canvas)
    contact = Image.new("RGB", (1120, sum(x.height for x in pieces)), "white")
    y = 0
    for piece in pieces:
        contact.paste(piece, (0, y))
        y += piece.height
    contact.save(output, quality=88)
