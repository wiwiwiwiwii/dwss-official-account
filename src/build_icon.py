from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


root = Path(__file__).resolve().parent
scale = 4
side = 256 * scale
image = Image.new("RGBA", (side, side))
draw = ImageDraw.Draw(image)
draw.rounded_rectangle((0, 0, side-1, side-1), radius=side*11/39, fill="#266550")
font = ImageFont.truetype("C:/Windows/Fonts/msyhbd.ttc", round(side*22/39))
box = draw.textbbox((0, 0), "文", font=font)
draw.text(((side-box[2]-box[0])/2, (side-box[3]-box[1])/2), "文", font=font, fill="white")
image = image.resize((256, 256), Image.Resampling.LANCZOS)
image.save(root / "assets/app-icon.png")
image.save(root / "assets/app-icon.ico", sizes=[(n,n) for n in (16,20,24,32,40,48,64,128,256)])
