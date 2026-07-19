"""Generate reproducible Word Order app icons and social card."""

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
WEB = ROOT / "packages" / "web" / "public"
ADDIN = ROOT / "packages" / "addin" / "public"

PAPER = "#f3efe5"
CREAM = "#fffdf7"
INK = "#17231f"
FOREST = "#123d31"
DEEP_FOREST = "#0b2f26"
ORANGE = "#ef5b35"
MINT = "#b8dec8"
LIME = "#d8f08b"

ARIAL = "/System/Library/Fonts/Supplemental/Arial.ttf"
ARIAL_BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
GEORGIA = "/System/Library/Fonts/Supplemental/Georgia.ttf"
GEORGIA_BOLD = "/System/Library/Fonts/Supplemental/Georgia Bold.ttf"


def font(path: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(path, size)


def make_icon() -> Image.Image:
    image = Image.new("RGB", (512, 512), PAPER)
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((0, 0, 511, 511), radius=68, fill=PAPER)
    draw.rounded_rectangle((256, 0, 511, 511), radius=68, fill=FOREST)
    draw.rectangle((256, 0, 443, 511), fill=FOREST)
    draw.rectangle((52, 72, 460, 440), outline=INK, width=8)
    draw.text((151, 281), "W", anchor="mm", font=font(ARIAL_BOLD, 178), fill=INK)
    draw.text((360, 281), "O", anchor="mm", font=font(ARIAL_BOLD, 178), fill=CREAM)
    draw.rectangle((84, 378, 428, 392), fill=ORANGE)
    return image


def document_card(after: bool) -> Image.Image:
    card = Image.new("RGBA", (278, 442), CREAM if after else "#ded8ca")
    draw = ImageDraw.Draw(card)
    label = "AFTER" if after else "BEFORE"
    draw.text((26, 28), label, font=font(ARIAL_BOLD, 13), fill="#287054" if after else "#8b3121")
    if after:
        draw.text((139, 78), "SERVICES AGREEMENT", anchor="mm", font=font(GEORGIA_BOLD, 18), fill=INK)
        headings = [(144, "1.  DEFINITIONS"), (238, "2.  PAYMENT TERMS"), (319, "3.  TERMINATION"), (370, "4.  CONFIDENTIALITY")]
        for y, text in headings:
            draw.text((30, y), text, font=font(GEORGIA_BOLD, 15), fill=INK)
        for y, width in [(174, 216), (190, 216), (206, 170), (268, 216), (284, 186), (348, 210), (399, 205)]:
            draw.line((30, y, 30 + width, y), fill="#b6b0a4", width=3)
        draw.ellipse((219, 389, 253, 423), fill=LIME)
        draw.line((227, 406, 233, 412), fill=FOREST, width=4)
        draw.line((233, 412, 245, 397), fill=FOREST, width=4)
    else:
        draw.text((28, 78), "SERVICES agreement", font=font(GEORGIA_BOLD, 22), fill=INK)
        draw.text((28, 145), "1. DEFINITIONS", font=font(GEORGIA_BOLD, 16), fill=INK)
        draw.text((69, 192), "3. payment terms", font=font(ARIAL, 15), fill=INK)
        draw.rectangle((69, 220, 190, 246), fill="#fff000")
        draw.text((74, 221), "[AMOUNT]", font=font(ARIAL_BOLD, 17), fill=INK)
        draw.text((28, 275), "3. Termination", font=font(GEORGIA_BOLD, 18), fill=INK)
        draw.text((42, 325), "2. CONFIDENTIALITY", font=font(GEORGIA_BOLD, 16), fill=INK)
        draw.line((28, 373, 233, 373), fill="#766f64", width=3)
        draw.line((70, 399, 234, 399), fill="#766f64", width=3)
    return card


def make_social_card() -> Image.Image:
    image = Image.new("RGB", (1200, 630), DEEP_FOREST)
    draw = ImageDraw.Draw(image)
    draw.rectangle((66, 64, 124, 122), fill=PAPER)
    draw.rectangle((95, 64, 124, 122), fill=FOREST)
    draw.text((80, 93), "W", anchor="mm", font=font(ARIAL_BOLD, 22), fill=INK)
    draw.text((109, 93), "O", anchor="mm", font=font(ARIAL_BOLD, 22), fill=CREAM)
    draw.text((142, 79), "Word Order", font=font(ARIAL_BOLD, 28), fill=CREAM)
    draw.text((66, 178), "Put broken Word", font=font(GEORGIA, 67), fill=CREAM)
    draw.text((66, 258), "documents back", font=font(GEORGIA, 67), fill=CREAM)
    draw.text((66, 338), "in order.", font=font(GEORGIA, 67), fill=ORANGE)
    draw.text((68, 464), "Native DOCX repair. Local-only. Open source.", font=font(ARIAL, 24), fill=MINT)
    draw.rectangle((68, 530, 300, 535), fill=ORANGE)

    before = document_card(False).rotate(5, expand=True, resample=Image.Resampling.BICUBIC)
    after = document_card(True).rotate(-4, expand=True, resample=Image.Resampling.BICUBIC)
    image.paste(before, (766, 88), before)
    image.paste(after, (904, 99), after)
    return image


def main() -> None:
    icon = make_icon()
    icon.save(WEB / "icon-512.png", optimize=True)
    icon.resize((192, 192), Image.Resampling.LANCZOS).save(WEB / "icon-192.png", optimize=True)
    icon.resize((80, 80), Image.Resampling.LANCZOS).save(ADDIN / "icon-80.png", optimize=True)
    icon.resize((32, 32), Image.Resampling.LANCZOS).save(ADDIN / "icon-32.png", optimize=True)
    make_social_card().save(WEB / "og.png", optimize=True)


if __name__ == "__main__":
    main()
