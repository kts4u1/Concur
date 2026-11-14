# -*- coding: utf-8 -*-
"""프로그램 아이콘(concur_auto.ico) 만들기: 파란 둥근 네모 + 흰 영수증 + 초록 체크.
다시 만들 때: python tools/make_icon.py  (Pillow 필요)"""

from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter

S = 1024                                   # 크게 그린 뒤 줄입니다
SIZES = [16, 20, 24, 32, 40, 48, 64, 96, 128, 256]
OUT = Path(__file__).resolve().parent.parent / "concur_auto.ico"


def gradient(size, top, bottom):
    g = Image.new("RGB", (1, 256))
    for y in range(256):
        t = y / 255
        g.putpixel((0, y), tuple(round(a + (b - a) * t) for a, b in zip(top, bottom)))
    return g.resize(size, Image.BILINEAR)


def rounded_mask(size, box, radius):
    m = Image.new("L", size, 0)
    ImageDraw.Draw(m).rounded_rectangle(box, radius=radius, fill=255)
    return m


def receipt_shape(x0, y0, x1, y1, teeth=6, depth=46):
    """위는 반듯하고 아래는 톱니 모양인 영수증 테두리 점들"""
    pts = [(x0, y0), (x1, y0), (x1, y1)]
    w = (x1 - x0) / teeth
    for i in range(teeth):
        pts += [(x1 - w * (i + 0.5), y1 - depth), (x1 - w * (i + 1), y1)]
    return pts


def draw(simple=False):
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    pad, radius = 40, 210
    # 바탕: 위는 밝은 SAP 파랑, 아래는 짙은 남색
    bg = gradient((S, S), (28, 132, 228), (12, 52, 104)).convert("RGBA")
    img.paste(bg, (0, 0), rounded_mask((S, S), (pad, pad, S - pad, S - pad), radius))
    # 위쪽 은은한 빛
    shine = Image.new("RGBA", (S, S), (255, 255, 255, 0))
    ImageDraw.Draw(shine).ellipse((-200, -620, S + 200, 420), fill=(255, 255, 255, 26))
    shine = shine.filter(ImageFilter.GaussianBlur(40))
    shine.putalpha(ImageChops.multiply(shine.getchannel("A"), rounded_mask((S, S), (pad, pad, S - pad, S - pad), radius)))
    img = Image.alpha_composite(img, shine)

    # 영수증 그림자 + 영수증
    x0, y0, x1, y1 = 262, 170, 700, 846
    pts = receipt_shape(x0, y0, x1, y1)
    shadow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).polygon([(x + 18, y + 26) for x, y in pts], fill=(0, 20, 50, 120))
    img = Image.alpha_composite(img, shadow.filter(ImageFilter.GaussianBlur(22)))
    d = ImageDraw.Draw(img)
    d.polygon(pts, fill=(255, 255, 255, 255))

    # 영수증 글줄
    blue, gray = (10, 110, 209, 255), (196, 206, 218, 255)
    lx0 = x0 + 62
    d.rounded_rectangle((lx0, y0 + 70, lx0 + 210, y0 + 118), radius=24, fill=blue)       # 가게 이름
    if not simple:
        for i, w in enumerate((300, 250, 280)):
            y = y0 + 190 + i * 74
            d.rounded_rectangle((lx0, y, lx0 + w, y + 30), radius=15, fill=gray)
        y = y0 + 430
        for x in range(lx0, x1 - 60, 44):                                             # 점선
            d.rounded_rectangle((x, y, x + 24, y + 10), radius=5, fill=gray)
        d.rounded_rectangle((lx0, y0 + 480, lx0 + 150, y0 + 524), radius=22, fill=(120, 136, 156, 255))  # 합계
    else:
        for i in range(2):
            y = y0 + 210 + i * 110
            d.rounded_rectangle((lx0, y, lx0 + 280, y + 52), radius=26, fill=gray)

    # 초록 체크 배지(흰 테두리)
    cx, cy, r = 690, 714, 196
    ring = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(ring).ellipse((cx - r - 10, cy - r + 6, cx + r + 22, cy + r + 34), fill=(0, 30, 60, 110))
    img = Image.alpha_composite(img, ring.filter(ImageFilter.GaussianBlur(16)))
    d = ImageDraw.Draw(img)
    d.ellipse((cx - r - 26, cy - r - 26, cx + r + 26, cy + r + 26), fill=(255, 255, 255, 255))
    badge = gradient((2 * r, 2 * r), (52, 199, 120), (22, 150, 86)).convert("RGBA")
    m = Image.new("L", (2 * r, 2 * r), 0)
    ImageDraw.Draw(m).ellipse((0, 0, 2 * r - 1, 2 * r - 1), fill=255)
    img.paste(badge, (cx - r, cy - r), m)
    d = ImageDraw.Draw(img)
    check = [(cx - 98, cy + 4), (cx - 26, cy + 76), (cx + 106, cy - 70)]
    d.line(check, fill=(255, 255, 255, 255), width=64 if not simple else 78, joint="curve")
    for x, y in (check[0], check[-1]):
        w = (64 if not simple else 78) // 2
        d.ellipse((x - w, y - w, x + w, y + w), fill=(255, 255, 255, 255))
    return img


def main():
    big, small = draw(), draw(simple=True)
    images = [(small if n <= 32 else big).resize((n, n), Image.LANCZOS) for n in SIZES]
    images[-1].save(OUT, format="ICO", sizes=[(n, n) for n in SIZES], append_images=images[:-1])
    print(f"만듦: {OUT}")


if __name__ == "__main__":
    main()
