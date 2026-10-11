# -*- coding: utf-8 -*-
"""生成使用说明文档所需的演示素材：示例图片 + 示例演示文稿。"""
import io
import math
import os
import sys

from PIL import Image, ImageDraw, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "output")
DOCS = os.path.abspath(os.path.join(HERE, "..", "docs"))
os.makedirs(OUT, exist_ok=True)
os.makedirs(os.path.join(DOCS, "screenshots"), exist_ok=True)


# ---------------------------------------------------------------- 示例图片
def make_photo(path, palette, seed=7):
    """画一张简洁的“风景照片”：天空渐变 + 太阳 + 山 + 水面。"""
    w, h = 1280, 800
    sky1, sky2, sun, m1, m2, water = palette
    img = Image.new("RGB", (w, h))
    d = ImageDraw.Draw(img)
    horizon = int(h * 0.62)
    for y in range(horizon):
        t = y / horizon
        c = tuple(int(a + (b - a) * t) for a, b in zip(sky1, sky2))
        d.line([(0, y), (w, y)], fill=c)
    # 太阳
    sx, sy, sr = int(w * 0.68), int(h * 0.30), int(h * 0.09)
    d.ellipse([sx - sr, sy - sr, sx + sr, sy + sr], fill=sun)
    glow = Image.new("RGB", (w, h), (0, 0, 0))
    gd = ImageDraw.Draw(glow)
    gr = int(sr * 2.2)
    gd.ellipse([sx - gr, sy - gr, sx + gr, sy + gr], fill=sun)
    glow = glow.filter(ImageFilter.GaussianBlur(sr * 0.8))
    img = Image.blend(img, Image.composite(glow, img, glow.convert("L").point(lambda v: min(v, 90))), 0.35)
    d = ImageDraw.Draw(img)
    # 远山 / 近山
    def ridge(base, amp, freq, phase, color):
        pts = [(0, h)]
        for x in range(0, w + 1, 8):
            y = base + amp * math.sin(freq * x / w * math.pi * 2 + phase) \
                + amp * 0.5 * math.sin(freq * 2.3 * x / w * math.pi * 2 + phase * 1.7)
            pts.append((x, y))
        pts.append((w, h))
        d.polygon(pts, fill=color)
    ridge(horizon - int(h * 0.10), int(h * 0.055), 2.0, seed, m2)
    ridge(horizon - int(h * 0.04), int(h * 0.045), 3.1, seed + 2, m1)
    # 水面（渐变 + 高光）
    for y in range(horizon, h):
        t = (y - horizon) / (h - horizon)
        c = tuple(int(a + (b - a) * t) for a, b in zip(water, tuple(int(v * 0.55) for v in water)))
        d.line([(0, y), (w, y)], fill=c)
    for k in range(24):
        y = horizon + 6 + k * int((h - horizon) / 26)
        ln = int(w * (0.10 + 0.16 * ((k * 7 + seed) % 5) / 5))
        x0 = sx - ln // 2 + ((k * 37 + seed * 11) % 60) - 30
        a = max(40, 150 - k * 6)
        d.line([(x0, y), (x0 + ln, y)], fill=tuple(min(255, v + a) for v in water), width=2)
    img.save(path, quality=90)


WARM = ((255, 176, 116), (94, 48, 84), (255, 236, 190),
        (66, 34, 60), (44, 24, 46), (255, 148, 102))
COOL = ((168, 216, 255), (240, 248, 255), (255, 250, 224),
        (52, 96, 88), (36, 66, 72), (120, 190, 235))

old_path = os.path.join(OUT, "_demo_old.jpg")
new_path = os.path.join(OUT, "_demo_new.jpg")
make_photo(old_path, WARM, seed=7)
make_photo(new_path, COOL, seed=3)
print("images:", old_path, new_path)


# ---------------------------------------------------------------- 示例演示文稿
def build_deck(pptx_path):
    """用 PowerPoint COM 建一份 5 页的示例文稿（不依赖 python-pptx）。"""
    import win32com.client as com

    const = com.Constants
    app = com.DispatchEx("PowerPoint.Application")
    app.Visible = True
    try:
        pres = app.Presentations.Add(0)  # WithWindow:=msoFalse
        pres.PageSetup.SlideWidth = 960
        pres.PageSetup.SlideHeight = 540

        INK = 0x40332B      # BGR
        ACCENT = 0x2E5AC0   # BGR
        SOFT = 0x80746B
        PAPER = 0xF1F5F7
        WHITE = 0xFFFFFF

        def rect(slide, l, t, w, h, color):
            s = slide.Shapes.AddShape(1, l, t, w, h)  # msoShapeRectangle
            s.Fill.Solid()
            s.Fill.ForeColor.RGB = color
            s.Line.Visible = 0
            s.Shadow.Visible = 0
            return s

        def txt(slide, l, t, w, h, text, size, color=INK, bold=False,
                align=1, font="Microsoft YaHei UI"):
            # align: 1=左 2=中
            box = slide.Shapes.AddTextbox(1, l, t, w, h)
            tf = box.TextFrame
            tf.WordWrap = -1
            tr = tf.TextRange
            tr.Text = text.replace("\n", "\r")
            tr.Font.Size = size
            tr.Font.Bold = bold
            tr.Font.Color.RGB = color
            tr.Font.Name = font
            tr.ParagraphFormat.Alignment = align
            return box

        EMU = 0  # 直接用点
        W, H = 960, 540

        # 1 封面
        s = pres.Slides.Add(1, 12)
        rect(s, 0, 0, W, H, PAPER)
        rect(s, 0, 424, W, 116, ACCENT)
        txt(s, 79, 158, 792, 116, "示例项目 · 方案汇报", 40, INK, True)
        txt(s, 79, 259, 792, 58, "PPT 增强插件套件 — 演示文稿", 18, SOFT)
        txt(s, 83, 450, 576, 43, "2026 年 10 月", 15, WHITE)

        # 2 效果图（放旧图，用于「图片替换」演示）
        s = pres.Slides.Add(2, 12)
        rect(s, 0, 0, W, H, PAPER)
        txt(s, 65, 36, 576, 58, "一、场景效果图", 26, INK, True)
        rect(s, 68, 97, 115, 5, ACCENT)
        s.Shapes.AddPicture(old_path, 0, -1, 65, 130, 547, 342)
        txt(s, 641, 144, 259, 43, "要点", 18, ACCENT, True)
        txt(s, 641, 187, 266, 245,
            "· 场地南侧为主要景观面\n· 建筑体量沿等高线展开\n· 材料以暖色石材为主", 13, SOFT)

        # 3 设计策略
        s = pres.Slides.Add(3, 12)
        rect(s, 0, 0, W, H, PAPER)
        txt(s, 65, 36, 576, 58, "二、设计策略", 26, INK, True)
        rect(s, 68, 97, 115, 5, ACCENT)
        items = [("01", "体量生成", "以功能分区为基础，\n通过庭院缝合体量。"),
                 ("02", "立面语言", "竖向线条统一节奏，\n虚实比例 6:4。"),
                 ("03", "材料选择", "本地石材 + 木色铝板，\n控制造价与质感。")]
        for i, (no, t1, t2) in enumerate(items):
            x = 65 + i * 288
            rect(s, x, 144, 252, 245, WHITE)
            rect(s, x, 144, 252, 9, ACCENT)
            txt(s, x + 22, 169, 202, 65, no, 30, ACCENT, True)
            txt(s, x + 22, 238, 202, 43, t1, 17, INK, True)
            txt(s, x + 22, 285, 209, 101, t2, 11, SOFT)

        # 4 数据
        s = pres.Slides.Add(4, 12)
        rect(s, 0, 0, W, H, PAPER)
        txt(s, 65, 36, 576, 58, "三、主要指标", 26, INK, True)
        rect(s, 68, 97, 115, 5, ACCENT)
        rows = [("总建筑面积", "42,300 ㎡"), ("容积率", "2.4"), ("绿地率", "35%"),
                ("建筑高度", "54 m"), ("停车位", "620 辆"), ("装配率", "60%")]
        for i, (k, v) in enumerate(rows):
            r_, c_ = divmod(i, 3)
            x = 65 + c_ * 288
            y = 144 + r_ * 151
            rect(s, x, y, 252, 115, WHITE)
            txt(s, x + 22, y + 16, 209, 36, k, 12, SOFT)
            txt(s, x + 22, y + 52, 209, 50, v, 21, INK, True)

        # 5 结尾
        s = pres.Slides.Add(5, 12)
        rect(s, 0, 0, W, H, INK)
        txt(s, 79, 216, 792, 86, "谢谢观看", 34, WHITE, True, align=2)
        txt(s, 79, 310, 792, 43, "本文件仅用于「PPT 增强插件套件」功能演示",
            12, 0xC8C0B8, align=2)

        if os.path.exists(pptx_path):
            os.remove(pptx_path)
        pres.SaveAs(os.path.abspath(pptx_path), 24)  # ppSaveAsOpenXMLPresentation
        pres.Close()
        print("deck:", pptx_path, "slides:", pres.Slides.Count if False else 5)
    finally:
        app.Quit()


if __name__ == "__main__":
    build_deck(os.path.join(OUT, "_demo_deck.pptx"))
