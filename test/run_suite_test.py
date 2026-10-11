# -*- coding: utf-8 -*-
"""
套件回归测试：加载自检版 .ppam，真实跑一遍三个插件并校验结果。

与旧项目的 e2e 测试不同，这里是**通过加载项本身**运行（而非向演示文稿注入模块），
因此同时验证了「套件能被 PowerPoint 加载 + 其 VBA 可执行」这条链路。

覆盖：
  插件① 导出图片  —— 真实导出并逐字节解析 PNG / JPG 头，核对输出像素
  插件② 选页另存  —— 无窗口打开 + 按序号保留页面
  插件③ 图片替换  —— 替换前后逐项属性断言 + OOXML 解包独立核对
"""

import os
import re
import sys
import zlib
import struct
import zipfile
import xml.etree.ElementTree as ET

import win32com.client as com

HERE = os.path.dirname(os.path.abspath(__file__))
SUITE = os.path.dirname(HERE)
ADDIN = os.path.join(SUITE, "dist", "PptEnhanceSuiteTest.ppam")
RESULT = os.path.join(HERE, "vbresult.txt")
OUTDIR = os.path.join(HERE, "output")
DECK5 = os.path.join(OUTDIR, "_deck5.pptx")
SELOUT = os.path.join(OUTDIR, "_selected.pptx")
IMG_OLD = os.path.join(OUTDIR, "_replace_old.png")
IMG_NEW = os.path.join(OUTDIR, "_replace_new.png")
REPLACE_IN = os.path.join(OUTDIR, "_replace_in.pptx")
REPLACE_OUT = os.path.join(OUTDIR, "_replace_out.pptx")
GROUP_IN = os.path.join(OUTDIR, "_group_in.pptx")
GROUP_OUT = os.path.join(OUTDIR, "_group_out.pptx")

PROJ = "PptEnhanceSuiteTest"


# ---------------------------------------------------------------- 图片尺寸

def png_size(path):
    with open(path, "rb") as f:
        d = f.read(26)
    return struct.unpack(">II", d[16:24])


def jpg_size(path):
    with open(path, "rb") as f:
        d = f.read()
    i = 2
    while i < len(d) - 9:
        if d[i] != 0xFF:
            i += 1
            continue
        m = d[i + 1]
        if m in (0xC0, 0xC1, 0xC2, 0xC3):
            h = struct.unpack(">H", d[i + 5:i + 7])[0]
            w = struct.unpack(">H", d[i + 7:i + 9])[0]
            return (w, h)
        if m in (0xD8, 0xD9) or 0xD0 <= m <= 0xD7:
            i += 2
            continue
        ln = struct.unpack(">H", d[i + 2:i + 4])[0]
        i += 2 + ln
    return None


def write_png(path, w, h, rgb):
    """手写最小 PNG（不依赖 Pillow）"""
    raw = bytearray()
    for _ in range(h):
        raw.append(0)
        raw.extend(bytes(rgb) * w)

    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data +
                struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(bytes(raw), 9))
    png += chunk(b"IEND", b"")
    with open(path, "wb") as f:
        f.write(png)


# ---------------------------------------------------------------- 造样张

def build_deck(app, n=3):
    pres = app.Presentations.Add()
    while pres.Slides.Count > 0:
        pres.Slides(1).Delete()
    pres.PageSetup.SlideWidth = 13.333 * 72
    pres.PageSetup.SlideHeight = 7.5 * 72
    colors = [0xC0504D, 0x9BBB59, 0x4F81BD]
    for i in range(n):
        s = pres.Slides.Add(i + 1, 12)          # ppLayoutBlank
        shp = s.Shapes.AddShape(1, 60, 60, 400, 240)   # msoShapeRectangle
        shp.Fill.ForeColor.RGB = colors[i % len(colors)]
        tb = s.Shapes.AddTextbox(1, 60, 330, 500, 60)
        tb.TextFrame.TextRange.Text = "Slide %d" % (i + 1)
    return pres


def build_replace_deck(app, deck_path, img_old):
    """构造「图片替换」样张。

    目标图片带全套属性，并且**被另一张图压在上面** ——
    只要 Z 序没有还原，它就会跑到最顶层，测试立刻暴露。
    """
    pres = app.Presentations.Add()
    while pres.Slides.Count > 0:
        pres.Slides(1).Delete()
    pres.PageSetup.SlideWidth = 13.333 * 72
    pres.PageSetup.SlideHeight = 7.5 * 72
    s = pres.Slides.Add(1, 12)                  # ppLayoutBlank

    bg = s.Shapes.AddShape(1, 30, 30, 800, 420)         # 底层衬板
    bg.Fill.ForeColor.RGB = 0xD8D8D8
    bg.Name = "BG_PANEL"

    tgt = s.Shapes.AddPicture(img_old, 0, -1, 100, 80, 320, 220)
    tgt.Name = "TGT_IMAGE"
    tgt.AlternativeText = "ALT-ORIG"
    tgt.Title = "TITLE-ORIG"
    tgt.PictureFormat.CropLeft = 12
    tgt.PictureFormat.CropTop = 6
    tgt.PictureFormat.CropRight = 9
    tgt.PictureFormat.CropBottom = 4
    tgt.PictureFormat.Brightness = 0.62
    tgt.PictureFormat.Contrast = 0.55
    tgt.Fill.Transparency = 0.35
    tgt.Line.Visible = -1
    tgt.Line.ForeColor.RGB = 0x0000FF
    tgt.Line.Weight = 3
    tgt.Rotation = 12.5
    tgt.Flip(0)                                 # msoFlipHorizontal（pywin32 下
    #  FlipHorizontal 属性被动态绑定视为只读，改用 Flip 方法达到同样效果）
    tgt.Shadow.Visible = -1
    tgt.ActionSettings(1).Hyperlink.Address = "https://example.com/orig"

    top = s.Shapes.AddPicture(img_old, 0, -1, 380, 180, 180, 140)
    top.Name = "TOP_OVERLAY"

    plain = s.Shapes.AddPicture(img_old, 0, -1, 600, 60, 120, 90)
    plain.Name = "PLAIN_IMG"

    pres.SaveAs(deck_path, 24)                  # ppSaveAsOpenXMLPresentation
    safe_close(pres, app)


def build_group_deck(app, deck_path, img_old):
    """构造「图片位于编组内」的样张。

    组合 GRP_TARGET = 目标图片 TGT_IMAGE + 装饰矩形 DECO_RECT，并且刻意带上
    三处「处理不当就会暴露」的特征：

      · 组合自身带**旋转和水平翻转** —— 解组时 PowerPoint 会把组合的变换
        「烘焙」进每个子形状，如果不先把变换归零，重组后组合的旋转就回不来；
      · 组合**下面压着 BG_PANEL、上面压着 TOP_OVERLAY** —— 组合的 Z 序
        如果没还原，层级关系立刻出错；
      · 组内图片自身还带裁剪 / 旋转 / 亮度 / 对比度 / 透明度 / 边框 / 超链接。
    """
    pres = app.Presentations.Add()
    while pres.Slides.Count > 0:
        pres.Slides(1).Delete()
    pres.PageSetup.SlideWidth = 13.333 * 72
    pres.PageSetup.SlideHeight = 7.5 * 72
    s = pres.Slides.Add(1, 12)                  # ppLayoutBlank

    bg = s.Shapes.AddShape(1, 30, 30, 800, 420)         # 组合下方的衬板
    bg.Fill.ForeColor.RGB = 0xD8D8D8
    bg.Name = "BG_PANEL"

    tgt = s.Shapes.AddPicture(img_old, 0, -1, 100, 80, 320, 220)
    tgt.Name = "TGT_IMAGE"
    tgt.AlternativeText = "ALT-GRP"
    tgt.Title = "TITLE-GRP"
    tgt.PictureFormat.CropLeft = 12
    tgt.PictureFormat.CropTop = 6
    tgt.PictureFormat.CropRight = 9
    tgt.PictureFormat.CropBottom = 4
    tgt.PictureFormat.Brightness = 0.62
    tgt.PictureFormat.Contrast = 0.55
    tgt.Fill.Transparency = 0.35
    tgt.Line.Visible = -1
    tgt.Line.ForeColor.RGB = 0x0000FF
    tgt.Line.Weight = 3
    tgt.Rotation = 8.0
    tgt.ActionSettings(1).Hyperlink.Address = "https://example.com/grp"

    deco = s.Shapes.AddShape(1, 460, 140, 150, 120)
    deco.Fill.ForeColor.RGB = 0x4F81BD
    deco.Name = "DECO_RECT"

    grp = s.Shapes.Range(["TGT_IMAGE", "DECO_RECT"]).Group()
    grp.Name = "GRP_TARGET"
    grp.Flip(0)                     # msoFlipHorizontal：必须先翻转
    grp.Rotation = 10.0             # 后旋转（Flip 会把 Rotation 取反）

    top = s.Shapes.AddPicture(img_old, 0, -1, 900, 300, 120, 90)
    top.Name = "TOP_OVERLAY"

    pres.SaveAs(deck_path, 24)
    safe_close(pres, app)


# ---------------------------------------------------------------- OOXML 核对

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _q(tag, ns):
    return "{%s}%s" % (ns, tag)


def _pic_fields(el):
    """从 p:pic 元素提取属性字段（pic_props 与 group_props 共用）。"""
    cNvPr = el.find(".//%s" % _q("cNvPr", P_NS))
    if cNvPr is None:
        return None
    xfrm = el.find(".//%s/%s" % (_q("spPr", P_NS), _q("xfrm", A_NS)))
    off = xfrm.find(_q("off", A_NS)) if xfrm is not None else None
    ext = xfrm.find(_q("ext", A_NS)) if xfrm is not None else None
    blipFill = el.find(".//%s" % _q("blipFill", P_NS))
    blip = blipFill.find(_q("blip", A_NS)) if blipFill is not None else None
    srcRect = blipFill.find(_q("srcRect", A_NS)) if blipFill is not None else None
    lum = blip.find(_q("lum", A_NS)) if blip is not None else None
    alpha = blip.find(_q("alphaModFix", A_NS)) if blip is not None else None
    ln = el.find(".//%s/%s" % (_q("spPr", P_NS), _q("ln", A_NS)))
    eff = el.find(".//%s/%s" % (_q("spPr", P_NS), _q("effectLst", A_NS)))
    hl = cNvPr.find(_q("hlinkClick", A_NS))
    clr = ln.find(".//%s" % _q("srgbClr", A_NS)) if ln is not None else None

    return {
        "name": cNvPr.get("name"),
        "descr": cNvPr.get("descr"),
        "title": cNvPr.get("title"),
        "link": (hl.get(_q("id", R_NS)) if hl is not None else None),
        "rot": (xfrm.get("rot") if xfrm is not None else None),
        "flipH": (xfrm.get("flipH") if xfrm is not None else None),
        "flipV": (xfrm.get("flipV") if xfrm is not None else None),
        "off.x": (off.get("x") if off is not None else None),
        "off.y": (off.get("y") if off is not None else None),
        "ext.cx": (ext.get("cx") if ext is not None else None),
        "ext.cy": (ext.get("cy") if ext is not None else None),
        "crop.l": (srcRect.get("l") if srcRect is not None else None),
        "crop.t": (srcRect.get("t") if srcRect is not None else None),
        "crop.r": (srcRect.get("r") if srcRect is not None else None),
        "crop.b": (srcRect.get("b") if srcRect is not None else None),
        "lum.bright": (lum.get("bright") if lum is not None else None),
        "lum.contrast": (lum.get("contrast") if lum is not None else None),
        "alpha": (alpha.get("amt") if alpha is not None else None),
        "line.w": (ln.get("w") if ln is not None else None),
        "line.clr": (clr.get("val") if clr is not None else None),
        "has_effect": (eff is not None),
        "embed": (blip.get(_q("embed", R_NS)) if blip is not None else None),
    }


def _slide_parts(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
        media = [n for n in z.namelist() if n.startswith("ppt/media/")]
    root = ET.fromstring(xml)
    return root.find(".//%s/%s" % (_q("cSld", P_NS), _q("spTree", P_NS))), media


def pic_props(path, name="TGT_IMAGE"):
    """从 pptx 的 slide1.xml 里提取名为 name 的图片节点的全部相关属性"""
    spTree, media = _slide_parts(path)
    kids = list(spTree)
    for i, el in enumerate(kids):
        if el.tag != _q("pic", P_NS):
            continue
        cNvPr = el.find(".//%s" % _q("cNvPr", P_NS))
        if cNvPr is None or cNvPr.get("name") != name:
            continue
        d = _pic_fields(el)
        if d is None:
            return None
        d["index"] = i
        d["media_count"] = len(media)
        return d
    return None


# p:grpSp 下直接子成员的标签名（顺序即 Z 序）
MEMBER_TAGS = ("pic", "sp", "grpSp", "graphicFrame", "cxnSp")


def group_props(path, gname="GRP_TARGET", pic_name="TGT_IMAGE"):
    """提取组合及其内部目标图片的属性。

    这是「图片位于编组内」场景的独立证据来源 —— 完全不经过 VBA，
    直接从 OOXML 看组合的 xfrm / 成员构成 / 组内图片属性有没有被改动。
    """
    spTree, media = _slide_parts(path)

    def find_group(node):
        for el in list(node):
            if el.tag != _q("grpSp", P_NS):
                continue
            nv = el.find("%s/%s" % (_q("nvGrpSpPr", P_NS), _q("cNvPr", P_NS)))
            if nv is not None and nv.get("name") == gname:
                return el
            deeper = find_group(el)
            if deeper is not None:
                return deeper
        return None

    g = find_group(spTree)
    if g is None:
        return None

    nv = g.find("%s/%s" % (_q("nvGrpSpPr", P_NS), _q("cNvPr", P_NS)))
    xfrm = g.find("%s/%s" % (_q("grpSpPr", P_NS), _q("xfrm", A_NS)))
    off = xfrm.find(_q("off", A_NS)) if xfrm is not None else None
    ext = xfrm.find(_q("ext", A_NS)) if xfrm is not None else None

    idx = -1
    for i, el in enumerate(list(spTree)):
        if el is g:
            idx = i
            break

    members = []
    for el in list(g):
        if el.tag.split("}")[-1] in MEMBER_TAGS:
            nn = el.find(".//%s" % _q("cNvPr", P_NS))
            members.append("%s:%s" % (el.tag.split("}")[-1],
                                      nn.get("name") if nn is not None else "?"))

    out = {
        "grp.name": nv.get("name") if nv is not None else None,
        "grp.index": idx,
        "grp.off.x": off.get("x") if off is not None else None,
        "grp.off.y": off.get("y") if off is not None else None,
        "grp.ext.cx": ext.get("cx") if ext is not None else None,
        "grp.ext.cy": ext.get("cy") if ext is not None else None,
        "grp.rot": xfrm.get("rot") if xfrm is not None else None,
        "grp.flipH": xfrm.get("flipH") if xfrm is not None else None,
        "grp.flipV": xfrm.get("flipV") if xfrm is not None else None,
        "grp.members": "|".join(members),
        "grp.member_count": len(members),
        "media_count": len(media),
    }

    for pel in g.findall(".//%s" % _q("pic", P_NS)):
        nn = pel.find(".//%s" % _q("cNvPr", P_NS))
        if nn is not None and nn.get("name") == pic_name:
            sub = _pic_fields(pel)
            if sub:
                for k, v in sub.items():
                    out["pic." + k] = v
            break
    return out


# 允许 ±NUM_TOL 浮点抖动的字段白名单。
#
# 为什么 off / ext 需要容差（已实测确认，不是猜测）：
#     测试目标形状带旋转。对旋转形状，Shape.Left / Top 读写的是**外接框**坐标，
#     PowerPoint 内部要在「外接框」与「未旋转的 off/ext」之间换算，而 OOXML 里
#     存的是整数 EMU。实际表现是：读出 shp.Top 再原值写回，落盘仍会与源相差
#     1 EMU —— 也就是说这个往返**本来就无法在 VBA 层面做到 bit 级一致**。
#     量级参考：914400 EMU = 1 英寸，故 1 EMU ≈ 0.00003 mm，纯舍入噪声。
#     而历史上真实出现过的缺陷量级是「off 差几十万、ext.cx 差几百」，
#     相差 1 与它们绝非同类问题。
#
# 为什么必须是「白名单」而不是「所有数值字段都给容差」：
#     最初一版把容差无差别套到所有数值上，结果 media_count 1 -> 2 被
#     「abs(1-2) <= 2」判成相等 —— 「替换到底有没有真的换图」这项最关键的
#     证据被静默吞掉。凡是**小整数**字段（media_count / index）与**比例值**
#     字段（crop.* / lum.*），差 1 就是真问题，必须严格比较。
TOL_KEYS = frozenset({"off.x", "off.y", "ext.cx", "ext.cy", "line.w"})
NUM_TOL = 2


def _as_int(v):
    if v is None:
        return None
    try:
        return int(str(v))
    except (TypeError, ValueError):
        return None


def _tol_key(key):
    """容差白名单匹配：忽略 grp. / pic. 这类前缀。

    组合场景的字段名带 grp./pic. 前缀（如 grp.off.x、pic.ext.cx），
    而白名单是按几何字段名写的，所以逐层剥前缀后再匹配。
    """
    k = key
    while k not in TOL_KEYS and "." in k:
        k = k.split(".", 1)[1]
    return k in TOL_KEYS


def val_eq(key, a, b):
    """属性相等判定：仅白名单内的几何字段允许 ±NUM_TOL 的浮点抖动。"""
    if a == b:
        return True
    if not _tol_key(key):
        return False
    na, nb = _as_int(a), _as_int(b)
    if na is None or nb is None:
        return False
    return abs(na - nb) <= NUM_TOL


# 这些字段本来就**应当**变化，不算差异（图片引用必然换新、媒体数 +1）。
# 按「后缀」匹配，这样组合场景的 pic.embed 也能命中。
IGNORE_SUFFIX = ("embed", "media_count")


def _is_ignored(key):
    return key in IGNORE_SUFFIX or key.split(".")[-1] in IGNORE_SUFFIX


def compare_xml(before, after):
    """逐项比对，返回 (是否全等, 差异说明列表, 变化的引用列表)"""
    diffs, changed = [], []
    for k in before:
        if _is_ignored(k):
            if not val_eq(k, before[k], after[k]):
                changed.append("%s: %s -> %s" % (k, before[k], after[k]))
            continue
        if not val_eq(k, before[k], after[k]):
            diffs.append("%s: %s -> %s" % (k, before[k], after[k]))
    return (not diffs), diffs, changed


# ---------------------------------------------------------------- 主流程

def safe_close(pres, app=None):
    """关闭文稿。

    反复强杀 PowerPoint 会留下「文档恢复」状态，Presentation.Close 偶尔会报
    「Close : Failed.」——这类失败与业务逻辑无关，屏蔽掉即可；
    实在关不掉就先丢弃引用，后面 SaveAs 仍会正常覆盖。
    """
    try:
        pres.Close()
        return True
    except Exception:
        try:
            if app is not None:
                app.DisplayAlerts = 1          # ppAlertsNone
                pres.Close()
                return True
        except Exception:
            pass
        return False


def pptx_slide_count(path):
    """直接读 OPC 包里 ppt/slides 的部件数，避免再次启动 PowerPoint"""
    with zipfile.ZipFile(path) as z:
        return len([n for n in z.namelist()
                    if re.match(r"ppt/slides/slide\d+\.xml$", n)])


def main():
    for p in (RESULT,):
        if os.path.exists(p):
            os.remove(p)
    os.makedirs(OUTDIR, exist_ok=True)
    for f in os.listdir(OUTDIR):
        try:
            os.remove(os.path.join(OUTDIR, f))
        except OSError:
            pass

    if not os.path.exists(ADDIN):
        print("缺少产物：%s\n请先运行 build_suite.py --selftest" % ADDIN)
        return 2

    # 两张测试图片：尺寸与颜色都不同，便于确认确实换了图
    write_png(IMG_OLD, 200, 140, (192, 80, 77))
    write_png(IMG_NEW, 300, 200, (79, 129, 189))
    print("IMG  : 源图 %s  新图 %s" % (os.path.basename(IMG_OLD),
                                       os.path.basename(IMG_NEW)))

    app = com.DispatchEx("PowerPoint.Application")
    try:
        app.Visible = True
        try:
            app.DisplayAlerts = 1          # ppAlertsNone，抑制「文档恢复」等提示
        except Exception:
            pass
        added = app.AddIns.Add(ADDIN)
        added.Loaded = True
        print("ADDIN: name=%s  loaded=%s" % (added.Name, added.Loaded))

        # ① 5 页样张，供「选页另存」测试读取
        if os.path.exists(DECK5):
            os.remove(DECK5)
        deck5 = build_deck(app, 5)
        deck5.SaveAs(DECK5, 24)
        safe_close(deck5, app)
        print("DECK5: 5 页样张 -> %s" % os.path.basename(DECK5))

        # ② 图片替换样张（保存后关闭）
        #    先删掉上一轮的产物：万一这一轮 VBA 侧提前失败、根本没写出新文件，
        #    残留的旧 _replace_out.pptx 会被当成「本轮结果」比对，从而假通过。
        for stale in (REPLACE_IN, REPLACE_OUT, SELOUT, GROUP_IN, GROUP_OUT):
            try:
                if os.path.exists(stale):
                    os.remove(stale)
            except OSError:
                pass
        build_replace_deck(app, REPLACE_IN, IMG_OLD)
        print("REPL : 替换样张 -> %s（目标图带裁剪/旋转/翻转/亮度/透明度/边框/超链接，"
              "并被另一张图压住）" % os.path.basename(REPLACE_IN))

        # ③ 编组场景样张：目标图片位于一个带旋转 + 水平翻转的组合里，
        #    组合上下各压着一个形状（用于验证组合的 Z 序有没有还原）。
        build_group_deck(app, GROUP_IN, IMG_OLD)
        print("GRP  : 编组样张 -> %s（图片在带旋转+翻转的组合内）"
              % os.path.basename(GROUP_IN))

        # ③ 3 页文稿保持打开，作为 ActivePresentation 供图片导出测试
        pres = build_deck(app, 3)
        print("DECK : %d 页  %.3f x %.3f in"
              % (pres.Slides.Count, pres.PageSetup.SlideWidth / 72.0,
                 pres.PageSetup.SlideHeight / 72.0))

        # 若类型库已被 makepy/gencache 缓存，DispatchEx 会返回强类型对象，
        # 此时 Run 直接可用，_FlagAsMethod 反而不存在 —— 两种都兼容。
        if hasattr(app, "_FlagAsMethod"):
            app._FlagAsMethod("Run")
        refs = [
            "'%s.ppam'!modSelfTest.RunAllTests" % PROJ,
            "'%s'!modSelfTest.RunAllTests" % PROJ,
            "%s.ppam!modSelfTest.RunAllTests" % PROJ,
            "modSelfTest.RunAllTests",
        ]
        ok = False
        for r in refs:
            try:
                app.Run(r)
                print("RUN  : 成功（引用形式 %s）" % r)
                ok = True
                break
            except Exception as e:
                print("RUN  : 失败（%s）-> %s" % (r, e))
        if not ok:
            print("RUN_ERROR: 所有引用形式均失败")
    finally:
        try:
            app.Quit()
        except Exception:
            pass

    print("=" * 66)
    if os.path.exists(RESULT):
        print(open(RESULT, encoding="utf-8", errors="replace").read().strip())
    else:
        print("(未生成 vbresult.txt)")
        return 1

    all_ok = True

    # ---- 插件①：图片导出像素 ----
    print("=" * 66)
    print("① 导出图片 —— 输出文件像素校验")
    print("=" * 66)
    expects = [("T4K_001.png", (3840, 2160), png_size),
               ("TFHD_002.jpg", (1920, 1080), jpg_size),
               ("T300_003.png", (4000, 2250), png_size)]
    for fn, want, fn_size in expects:
        p = os.path.join(OUTDIR, fn)
        if not os.path.exists(p):
            print("  %-14s 缺失" % fn)
            all_ok = False
            continue
        got = fn_size(p)
        good = (got == want)
        all_ok = all_ok and good
        print("  %-14s 期望 %s  实际 %s  %s  (%d 字节)"
              % (fn, want, got, "通过" if good else "不通过", os.path.getsize(p)))

    # ---- 插件②：选页另存 ----
    print("=" * 66)
    print("② 选页另存 —— 选中 5 页中的第 2/4/5 页，期望保留 3 页")
    print("=" * 66)
    if os.path.exists(SELOUT):
        n = pptx_slide_count(SELOUT)
        good = (n == 3)
        all_ok = all_ok and good
        print("  _selected.pptx  期望 3 页  实际 %d 页  %s"
              % (n, "通过" if good else "不通过"))
    else:
        all_ok = False
        print("  _selected.pptx  缺失")

    # ---- 插件③：图片替换 ----
    print("=" * 66)
    print("③ 图片替换 —— 替换前后 OOXML 逐项比对（独立于 VBA 的第二重证据）")
    print("=" * 66)
    if os.path.exists(REPLACE_IN) and os.path.exists(REPLACE_OUT):
        b = pic_props(REPLACE_IN)
        a = pic_props(REPLACE_OUT)
        if b is None or a is None:
            all_ok = False
            print("  未能在样张中找到目标图片节点")
        else:
            same, diffs, changed = compare_xml(b, a)
            print("  目标图片在 spTree 中的序号：%d -> %d" % (b["index"], a["index"]))
            print("  图片引用变化            ：%s" % ("; ".join(changed) or "(无)"))
            if same:
                print("  逐项比对：全部一致（%d 项）✅" % (len(b) - 2))
            else:
                all_ok = False
                print("  逐项比对：发现 %d 处差异 ❌" % len(diffs))
                for d in diffs:
                    print("     [!] %s" % d)
            if not changed:
                all_ok = False
                print("     [!] 图片引用未变化，替换可能没有真正换图")
    else:
        all_ok = False
        print("  替换样张或替换结果缺失")

    # ---- 插件③（编组场景）：图片位于组合内的替换 ----
    print("=" * 66)
    print("③b 图片替换（编组内）—— 组合结构 + 图片属性的 OOXML 核对")
    print("=" * 66)
    vb = ""
    if os.path.exists(RESULT):
        vb = open(RESULT, encoding="utf-8", errors="replace").read()
    grp_line = ""
    for ln in vb.splitlines():
        if ln.startswith("IMG_GROUP="):
            grp_line = ln[len("IMG_GROUP="):]
            break
    if not grp_line:
        all_ok = False
        print("  未拿到 IMG_GROUP 结果行（VBA 自检可能没跑到该用例）")
    else:
        bad = [t for t in grp_line.split("|") if t and ("=DIFF" in t or "=FAIL" in t
                                                        or "=NOT_INGROUP" in t
                                                        or "=FIELDCOUNT" in t
                                                        or "=KEYDIFF" in t
                                                        or "=BADFMT" in t)]
        if bad:
            all_ok = False
            print("  VBA 自检发现问题 ❌：")
            for t in bad:
                print("     [!] %s" % t)
        else:
            print("  VBA 自检：组合与图片逐项一致 ✅")

        # 选区解析：覆盖用户的实际操作方式
        sel_tokens = [t for t in grp_line.split("|") if t.startswith("SEL.")]
        for t in sel_tokens:
            if t.startswith("SEL.CNT="):
                continue                      # 纯信息字段，不作断言
            if t == "SEL.WHOLE=OK:TGT_IMAGE":
                print("  单击组合（内含唯一图片）     -> 解析为 TGT_IMAGE ✅")
            elif t.startswith("SEL.IN.TGT_IMAGE="):
                if t == "SEL.IN.TGT_IMAGE=OK:TGT_IMAGE":
                    print("  进入组合点选图片             -> 解析为 TGT_IMAGE ✅")
                else:
                    all_ok = False
                    print("     [!] 进入组合点选图片却解析失败：%s" % t)
            elif t.startswith("SEL.IN.DECO_RECT="):
                # 点选组内的矩形（非图片）必须被拒绝。
                # 这里必须硬断言：早期版本只打印不判定，导致
                # 「点选矩形却被当成点选了图片」的错误一路蒙混过关。
                if t == "SEL.IN.DECO_RECT=REJECT":
                    print("  进入组合点选非图片（矩形）   -> 正确拒绝 ✅")
                else:
                    all_ok = False
                    print("     [!] 点选非图片却被解析成了图片：%s" % t)
            elif t == "SEL.MULTI=REJECTED":
                print("  选中含两张图片的组合         -> 正确拒绝并要求进组合点选 ✅")
            else:
                all_ok = False
                print("     [!] 选区解析异常：%s" % t)

    if os.path.exists(GROUP_IN) and os.path.exists(GROUP_OUT):
        b = group_props(GROUP_IN)
        a = group_props(GROUP_OUT)
        if b is None or a is None:
            all_ok = False
            print("  未能在样张中找到组合 GRP_TARGET")
        else:
            print("  组合成员构成：%s" % b["grp.members"])
            print("  组合在 spTree 中的序号：%s -> %s" % (b["grp.index"], a["grp.index"]))
            print("  组合旋转：%s -> %s  翻转H：%s -> %s"
                  % (b["grp.rot"], a["grp.rot"], b["grp.flipH"], a["grp.flipH"]))
            print("  图片引用变化：%s -> %s（包内媒体 %s -> %s）"
                  % (b.get("pic.embed"), a.get("pic.embed"),
                     b["media_count"], a["media_count"]))
            same, diffs, changed = compare_xml(b, a)
            if same:
                print("  逐项比对：全部一致（%d 项）✅" % (len(b) - 1))
            else:
                all_ok = False
                print("  逐项比对：发现 %d 处差异 ❌" % len(diffs))
                for d in diffs:
                    print("     [!] %s" % d)
            # 「真的换图了吗」的判据：包内媒体数必须增加。
            # 不能只看 r:embed —— 组内图片被替换后占用的仍是同一个 rId 槽位，
            # rId 可以完全不变（实测 rId3 -> rId3），但该 rId 已指向新的媒体文件。
            if a["media_count"] <= b["media_count"]:
                all_ok = False
                print("     [!] 包内媒体数未增加，替换可能没有真正换图")
            # 组内图片的裁剪比例必须严格保持（磅值会变，比例不该变）
            for ck in ("crop.l", "crop.t", "crop.r", "crop.b"):
                key = "pic." + ck
                if b.get(key) != a.get(key):
                    all_ok = False
                    print("     [!] %s 比例变化：%s -> %s" % (key, b.get(key), a.get(key)))
    else:
        all_ok = False
        print("  编组样张或结果缺失")

    print("-" * 66)
    print("回归测试总体：%s" % ("通过" if all_ok else "不通过"))

    trace = os.path.join(os.environ.get("TEMP", ""), "PptImageReplace-trace.log")
    if os.path.exists(trace):
        print("=" * 66)
        print("图片替换 · 步骤追踪（%s）" % trace)
        print("=" * 66)
        print(open(trace, encoding="utf-8", errors="replace").read().strip())

    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
