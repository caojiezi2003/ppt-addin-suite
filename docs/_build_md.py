# -*- coding: utf-8 -*-
"""由《使用说明.html》生成 Markdown 版《使用说明.md》。

为什么要自动转换而不是另写一份：
    HTML 版是自包含单文件（截图 base64 内嵌），Markdown 版则引用同目录 screenshots/。
    两份内容由同一脚本链自动保持一致 —— 改内容只需要改 _build_doc.py 里的 HTML 模板，
    重跑一次即可同步产出 HTML / PDF / MD 三份，不会出现版本漂移。

    python _build_md.py        # 也可由 _build_doc.py 末尾自动调用

Markdown 版约定：
    * 宽幅截图（原始宽度 >960px）用 <img width="960"> 约束显示宽度，
      Markdown 原生语法不支持指定宽度，HTML 标签是最兼容的折中；
    * 提示框转成引用块，参数表格转成 Markdown 表格；
    * 图片路径为相对路径，分发时请连同 screenshots/ 目录一起带走。
"""
import io
import os
import re
from html.parser import HTMLParser

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "使用说明.html")
OUT = os.path.join(HERE, "使用说明.md")

VOID = {"br", "img", "hr", "meta", "link", "input"}
SKIP = {"head", "style", "script", "title", "meta", "link"}

# 宽图的显示宽度上限（Markdown 里靠 HTML img 的 width 属性实现）
WIDE = {
    "ribbon_full.png": 960,
    "slidesel.png": 960,
    "selimg.png": 960,
    "replace_before_after.png": 960,
    "slidesave.png": 880,
    "replace_pick.png": 880,
}


class Node(object):
    __slots__ = ("tag", "attrs", "text", "children")

    def __init__(self, tag=None, attrs=None, text=None):
        self.tag = tag
        self.attrs = dict(attrs or {})
        self.text = text
        self.children = []

    @property
    def cls(self):
        return self.attrs.get("class", "").split()


class DocParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self, convert_charrefs=True)
        self.root = Node("root")
        self.stack = [self.root]

    def handle_starttag(self, tag, attrs):
        node = Node(tag, attrs)
        self.stack[-1].children.append(node)
        if tag not in VOID:
            self.stack.append(node)

    def handle_startendtag(self, tag, attrs):
        self.stack[-1].children.append(Node(tag, attrs))

    def handle_endtag(self, tag):
        if tag in VOID:
            return
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                return

    def handle_data(self, data):
        # 丢弃纯缩进/换行空白
        if not data.strip():
            return
        # HTML 源码里的缩进换行只是为了源码可读（浏览器渲染成空格），
        # 归一化成单个空格：否则在 Markdown 表格里会撑断表格、在列表里会变成缩进
        self.stack[-1].children.append(
            Node(None, None, re.sub(r"\s*\n\s*", " ", data)))


def text_of(node):
    if node.tag is None:
        return node.text
    return "".join(text_of(c) for c in node.children)


def img_md(node):
    name = node.attrs.get("alt", "")
    src = "screenshots/" + name
    label = os.path.splitext(name)[0]
    w = WIDE.get(name)
    if w:
        return '<img src="%s" alt="%s" width="%d">' % (src, label, w)
    return "![%s](%s)" % (label, src)


def inner(kids):
    """行内渲染；连续的「药丸」标签之间补上分隔符。"""
    parts = []
    prev_pill = False
    for k in kids:
        cur_pill = k.tag == "span" and "pill" in k.cls
        if prev_pill and cur_pill:
            parts.append(" / ")
        parts.append(inline(k))
        if not (k.tag is None and not k.text.strip()):
            prev_pill = cur_pill
    return "".join(parts)


def inline(node):
    if node.tag is None:
        return node.text
    tag, cls = node.tag, node.cls
    if tag == "br":
        return "<br>"
    if tag == "img":
        return img_md(node)
    if tag in ("b", "strong"):
        return "**" + inner(node.children).strip() + "**"
    if tag == "code":
        return "`" + text_of(node).strip() + "`"
    if tag == "a":                       # 锚点在 Markdown 里不可靠，只保留文字
        return inner(node.children).strip()
    if tag == "span" and "pill" in cls:
        return "`" + text_of(node).strip() + "`"
    if tag == "span" and "en" in cls:
        return "（" + text_of(node).strip() + "）"
    if tag == "span" and "ver" in cls:
        return " " + text_of(node).strip()
    return inner(node.children)


def emit_table(node, out):
    rows = []
    for tr in node.children:
        if tr.tag == "tr":
            rows.append([inline(td) for td in tr.children if td.tag in ("td", "th")])
    if not rows:
        return
    ncol = max(len(r) for r in rows)
    rows = [r + [""] * (ncol - len(r)) for r in rows]

    def clean(s):
        s = re.sub(r"\s*<br\s*/?>\s*", "<br>", s)   # <br> 两侧不留空格
        s = re.sub(r"\s+", " ", s)
        return s.replace("|", "\\|").strip()

    out.append("| " + " | ".join(clean(c) for c in rows[0]) + " |")
    out.append("| " + " | ".join(["---"] * ncol) + " |")
    for r in rows[1:]:
        out.append("| " + " | ".join(clean(c) for c in r) + " |")
    out.append("")


def render(node, out):
    tag = node.tag
    if tag is None:
        s = node.text.strip()
        if s:
            out.append(s)
        return
    cls = node.cls
    if tag in SKIP:
        return
    if tag in ("html", "body") or (tag == "div" and "page" in cls):
        for c in node.children:
            render(c, out)
        return

    if tag == "header":
        out.append("# " + inner(node.children[0].children).strip())
        out.append("")
        for c in node.children[1:]:
            if c.tag == "div" and "sub" in c.cls:
                out.append("*" + text_of(c).strip() + "*")
        out.append("")
        return

    if tag == "nav":                     # 目录：不用链接，避免各渲染器锚点规则不一造成死链
        out.append("## 目录")
        out.append("")
        for ol in node.children:
            if ol.tag != "ol":
                continue
            n = 0
            for li in ol.children:
                if li.tag == "li":
                    n += 1
                    out.append("%d. %s" % (n, inline(li).strip()))
        out.append("")
        return

    if tag in ("h2", "h3"):
        out.append("")
        out.append(("#" * int(tag[1])) + " " + inner(node.children).strip())
        out.append("")
        return

    if tag == "p":
        txt = inner(node.children).strip()
        if txt:
            out.append(txt)
            out.append("")
        return

    if tag == "table":
        emit_table(node, out)
        return

    if tag == "ol":
        start = int(node.attrs.get("start", 1))
        i = 0
        for li in node.children:
            if li.tag == "li":
                out.append("%d. %s" % (start + i, inner(li.children).strip()))
                i += 1
        out.append("")
        return

    if tag == "div" and "box" in cls:
        out.append("> " + inner(node.children).strip())
        out.append("")
        return

    if tag == "img":
        out.append("")
        out.append(img_md(node))
        out.append("")
        return

    if tag == "footer":
        out.append("---")
        out.append("")
        for line in inner(node.children).split("<br>"):
            line = line.strip()
            if line:
                out.append(line)
        out.append("")
        return

    for c in node.children:              # 兜底：继续下钻
        render(c, out)


def main():
    src = io.open(SRC, encoding="utf-8").read()
    parser = DocParser()
    parser.feed(src)
    parser.close()

    out = []
    render(parser.root, out)
    md = "\n".join(out)
    md = re.sub(r"[ \t]+\n", "\n", md)
    # 软换行在浏览器里会被渲染成空格，中文之间这个空格删掉更符合中文排版
    cjk = r"\u3000-\u303f\u4e00-\u9fff\uff00-\uffef"
    md = re.sub(r"(?<=[%s])[ \t]+(?=[%s])" % (cjk, cjk), "", md)
    # 中文标点后紧跟加粗标记时也不要留空格
    md = re.sub(r"(?<=[，。、：；！？）】」])[ \t]+(?=\*\*)", "", md)
    md = re.sub(r"\n{3,}", "\n\n", md).strip() + "\n"
    md += ("\n<!-- 配图位于同目录 screenshots/ 下；本文与 使用说明.html / PDF 版内容一致，"
           "由 _build_doc.py 统一生成。 -->\n")

    io.open(OUT, "w", encoding="utf-8", newline="\n").write(md)
    print("[md  ] written: %s  (%.1f KB, %d 行)" %
          (OUT, len(md.encode("utf-8")) / 1024.0, md.count("\n")))
    return OUT


if __name__ == "__main__":
    main()
