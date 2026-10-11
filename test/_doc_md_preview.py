# -*- coding: utf-8 -*-
"""把 docs/使用说明.md 用 GitHub 风格渲染成 HTML，便于浏览器截图核对 Markdown 版效果。

    python test/_doc_md_preview.py <输出html路径，默认 docs/_md_preview.html>

只用于验证，不参与交付。
"""
import io
import os
import sys

import markdown

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(os.path.dirname(HERE), "docs")
SRC = os.path.join(DOCS, "使用说明.md")
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DOCS, "_md_preview.html")

CSS = """
body{max-width:1012px;margin:0 auto;padding:24px 32px;color:#1f2328;background:#fff;
 font-family:"Microsoft YaHei UI","Segoe UI",sans-serif;font-size:16px;line-height:1.75}
h1{font-size:2em;border-bottom:1px solid #d1d9e0;padding-bottom:.3em}
h2{border-bottom:1px solid #d1d9e0;padding-bottom:.3em;margin-top:36px;font-size:1.5em}
h3{margin-top:28px;font-size:1.2em}
table{border-collapse:collapse;width:100%;font-size:15px;margin:16px 0}
th,td{border:1px solid #d1d9e0;padding:6px 13px;vertical-align:top}
th{background:#f6f8fa}
blockquote{border-left:4px solid #d1d9e0;margin:0 0 16px;padding:0 1em;color:#59636e}
code{background:#eff1f3;padding:.2em .4em;border-radius:6px;font-size:85%}
img{max-width:100%;height:auto;border:1px solid #e4e8ec;border-radius:6px}
li{margin:4px 0}
"""


def main():
    md = io.open(SRC, encoding="utf-8").read()
    body = markdown.markdown(md, extensions=["tables", "fenced_code",
                                             "sane_lists", "md_in_html"])
    html = ('<!DOCTYPE html><html><head><meta charset="utf-8">'
            "<style>%s</style></head><body>%s</body></html>" % (CSS, body))
    io.open(OUT, "w", encoding="utf-8", newline="\n").write(html)
    print("preview:", OUT, len(html), "bytes")


if __name__ == "__main__":
    main()
