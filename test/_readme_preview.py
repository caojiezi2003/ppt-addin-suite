# -*- coding: utf-8 -*-
"""把 README.md 按 GitHub 风格渲染成临时 HTML，用于截图目视检查。

放在仓库根目录下生成，图片相对路径 docs/screenshots/… 才能正常加载。
"""
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "README.md")
OUT = os.path.join(ROOT, "_readme_preview.html")

CSS = """
body{max-width:1012px;margin:0 auto;padding:24px 32px;
 font-family:-apple-system,"Segoe UI","Microsoft YaHei UI",sans-serif;
 font-size:16px;line-height:1.75;color:#1f2328;background:#fff}
h1{font-size:2em;border-bottom:1px solid #d1d9e0;padding-bottom:.3em}
h2{border-bottom:1px solid #d1d9e0;padding-bottom:.3em;margin-top:40px;font-size:1.5em}
h3{margin-top:28px;font-size:1.2em}
table{border-collapse:collapse;width:100%;font-size:15px;margin:16px 0}
th,td{border:1px solid #d1d9e0;padding:6px 13px;vertical-align:top}
th{background:#f6f8fa}
blockquote{border-left:4px solid #d1d9e0;margin:0 0 16px;padding:0 1em;color:#59636e}
code{background:#eff1f3;padding:.2em .4em;border-radius:6px;font-size:85%}
img{max-width:100%;height:auto;border:1px solid #e4e8ec;border-radius:6px}
li{margin:4px 0}
"""

try:
    import markdown
except ImportError:
    sys.stderr.write("需要 python-markdown：pip install markdown\n")
    sys.exit(1)

md = io.open(SRC, encoding="utf-8").read()
body = markdown.markdown(md, extensions=["tables", "fenced_code", "sane_lists"])
h = ("<!DOCTYPE html><html><head><meta charset=\"utf-8\"><style>" + CSS +
     "</style></head><body>" + body + "</body></html>")
io.open(OUT, "w", encoding="utf-8", newline="\n").write(h)
# 统计渲染结果里的图片与链接，便于核对
print("preview: %s  (%d 字符)" % (OUT, len(h)))
print("  <img> 标签：%d 个" % body.count("<img"))
print("  <a>   链接：%d 个" % body.count("<a "))
print("  <table>：%d 个" % body.count("<table"))
