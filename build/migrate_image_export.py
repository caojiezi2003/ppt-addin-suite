# -*- coding: utf-8 -*-
"""
把独立项目 ppt-image-exporter 的源码迁入套件目录，并做命名规范化。

一次性脚本；迁移完成后源码即归套件维护，本脚本保留作追溯用途。

命名映射：
    modPptImageExport   -> m_ImageExport          （模块）
    frmPptImageExport   -> frmImageExport         （窗体）
    RibbonExportImages  -> Ribbon_ImageExport_Export（Ribbon 回调）
    RibbonGetEnabled    -> 移交 m_Core.bas（通用回调，避免跨插件重名）
"""

import os
import re
import shutil

HERE = os.path.dirname(os.path.abspath(__file__))
SUITE = os.path.dirname(HERE)
OLD_SRC = os.path.join(os.path.dirname(SUITE), "ppt-image-exporter", "src")
NEW_DIR = os.path.join(SUITE, "src", "image-export")

os.makedirs(NEW_DIR, exist_ok=True)

report = []


def migrate(text, name):
    """对单个源文件做重命名替换"""
    hits = []

    def sub(pattern, repl, label):
        nonlocal text
        new, n = re.subn(pattern, repl, text)
        if n:
            hits.append("%s x%d" % (label, n))
        text = new

    sub(r"\bmodPptImageExport\b", "m_ImageExport", "modPptImageExport->m_ImageExport")
    sub(r"\bfrmPptImageExport\b", "frmImageExport", "frmPptImageExport->frmImageExport")
    sub(r"\bRibbonExportImages\b", "Ribbon_ImageExport_Export",
        "RibbonExportImages->Ribbon_ImageExport_Export")

    # RibbonGetEnabled 是通用回调，统一放在 m_Core.bas，这里移除以免重名
    new, n = re.subn(
        r"Public Sub RibbonGetEnabled\(control As IRibbonControl, ByRef returnedVal\)"
        r"\r?\n[^\n]*returnedVal = True[^\n]*\r?\nEnd Sub\r?\n\r?\n",
        "", text)
    if n:
        hits.append("移除 RibbonGetEnabled x%d（已移至 _core/m_Core.bas）" % n)
    text = new

    # 文件头 Attribute 行同步改名
    sub(r'Attribute VB_Name = "m_ImageExport"', 'Attribute VB_Name = "m_ImageExport"', "VB_Name")
    report.append("  %-34s %s" % (name, "；".join(hits) if hits else "(无改动)"))
    return text


print("=" * 66)
print("迁入 image-export")
print("=" * 66)

# 1) 引擎模块
src = os.path.join(OLD_SRC, "modPptImageExport.bas")
code = open(src, encoding="utf-8").read()
code = migrate(code, "modPptImageExport.bas")
code = code.replace('Attribute VB_Name = "modPptImageExport"',
                    'Attribute VB_Name = "m_ImageExport"')
with open(os.path.join(NEW_DIR, "m_ImageExport.bas"), "w", encoding="utf-8",
          newline="\r\n") as f:
    f.write(code)

# 2) 窗体代码模块
src = os.path.join(OLD_SRC, "frmPptImageExport.code.vba")
form = open(src, encoding="utf-8").read()
form = migrate(form, "frmPptImageExport.code.vba")
with open(os.path.join(NEW_DIR, "frmImageExport.code.vba"), "w", encoding="utf-8",
          newline="\r\n") as f:
    f.write(form)

print("\n".join(report))

# 3) 残留检查
print("\n--- 残留旧名检查 ---")
bad = 0
for fn in ("m_ImageExport.bas", "frmImageExport.code.vba"):
    t = open(os.path.join(NEW_DIR, fn), encoding="utf-8").read()
    for token in ("modPptImageExport", "frmPptImageExport", "RibbonExportImages",
                  "RibbonGetEnabled"):
        if token in t:
            print("  [!] %s 仍含 %s" % (fn, token))
            bad += 1
print("  残留：%d 处" % bad)

print("\n--- 产出 ---")
for fn in sorted(os.listdir(NEW_DIR)):
    p = os.path.join(NEW_DIR, fn)
    n = open(p, encoding="utf-8").read().count("\n") + 1
    print("  %-34s %5d 行  %6.1f KB" % (fn, n, os.path.getsize(p) / 1024.0))
