# -*- coding: utf-8 -*-
"""
构建产物代码核对（回归闸门）。

把「源文件 + 注入代码」——也就是构建时*应当*写进 VBA 工程的内容——
与「从 .ppam 里实际提取出来的代码」逐行比对。

这道检查专抓一类极其隐蔽的问题：

    AddFromString 注入之后立刻运行**不会**暴露的问题，
    只有保存进包、被 VBA 重新解析时才显现。
    本项目的实例：模块级声明被写在过程之后 → VBE 静默丢弃 → 保存后
    所有引用它的过程都变成「变量未定义」；而加载项的工程是「隐藏」的，
    PowerPoint 只会报一句无法定位的「隐藏的模块中的编译错误」。
    先在这里逐行核对「构建时该写什么」，能把问题挡在装进 PowerPoint 之前。

比对口径（重要）：
    VBE 保存时会按自己的规则重写标识符大小写 —— 凡它没认出是内置函数 /
    属性的标识符一律强制转小写（P( → p(、Value → value、Path → path ...）。
    这些改写对 VBA 完全无害（VBA 本来就大小写不敏感），却是逐行比对时
    最大的噪声源。因此本工具在**字符串字面量之外**忽略大小写，
    字符串内部（自定义功能区回调名、文件路径等）仍严格比较。
    曾经把这类无害改写当成「Dir$ 不存在」的证据，是走过的弯路。

用法：
    python test/verify_build_code.py                    # 核对含自检注入的测试产物
    python test/verify_build_code.py --release          # 核对正式产物（不含注入）
"""

import difflib
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SUITE = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(SUITE, "build"))
import build_suite as bs           # noqa: E402

try:
    from oletools.olevba import VBA_Parser
except ImportError:
    sys.stderr.write("需要 oletools：pip install oletools\n")
    sys.exit(2)


def strip_attr(code):
    return "\r\n".join(
        l for l in code.replace("\r\n", "\n").split("\n")
        if not l.strip().lower().startswith("attribute "))


def build_expected(plugins, patches):
    """产出 {模块名: 应当写入的代码}"""
    expect = {}
    for p in plugins:
        for comp in p["components"]:
            path = os.path.join(p["_dir"], comp["file"])
            code = strip_attr(bs.read_text(path))
            patch = patches.get(comp["name"])
            if patch:
                code = code + "\r\n" + \
                    "\r\n".join(patch.replace("\r\n", "\n").split("\n"))
            expect[comp["name"]] = code
    return expect


def fold_outside_strings(line):
    """把「字符串字面量之外」的部分转小写，字符串内部原样保留。

    为什么需要：VBE 保存时会按自己的规则重写标识符大小写 —— 凡它没认出
    是内置函数 / 属性的标识符一律强制转小写（P( → p(、Value → value、
    Path → path、OutDir → outdir ...）。这些改写对 VBA 完全无害（VBA 本身
    大小写不敏感），却是逐行比对时最大的噪声来源。

    为什么不干脆整行转小写：自定义功能区回调名、文件路径等都在字符串
    字面量里，大小写是有意义的，折叠掉就可能漏掉真问题。
    """
    out = []
    i, n = 0, len(line)
    in_str = False
    while i < n:
        ch = line[i]
        if ch == '"':
            out.append(ch)
            if in_str and i + 1 < n and line[i + 1] == '"':
                out.append('"')          # VBA 里 "" 表示一个字面双引号
                i += 2
                continue
            in_str = not in_str
        else:
            out.append(ch if in_str else ch.lower())
        i += 1
    return "".join(out)


def normalize(code):
    lines = [l.rstrip() for l in code.replace("\r\n", "\n").split("\n")]
    lines = [l for l in lines if not l.strip().lower().startswith("attribute ")]
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return lines


def main():
    release = "--release" in sys.argv
    suite = bs.load_suite()
    addin = os.path.join(SUITE, "dist",
                         suite["suite"]["addin_filename"] if release
                         else "PptEnhanceSuiteTest.ppam")
    if not os.path.exists(addin):
        print("缺少产物：%s" % addin)
        return 2

    plugins = bs.load_plugins()
    patches = {}
    if not release:
        patches = {
            "frmImageExport": bs.read_text(os.path.join(HERE, "selftest.vba")),
            "m_SlideExport": bs.read_text(os.path.join(HERE, "selftest_slide.vba")),
            "m_ImageReplace": bs.read_text(os.path.join(HERE, "selftest_replace.vba")),
        }
        plugins.append({
            "id": "_test", "order": 9999, "title": "自检驱动", "group": None,
            "buttons": [],
            "components": [{"type": "module", "name": "modSelfTest",
                            "file": "modSelfTest.bas"}],
            "_dir": HERE, "_dirname": "_test",
        })

    got_map = {}
    vp = VBA_Parser(addin)
    for (fn, stream, name, code) in vp.extract_macros():
        got_map[str(name).split(".")[0]] = code
    vp.close()

    print("=" * 66)
    print("构建产物代码核对：%s" % os.path.basename(addin))
    print("=" * 66)

    expect = build_expected(plugins, patches)
    all_ok = True
    for name, exp in expect.items():
        got = got_map.get(name)
        if got is None:
            print("  %-18s 产物内缺失 ❌" % name)
            all_ok = False
            continue
        e, g = normalize(exp), normalize(got)
        # 判据：字符串字面量之外忽略大小写（VBA 大小写不敏感）
        ef = [fold_outside_strings(x) for x in e]
        gf = [fold_outside_strings(x) for x in g]
        if ef == gf:
            note = ""
            if e != g:
                ndiff = sum(1 for x, y in zip(e, g) if x != y)
                note = "（其中 %d 行仅 VBE 大小写规范化差异，无害）" % ndiff
            print("  %-18s %4d 行  一致 ✅%s" % (name, len(e), note))
            continue
        all_ok = False
        print("  %-18s %4d 行 vs %4d 行  存在差异 ❌" % (name, len(e), len(g)))
        diff = list(difflib.unified_diff(e, g, "源文件", "产物内", lineterm="", n=2))
        for line in diff[:40]:
            print("      " + line)
        if len(diff) > 40:
            print("      ... 共 %d 行差异" % len(diff))

    print("-" * 66)
    print("代码核对总体：%s" % ("通过" if all_ok else "不通过"))
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
