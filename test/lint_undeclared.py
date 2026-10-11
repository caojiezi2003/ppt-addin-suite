# -*- coding: utf-8 -*-
"""
未声明标识符静态扫描（`Option Explicit` 编译失败的专用排查工具）。

为什么需要它：
    `.ppam` 加载项的 VBA 工程是「隐藏」的，编译失败只会弹一句
    「隐藏的模块中的编译错误: <模块名>」，既不说是哪一行，也不说是哪个标识符。
    而这类错误的成因往往只有一个：**某个标识符在 Option Explicit 下没有声明**。

思路（借用 Bootstrap 的巧劲）：
    本项目里已经有一批**确认能编译**的模块（m_Core / m_ImageExport /
    m_SlideExport / frmImageExport，它们早已通过真机验证）。
    那么「这些模块里出现过的裸标识符」就构成了一份可信的白名单。
    新模块里一旦冒出白名单外的裸标识符，就极可能是拼写错误或漏声明。

用法：
    python test/lint_undeclared.py            # 扫描所有未被列为「可信」的模块
    python test/lint_undeclared.py -v         # 连未在白名单但可能是内置的也列出
"""

import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SUITE = os.path.dirname(HERE)
SRC = os.path.join(SUITE, "src")

# 已经真机验证过、确认可编译的模块 —— 作为可信语料库
TRUSTED = [
    os.path.join(SRC, "_core", "m_Core.bas"),
    os.path.join(SRC, "image-export", "m_ImageExport.bas"),
    os.path.join(SRC, "image-export", "frmImageExport.code.vba"),
    os.path.join(SRC, "slide-export", "m_SlideExport.bas"),
]

# 待扫描的模块
TARGETS = [
    os.path.join(SRC, "image-replace", "m_ImageReplace.bas"),
    os.path.join(HERE, "selftest_replace.vba"),
    os.path.join(HERE, "modSelfTest.bas"),
]

VBA_KEYWORDS = set("""
and as boolean byref byval call case cbool cbyte ccur cdate cdbl cdec chr cint
circle class close clng const csng cstr cvar date declare defbool defbyte defcur
defdate defdbl defdec defint deflng deflngptr defobj defsng defstr defvar dim do
double each else elseif empty end endif enum eqv erase error event exit false for
friend function get global gosub goto if imp implements in integer is let lib like
lock long longlong loop lset me mod new next not nothing null object on option
optional or paramarray preserve private property public raiseevent redim rem
resume return select set single static step stop string sub then to true type
typeof unload until variant wend while with withevents xor
""".split())

VBA_BUILTINS = set("""
abs array asc ascw atn cbool cbyte ccur cdate cdbl cdec chr chrw cint clng clnglng
clngptr cos createobject csng cstr cvar cvdate cverr date dateadd datediff
datepart dateserial datevalue day ddb dir doevents environ eof error exp filelen
fix format freefile fv getobject hex hour iif imestatus input inputbox instr
instrrev int ipmt irr isarray isdate isempty iserror ismissing isnull isnumeric
isobject join lbound lcase left len loc lof log ltrim mid minute mirr month
monthname msgbox now nper npv oct partition pmt ppmt pv qbcolor rate replace rgb
right rnd round rtrim second seek sgn sin slk space spc split sqr strcomp strconv
string strreverse switch syscolor tab tan time timer timeserial timevalue trim
typename ubound ucase val vartype weekday weekdayname year
""".split())

# 常见 VBA / Office 枚举常量与全局对象（可信模块里若用过就已被白名单覆盖，
# 这里再兜一层，避免把明显合法的东西报出来）
COMMON_ENUMS = set("""
vbcr vbcrlf vblf vbtab vbnullstring vbnullchar vbobjecterror vbexclamation
vbinformation vbquestion vbcritical vbokonly vbokcancel vbyesno vbyescancel
vbapplicationmodal vbsystemmodal vbmodal vbmodeless vbnormal vbnormalFocus
vbmaximizedfocus vbminimizedfocus vbhide vbshownormal vbhidden vbreadonly
vbdirectory vbnormalfocus vbtextcompare vbbinarycompare vbfromunicode vbunicode
vbleftbuttondown vbleftbuttonup vbleftbuttonclick vbkeydown vbkeyup vbkeyreturn
vbkeyescape vbkeydelete vbnewline vbtab vbaggregate vbdatestring vbnumeric
msofalse msotrue msofliphorizontal msoflipvertical msobringtofront msosendtoback
msobringforward msosendbackward msoautomatic msoautosize msofit
msoshaperectangle msoshapeoval msopictureautomatic msopicturegrayscale
msohyperlink msoactionhyperlink msoanimnone ppLayoutBlank pplayouttitle
ppsaveasopencxmlpresentation ppSaveAsOpenXMLPresentation
ppSaveAsOpenXMLPresentationMacroEnabled ppsaveasdefault
ppwindowminimized ppwindownormal ppwindowmaximized
ppSelectionNone ppSelectionSlides ppSelectionShapes ppSelectionText
ppeffectnone ppEffectNone
xlcontinuous xlinsidehorizontal
""".lower().split())


def strip_comments_and_strings(text):
    """去掉注释与字符串字面量，保留行结构"""
    out = []
    for line in text.replace("\r\n", "\n").split("\n"):
        buf, i, in_str = [], 0, False
        while i < len(line):
            ch = line[i]
            if in_str:
                if ch == '"':
                    if i + 1 < len(line) and line[i + 1] == '"':
                        i += 2
                        continue
                    in_str = False
                i += 1
                continue
            if ch == '"':
                in_str = True
                buf.append('""')
                i += 1
                continue
            if ch == "'":
                break
            buf.append(ch)
            i += 1
        out.append("".join(buf))
    return "\n".join(out)


DECL_PATTERNS = [
    # Dim / Public / Private / Global / Static [Const] Name ...
    re.compile(r"^\s*(?:Public|Private|Global|Dim|Static|Friend)\s+"
               r"(?:Const\s+)?([A-Za-z_]\w*)", re.I | re.M),
    # 同行逗号连续声明： Dim a As Long, b As String
    re.compile(r",\s*([A-Za-z_]\w*)\s+As\b", re.I),
    re.compile(r",\s*([A-Za-z_]\w*)\s*(?:\(|,|$)", re.I | re.M),
    # 过程声明
    re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*"
               r"(?:Sub|Function)\s+([A-Za-z_]\w*)", re.I | re.M),
    re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)*"
               r"Property\s+(?:Get|Let|Set)\s+([A-Za-z_]\w*)", re.I | re.M),
    # Type / Enum 及其成员
    re.compile(r"^\s*(?:Public\s+|Private\s+)?(?:Type|Enum)\s+([A-Za-z_]\w*)",
               re.I | re.M),
    # 参数
    re.compile(r"(?:ByVal|ByRef|Optional|ParamArray)\s+([A-Za-z_]\w*)", re.I),
    # For / For Each
    re.compile(r"\bFor\s+Each\s+([A-Za-z_]\w*)", re.I),
    re.compile(r"\bFor\s+([A-Za-z_]\w*)\s*=", re.I),
    # 行号标签
    re.compile(r"^\s*([A-Za-z_]\w*):\s*$", re.M),
    # Line Input #n, x  / Open ... As #n
    re.compile(r"\bAs\s+#([A-Za-z_]\w*)", re.I),
]


def declared_names(text):
    code = strip_comments_and_strings(text)
    names = set()
    for pat in DECL_PATTERNS:
        for m in pat.finditer(code):
            names.add(m.group(1).lower())
    # 参数表：取 Sub/Function 行括号里的形参
    for m in re.finditer(r"(?:Sub|Function)\s+\w+\s*\(([^)]*)\)",
                         code, re.I | re.S):
        for part in m.group(1).split(","):
            part = part.strip()
            if not part:
                continue
            part = re.sub(r"^(?:ByVal|ByRef|Optional|ParamArray)\s+", "",
                          part, flags=re.I)
            mm = re.match(r"([A-Za-z_]\w*)", part)
            if mm:
                names.add(mm.group(1).lower())
    return names


IDENT = re.compile(r"(?<![.\w$])([A-Za-z_]\w*)\$?(?![\w$])")


def bare_identifiers(text):
    """取所有「裸标识符」——跳过 .成员、跳过 := 命名参数、跳过声明关键字行"""
    code = strip_comments_and_strings(text)
    used = {}
    for ln, line in enumerate(code.split("\n"), 1):
        s = line
        if s.lstrip().lower().startswith(("attribute ", "rem ")):
            continue
        for m in IDENT.finditer(s):
            nm = m.group(1)
            # 形如 "obj.Name" 的成员：前面紧跟 '.'，正则已排除
            # 形如 "Name:=x" 的命名参数
            after = s[m.end():]
            if after.lstrip().startswith(":="):
                continue
            used.setdefault(nm.lower(), []).append((ln, nm, s.strip()))
    return used


def collect_known(paths):
    known = set()
    for p in paths:
        if not os.path.exists(p):
            continue
        txt = open(p, encoding="utf-8").read()
        known |= declared_names(txt)
        for k in bare_identifiers(txt):
            known.add(k)
    return known


def main():
    verbose = "-v" in sys.argv
    known = collect_known(TRUSTED)
    print("=" * 70)
    print("未声明标识符扫描")
    print("=" * 70)
    print("可信语料库：%d 个已知合法标识符（来自已真机验证的模块）" % len(known))
    print()

    bad_total = 0
    for t in TARGETS:
        if not os.path.exists(t):
            print("  %-24s 文件不存在，跳过" % os.path.basename(t))
            continue
        txt = open(t, encoding="utf-8").read()
        decl = declared_names(txt)
        used = bare_identifiers(txt)
        suspects = {}
        for k, spots in used.items():
            if k in known or k in decl:
                continue
            if k in VBA_KEYWORDS or k in VBA_BUILTINS or k in COMMON_ENUMS:
                continue
            suspects[k] = spots

        print("-" * 70)
        print("%s   （声明 %d 个 / 使用 %d 个裸标识符）"
              % (os.path.basename(t), len(decl), len(used)))
        if not suspects:
            print("    未发现可疑标识符 ✅")
            continue
        bad_total += len(suspects)
        for k in sorted(suspects):
            spots = suspects[k]
            print("    [?] %-24s 出现 %d 次，首次在第 %d 行：%s"
                  % (k, len(spots), spots[0][0], spots[0][2][:78]))
            if verbose:
                for ln, orig, s in spots[1:6]:
                    print("          第 %d 行：%s" % (ln, s[:78]))
        print()

    print("=" * 70)
    print("可疑标识符合计：%d" % bad_total)
    return 1 if bad_total else 0


if __name__ == "__main__":
    sys.exit(main())
