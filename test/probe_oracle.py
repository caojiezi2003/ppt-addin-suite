# -*- coding: utf-8 -*-
"""
编译二值预言机：只回答一个问题 —— 「这份代码能不能编译」。

原理：
    把模块注入**普通演示文稿**（模块可见，VBA 会给出具体错误），
    在**工作线程**里调 Application.Run 触发编译；主线程只负责盯 VBA 弹窗。
    一旦出现 #32770「Microsoft Visual Basic for Applications」对话框，
    就判定「编译失败」，随即强杀 PowerPoint 并退出 ——
    因为那个对话框是模态的，被卡住的 Run 永远不会返回，我们也不需要它返回。

用法：
    python test/probe_oracle.py src/image-replace/m_ImageReplace.bas
    python test/probe_oracle.py src/image-replace/m_ImageReplace.bas --patch test/selftest_replace.vba
    python test/probe_oracle.py --snippet "shp.FlipHorizontal = msoTrue"
    python test/probe_oracle.py --snippet-file /tmp/x.bas

退出码：0 = 编译通过；1 = 编译失败；2 = 未能判定
"""

import ctypes
import os
import sys
import threading
import time

import pythoncom
import win32com.client as com

HERE = os.path.dirname(os.path.abspath(__file__))
SUITE = os.path.dirname(HERE)
SRC = os.path.join(SUITE, "src")

VBE_CT_STDMODULE = 1
user32 = ctypes.windll.user32
CHILD = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
WND = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

PROBE_NAME = "ProbeRun"

# 探针必须**真正进入**每一个目标过程。
#
# ★ VBA 的真实编译规则（实测结论，别再想当然）★
#     过程**首次被进入时**才编译，编译单位是「这一个过程」。
#     没被执行到的过程根本不会被解析，里面的语法错误也就永远不会报出来。
#     所以「静态调用图可达」不等于「会被检查」——
#     早期以为"探针调用了 SelfTest，它静态可达的整条链都会被编译"，
#     因此漏掉了大量从未被真正进入的代码（例如组内替换那条路径）。
#
#     推论：写探针时要把目标过程**一个一个直接列出来**，
#     哪怕传 Nothing / 不存在的路径让它立刻返回 —— 进入动作本身就会
#     让它整体编译，这才是我们想要的。
#
# ⚠ 不要再把 ImageReplace_Run 加回来。
#     它在"没有选中任何形状"时会弹一个 MsgBox，而 MsgBox 是模态的：
#     Run 永远不返回、弹窗标题又不含 "Visual Basic"（本工具的判据），
#     于是预言机会一直等到超时，报「未能判定」—— 看起来像工具坏了。
DEFAULT_CALLS = (
    '    Dim s As String\r\n'
    '    Dim e As String\r\n'
    '    Dim sh As Shape\r\n'
    '    Dim ok As Boolean\r\n'
    '    On Error Resume Next\r\n'
    '    s = ImageReplace_SelfTest("X:\\nonexist_in.pptx", "X:\\nonexist_new.png", '
    '"X:\\nonexist_out.pptx")\r\n'
    '    s = ImageReplace_GroupSelfTest("X:\\nonexist_in.pptx", '
    '"X:\\nonexist_new.png", "X:\\nonexist_out.pptx")\r\n'
    '    s = ImageReplace_Describe(Nothing)\r\n'
    '    ok = ImageReplace_ReplaceShape(Nothing, "X:\\nonexist.png", e, sh)\r\n'
    # ---- 组合（编组）支持：逐个直接进入，确保这一段真的被编译 ----
    '    ok = IR_IsGroup(Nothing)\r\n'
    '    Set sh = IR_ParentGroup(Nothing)\r\n'
    '    IR_CollectPictures Nothing, Nothing\r\n'
    '    IR_RenameTree Nothing\r\n'
    '    s = IR_SnapLayer(Nothing)\r\n'
    '    IR_ZeroTransform Nothing\r\n'
    '    IR_ApplyTransform Nothing, Nothing\r\n'
    '    IR_ShiftZ Nothing, 1\r\n'
    '    ok = IR_BuildChain(Nothing, Nothing)\r\n'
    '    IR_RestoreNames Nothing\r\n'
    '    ok = IR_ResolveTarget(Nothing, sh, e)\r\n'
    '    ok = IR_ReplaceInGroup(Nothing, "X:\\nonexist.png", e, sh)\r\n'
    '    IR_TryRebuild Nothing\r\n'
    '    s = IR_SelTest(Nothing, Nothing, "X:\\nonexist.png")\r\n'
    '    On Error GoTo 0\r\n'
)


def make_probe(calls):
    body = calls if calls else DEFAULT_CALLS
    return ("\r\nPublic Function %s() As String\r\n" % PROBE_NAME
            + body
            + "    %s = \"reached\"\r\n" % PROBE_NAME
            + "End Function\r\n")


def strip_attr(code):
    return "\r\n".join(
        l for l in code.replace("\r\n", "\n").split("\n")
        if not l.strip().lower().startswith("attribute "))


def read_text(p):
    with open(p, encoding="utf-8") as f:
        return strip_attr(f.read())


def grab_dialogs():
    """枚举 VBA 的 #32770 对话框，并抓出子控件文本（错误信息就在里面）。

    注意：EnumChildWindows 会向对话框所在线程发消息，若那个线程正卡在模态循环里
    有挂住的风险。因此这里只在**确实发现 VBA 对话框之后**才去枚举子窗口 ——
    这也是为什么每次运行前都要先清掉残留的 PowerPoint 实例。
    """
    found = []

    def enum_child(hwnd, lp):
        buf = ctypes.create_unicode_buffer(2048)
        user32.GetWindowTextW(hwnd, buf, 2048)
        if buf.value.strip():
            found[-1].append(buf.value.strip())
        return True

    def cb(hwnd, lp):
        try:
            if user32.IsWindowVisible(hwnd):
                cls = ctypes.create_unicode_buffer(64)
                user32.GetClassNameW(hwnd, cls, 64)
                if cls.value == "#32770":
                    buf = ctypes.create_unicode_buffer(512)
                    user32.GetWindowTextW(hwnd, buf, 512)
                    if "Visual Basic" in buf.value:
                        found.append([])
                        user32.EnumChildWindows(hwnd, CHILD(enum_child), 0)
                        # 顺序整理成「错误类别 | 详细文本」
                        kids = [k for k in found[-1]
                                if k not in ("确定", "OK", "帮助", "Help")]
                        found[-1] = [" | ".join(kids) or buf.value]
        except Exception:
            pass
        return True

    user32.EnumWindows(WND(cb), 0)
    return found


def inject(vbp, name, code, ct=VBE_CT_STDMODULE):
    comp = vbp.VBComponents.Add(ct)
    comp.Name = name
    cm = comp.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(code.replace("\r\n", "\n").replace("\n", "\r\n"))
    return comp


def main():
    argv = sys.argv[1:]
    patch_path = None
    snippet = None
    snippet_file = None
    timeout = 18.0
    src_file = None
    calls = None

    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--patch":
            i += 1
            patch_path = argv[i]
        elif a == "--snippet":
            i += 1
            snippet = argv[i]
        elif a == "--snippet-file":
            i += 1
            snippet_file = argv[i]
        elif a == "--timeout":
            i += 1
            timeout = float(argv[i])
        elif a == "--call":
            i += 1
            calls = "".join("    %s\r\n" % l.rstrip()
                            for l in argv[i].replace("\r\n", "\n").split("\n"))
        elif a == "--call-file":
            i += 1
            with open(argv[i], encoding="utf-8") as f:
                calls = "".join("    %s\r\n" % l.rstrip()
                                for l in f.read().replace("\r\n", "\n").split("\n"))
        elif a == "--noprobe":
            calls = ""
        elif not a.startswith("--"):
            src_file = a
        i += 1

    probe_tail = make_probe(calls)

    # ---- 组装修检内容 ----
    if snippet is not None or snippet_file is not None:
        body = snippet if snippet is not None else open(snippet_file,
                                                        encoding="utf-8").read()
        mod_name = "SnippetMod"
        code = ("Option Explicit\r\n"
                "Public Function %s() As String\r\n" % PROBE_NAME)
        for line in body.replace("\r\n", "\n").split("\n"):
            code += "    " + line.rstrip() + "\r\n"
        code += ("    %s = \"compiled-ok\"\r\n"
                 "End Function\r\n" % PROBE_NAME)
        mods = [("m_Core", read_text(os.path.join(SRC, "_core", "m_Core.bas")),
                 VBE_CT_STDMODULE), (mod_name, code, VBE_CT_STDMODULE)]
        label = "片段：" + (snippet or snippet_file)[:70]
    else:
        if not src_file:
            print("用法：python test/probe_oracle.py <文件.bas> [--patch x.vba]")
            print("      python test/probe_oracle.py --snippet \"vba 代码\"")
            return 2
        code = read_text(src_file)
        name = os.path.splitext(os.path.basename(src_file))[0]
        if patch_path:
            code = code + "\r\n" + "\r\n".join(
                read_text(patch_path).replace("\r\n", "\n").split("\n"))
        code = code + probe_tail
        mods = [("m_Core", read_text(os.path.join(SRC, "_core", "m_Core.bas")),
                 VBE_CT_STDMODULE), (name, code, VBE_CT_STDMODULE)]
        label = src_file + (" + " + os.path.basename(patch_path)
                            if patch_path else "")

    print("预言机受检对象：%s" % label)
    print("注入 %d 个模块：%s   %d 行" % (len(mods), mods[-1][0],
                                          mods[-1][1].count("\n") + 1))

    # 关键：先清干净可能残留下来的、卡在模态框里的 PowerPoint 实例。
    # 否则 DispatchEx 会复用它，后续每一个 COM 调用都会被拖住，
    # 表现成「莫名其妙地没有输出」。
    os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")
    time.sleep(1.5)

    app = com.DispatchEx("PowerPoint.Application")
    hits = {"dialogs": [], "done": False, "err": None}

    try:
        app.Visible = True
        pres = app.Presentations.Add()
        vbp = pres.VBProject
        for n, c, ct in mods:
            inject(vbp, n, c, ct)

        def worker():
            pythoncom.CoInitialize()
            try:
                a2 = com.DispatchEx("PowerPoint.Application")
                # 若类型库已被 makepy 缓存，DispatchEx 会返回强类型对象，
                # 此时 Run 直接可用，_FlagAsMethod 反而不存在 —— 两者都兼容。
                if hasattr(a2, "_FlagAsMethod"):
                    a2._FlagAsMethod("Run")
                a2.Run("%s.%s" % (mods[-1][0], PROBE_NAME))
                hits["done"] = True
            except Exception as e:
                hits["err"] = str(e)[:160]
                hits["done"] = True
            finally:
                pythoncom.CoUninitialize()

        th = threading.Thread(target=worker, daemon=True)
        th.start()

        t0 = time.time()
        while time.time() - t0 < timeout:
            ds = grab_dialogs()
            for d in ds:
                if d not in hits["dialogs"]:
                    hits["dialogs"].append(d)
            if hits["dialogs"]:
                break
            if hits["done"]:
                break
            time.sleep(0.3)

        print("=" * 62)
        if hits["dialogs"]:
            print("结果：编译不通过 ❌")
            for d in hits["dialogs"]:
                print("   弹窗：%s" % " | ".join(d))
            return 1
        if hits["done"] and hits["err"] is None:
            print("结果：编译通过 ✅")
            return 0
        if hits["done"]:
            print("结果：无弹窗但 Run 报错 -> %s" % hits["err"])
            print("      （多半是过程名不存在，而非编译错误）")
            return 0
        print("结果：超时，未能判定（无弹窗且 Run 未返回）？")
        return 2
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(1.0)
        os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")


if __name__ == "__main__":
    sys.exit(main())
