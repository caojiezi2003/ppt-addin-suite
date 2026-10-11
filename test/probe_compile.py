# -*- coding: utf-8 -*-
"""
源码编译体检工具。

把插件源码注入到一个**普通演示文稿**的 VBA 工程里并触发编译。

为什么不在加载项里查：`.ppam` 加载项的 VBA 工程是「隐藏」的，一旦编译失败，
PowerPoint 只会给出一句含糊的「隐藏的模块中的编译错误」，无法定位。
而普通演示文稿的模块是「可见」的，VBA 会指明具体错误、并高亮出错行 ——
我们可以在对话框关掉之后从 VBE 读回光标所在行。

用法：
    python test/probe_compile.py                     # 体检 image-replace
    python test/probe_compile.py image-export        # 体检指定插件目录
"""

import ctypes
import os
import sys
import threading
import time

import win32com.client as com

HERE = os.path.dirname(os.path.abspath(__file__))
SUITE = os.path.dirname(HERE)
WINDOWS = os.path.join(HERE, "probe-windows.txt")
VBE_CT_STDMODULE = 1
BM_CLICK = 0x00F5


def read_src(path):
    with open(path, encoding="utf-8") as f:
        code = f.read()
    # 与 build_suite.py 完全一致的处理：滤除 Attribute 行（组件名由 Name 设定）
    return "\r\n".join(
        l for l in code.replace("\r\n", "\n").split("\n")
        if not l.strip().lower().startswith("attribute "))


def dump_and_dismiss(delay, out_path):
    """延时抓取 VBA 对话框文本，然后点「确定」把它关掉，
    好让被阻塞的 Application.Run 能够返回，我们才拿得到行号。"""
    def run():
        time.sleep(delay)
        try:
            user32 = ctypes.windll.user32
            lines, buttons = [], []
            CHILD = ctypes.WINFUNCTYPE(ctypes.c_bool,
                                       ctypes.c_void_p, ctypes.c_void_p)

            def enum_child(hwnd, lp):
                buf = ctypes.create_unicode_buffer(2048)
                user32.GetWindowTextW(hwnd, buf, 2048)
                cls = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(hwnd, cls, 256)
                if buf.value.strip():
                    lines.append("   [%s] %s" % (cls.value, buf.value))
                    if cls.value == "Button" and buf.value.strip() in ("确定", "OK"):
                        buttons.append(hwnd)
                return True

            def cb(hwnd, lp):
                if user32.IsWindowVisible(hwnd):
                    cls = ctypes.create_unicode_buffer(256)
                    user32.GetClassNameW(hwnd, cls, 256)
                    if cls.value == "#32770":
                        buf = ctypes.create_unicode_buffer(512)
                        user32.GetWindowTextW(hwnd, buf, 512)
                        lines.append("DIALOG #32770 : %s" % buf.value)
                        user32.EnumChildWindows(hwnd, CHILD(enum_child), 0)
                return True

            user32.EnumWindows(
                ctypes.WINFUNCTYPE(ctypes.c_bool,
                                   ctypes.c_void_p, ctypes.c_void_p)(cb), 0)
            with open(out_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines) if lines else "(未发现对话框)")
            for h in buttons:
                user32.SendMessageW(h, BM_CLICK, 0, 0)
        except Exception as e:
            with open(out_path, "w", encoding="utf-8") as f:
                f.write("dump failed: %r" % e)
    threading.Thread(target=run, daemon=True).start()


def inject(proj, name, code):
    comp = proj.VBComponents.Add(VBE_CT_STDMODULE)
    comp.Name = name
    cm = comp.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(code.replace("\r\n", "\n").replace("\n", "\r\n"))
    return comp


def main():
    which = sys.argv[1] if len(sys.argv) > 1 else "image-replace"
    d = os.path.join(SUITE, "src", which)
    if not os.path.isdir(d):
        print("找不到插件目录：%s" % d)
        return 2

    core = read_src(os.path.join(SUITE, "src", "_core", "m_Core.bas"))
    mods = []
    for f in sorted(os.listdir(d)):
        if f.lower().endswith(".bas"):
            mods.append((os.path.splitext(f)[0], read_src(os.path.join(d, f))))
    if not mods:
        print("该插件目录下没有 .bas 文件")
        return 2

    # 给第一个模块追加一个无参探针，用来触发编译
    name0, code0 = mods[0]

    # 可选第二参数：追加「自检注入」代码，复现构建时的拼接顺序
    if len(sys.argv) > 2:
        extra = os.path.join(HERE, sys.argv[2])
        if os.path.exists(extra):
            with open(extra, encoding="utf-8") as f:
                patch = f.read()
            code0 = code0 + "\r\n" + "\r\n".join(
                patch.replace("\r\n", "\n").split("\n"))
            print("已追加注入代码：%s" % sys.argv[2])
        else:
            print("注入代码文件不存在：%s" % extra)

    code0 += ("\r\nPublic Function ProbeRun() As String\r\n"
              "    ProbeRun = \"compiled-ok\"\r\n"
              "End Function\r\n")
    mods[0] = (name0, code0)

    print("体检插件：%s  ->  模块 %s" % (which, [m[0] for m in mods]))

    app = com.DispatchEx("PowerPoint.Application")
    try:
        app.Visible = True
        pres = app.Presentations.Add()
        vbp = pres.VBProject
        inject(vbp, "mCore", core)
        for n, c in mods:
            inject(vbp, n, c)
        print("已注入 %d 个模块" % (len(mods) + 1))

        dump_and_dismiss(8, WINDOWS)
        app._FlagAsMethod("Run")
        ok = False
        try:
            r = app.Run("%s.ProbeRun" % name0)
            print("COMPILE: 通过  ->  %s" % r)
            ok = True
        except Exception as e:
            print("COMPILE: 失败  ->  %s" % e)
            try:
                cp = app.VBE.ActiveCodePane
                cm = cp.CodeModule
                ln, col = cp.GetSelection()
                print("VBE 光标定位：第 %d 行  共 %d 行" % (ln, cm.CountOfLines))
                lo = max(1, ln - 3)
                for k in range(lo, min(ln + 3, cm.CountOfLines) + 1):
                    mark = " <<< 出错行" if k == ln else ""
                    print("   %4d | %s%s" % (k, cm.Lines(k, 1), mark))
            except Exception as e2:
                print("读取 VBE 失败：%s" % e2)
        return 0 if ok else 1
    finally:
        try:
            app.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(main())
