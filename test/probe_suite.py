# -*- coding: utf-8 -*-
"""
整套套件编译体检。

完全复现 build_suite.py 的工程内容（含窗体、含自检注入），
但注入到**普通演示文稿**而非 .ppam —— 这样模块是「可见」的，
编译错误能定位到具体行，而不是加载项那句含糊的「隐藏的模块中的编译错误」。

用法：
    python test/probe_suite.py
"""

import ctypes
import os
import sys
import threading
import time

import win32com.client as com

HERE = os.path.dirname(os.path.abspath(__file__))
SUITE = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(SUITE, "build"))
import build_suite as bs           # noqa: E402

WINDOWS = os.path.join(HERE, "probe-windows.txt")
VBE_CT_STDMODULE = 1
VBE_CT_MSFORM = 3
BM_CLICK = 0x00F5

PROBE = ("\r\nPublic Function ProbeCompile() As String\r\n"
         "    Dim s As String\r\n"
         "    s = ImageReplace_SelfTest(\"X:\\nonexist_in.pptx\", \"X:\\nonexist.png\","
         " \"X:\\nonexist_out.pptx\")\r\n"
         "    ProbeCompile = \"reached: \" & s\r\n"
         "End Function\r\n")


def read_text(p):
    with open(p, encoding="utf-8") as f:
        return f.read()


def strip_attr(code):
    return "\r\n".join(
        l for l in code.replace("\r\n", "\n").split("\n")
        if not l.strip().lower().startswith("attribute "))


def dump_and_dismiss(delay, out_path, rounds=14, interval=2.0):
    """反复抓取并点掉 VBA 对话框。

    编译错误可能不止一个，点掉一个之后 VBA 会继续编译并再弹，
    因此必须循环点击，被阻塞的 Application.Run 才能最终返回。
    """
    def run():
        time.sleep(delay)
        try:
            user32 = ctypes.windll.user32
            lines = []
            CHILD = ctypes.WINFUNCTYPE(ctypes.c_bool,
                                       ctypes.c_void_p, ctypes.c_void_p)
            WND = ctypes.WINFUNCTYPE(ctypes.c_bool,
                                     ctypes.c_void_p, ctypes.c_void_p)
            for _ in range(rounds):
                buttons = []

                def enum_child(hwnd, lp):
                    buf = ctypes.create_unicode_buffer(2048)
                    user32.GetWindowTextW(hwnd, buf, 2048)
                    cls = ctypes.create_unicode_buffer(256)
                    user32.GetClassNameW(hwnd, cls, 256)
                    if buf.value.strip():
                        lines.append("   [%s] %s" % (cls.value, buf.value))
                        if cls.value == "Button" and \
                                buf.value.strip() in ("确定", "OK"):
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

                user32.EnumWindows(WND(cb), 0)
                with open(out_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(lines) if lines else "(未发现对话框)")
                for h in buttons:
                    user32.SendMessageW(h, BM_CLICK, 0, 0)
                time.sleep(interval)
        except Exception as e:
            with open(out_path, "w", encoding="utf-8") as f:
                f.write("dump failed: %r" % e)
    threading.Thread(target=run, daemon=True).start()


def inject(vbp, name, code, ct):
    comp = vbp.VBComponents.Add(ct)
    comp.Name = name
    cm = comp.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(code.replace("\r\n", "\n").replace("\n", "\r\n"))
    return comp


def main():
    plugins = bs.load_plugins()

    # 复现 --selftest 构建：注入各模块的自检代码
    patches = {
        "frmImageExport": read_text(os.path.join(HERE, "selftest.vba")),
        "m_SlideExport": read_text(os.path.join(HERE, "selftest_slide.vba")),
        "m_ImageReplace": read_text(os.path.join(HERE, "selftest_replace.vba"))
                          + PROBE,
    }
    plugins.append({
        "id": "_test", "order": 9999, "title": "自检驱动", "group": None,
        "buttons": [],
        "components": [{"type": "module", "name": "modSelfTest",
                        "file": "modSelfTest.bas"}],
        "_dir": HERE, "_dirname": "_test",
    })

    app = com.DispatchEx("PowerPoint.Application")
    try:
        app.Visible = True
        pres = app.Presentations.Add()
        vbp = pres.VBProject

        for p in plugins:
            for comp in p["components"]:
                path = os.path.join(p["_dir"], comp["file"])
                code = strip_attr(read_text(path))
                patch = patches.get(comp["name"])
                if patch:
                    code = code + "\r\n" + \
                        "\r\n".join(patch.replace("\r\n", "\n").split("\n"))
                ct = VBE_CT_STDMODULE if comp["type"] == "module" else VBE_CT_MSFORM
                inject(vbp, comp["name"], code, ct)
                print("  注入 %-18s (%s)  %4d 行"
                      % (comp["name"], comp["type"], code.count("\n") + 1))

        dump_and_dismiss(8, WINDOWS)
        app._FlagAsMethod("Run")
        ok = False
        try:
            r = app.Run("m_ImageReplace.ProbeCompile")
            print("COMPILE: 通过  ->  %s" % r)
            ok = True
        except Exception as e:
            print("COMPILE: 失败  ->  %s" % e)
            try:
                cp = app.VBE.ActiveCodePane
                cm = cp.CodeModule
                ln, col = cp.GetSelection()
                print("VBE 光标定位：第 %d 行 / 共 %d 行" % (ln, cm.CountOfLines))
                for k in range(max(1, ln - 4), min(ln + 4, cm.CountOfLines) + 1):
                    mark = "  <<< 出错行" if k == ln else ""
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
