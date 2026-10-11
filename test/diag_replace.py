# -*- coding: utf-8 -*-
"""
「图片替换」专项诊断：只跑插件③，快速定位卡点。

复用 run_suite_test 里的样张构造代码，但跳过图片导出与选页另存，
直接把替换过程的逐步日志（ir-steps.txt / ir-trace.log）打出来。
"""

import ctypes
import os
import sys
import threading
import time

import win32com.client as com

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import run_suite_test as R          # noqa: E402

STEPS = os.path.join(HERE, "ir-steps.txt")
TRACE = os.path.join(HERE, "ir-trace.log")
WINDOWS = os.path.join(HERE, "ir-windows.txt")


def dump_windows_after(delay, out_path):
    """延时枚举顶层窗口，并把 #32770 对话框的子控件文本一并抓出来
    —— VBA 的「编译错误 / 运行时错误」对话框正是这个窗口类。"""
    def run():
        time.sleep(delay)
        try:
            user32 = ctypes.windll.user32
            lines = []
            CHILD = ctypes.WINFUNCTYPE(ctypes.c_bool,
                                       ctypes.c_void_p, ctypes.c_void_p)

            def enum_child(hwnd, lp):
                buf = ctypes.create_unicode_buffer(1024)
                user32.GetWindowTextW(hwnd, buf, 1024)
                cls = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(hwnd, cls, 256)
                if buf.value.strip():
                    lines.append("      [child] %-18s | %s"
                                 % (cls.value, buf.value))
                return True

            def cb(hwnd, lp):
                if user32.IsWindowVisible(hwnd):
                    buf = ctypes.create_unicode_buffer(512)
                    user32.GetWindowTextW(hwnd, buf, 512)
                    cls = ctypes.create_unicode_buffer(256)
                    user32.GetClassNameW(hwnd, cls, 256)
                    if cls.value == "#32770":
                        lines.append("== DIALOG #32770 | %s" % buf.value)
                        user32.EnumChildWindows(hwnd, CHILD(enum_child), 0)
                    elif buf.value:
                        lines.append("%-24s | %s" % (cls.value, buf.value))
                return True

            user32.EnumWindows(
                ctypes.WINFUNCTYPE(ctypes.c_bool,
                                   ctypes.c_void_p, ctypes.c_void_p)(cb), 0)
            with open(out_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
        except Exception as e:
            with open(out_path, "w", encoding="utf-8") as f:
                f.write("dump failed: %r" % e)
    threading.Thread(target=run, daemon=True).start()


def main():
    os.makedirs(R.OUTDIR, exist_ok=True)
    for f in (R.REPLACE_IN, R.REPLACE_OUT, R.IMG_OLD, R.IMG_NEW, STEPS, TRACE):
        try:
            os.remove(f)
        except OSError:
            pass

    R.write_png(R.IMG_OLD, 200, 140, (192, 80, 77))
    R.write_png(R.IMG_NEW, 300, 200, (79, 129, 189))

    if not os.path.exists(R.ADDIN):
        print("缺少产物：%s" % R.ADDIN)
        return 2

    app = com.DispatchEx("PowerPoint.Application")
    try:
        app.Visible = True
        added = app.AddIns.Add(R.ADDIN)
        added.Loaded = True
        print("ADDIN: loaded=%s" % added.Loaded)

        R.build_replace_deck(app, R.REPLACE_IN, R.IMG_OLD)
        print("样张已生成：%s" % os.path.basename(R.REPLACE_IN))

        app._FlagAsMethod("Run")
        refs = [
            "'%s.ppam'!m_ImageReplace.ImageReplace_SelfTestAuto" % R.PROJ,
            "'%s'!m_ImageReplace.ImageReplace_SelfTestAuto" % R.PROJ,
            "m_ImageReplace.ImageReplace_SelfTestAuto",
        ]
        dump_windows_after(10, WINDOWS)
        done = False
        for r in refs:
            try:
                res = app.Run(r)
                print("RUN  : 成功（%s）" % r)
                print("RESULT: %s" % res)
                done = True
                break
            except Exception as e:
                print("RUN  : 失败（%s）-> %s" % (r, e))
        if not done:
            print("RUN_ERROR: 所有引用形式均失败")
    finally:
        try:
            app.Quit()
        except Exception:
            pass

    for name, p in (("ir-steps.txt", STEPS), ("ir-trace.log", TRACE),
                    ("ir-windows.txt", WINDOWS)):
        print("=" * 62)
        print(name)
        print("=" * 62)
        if os.path.exists(p):
            print(open(p, encoding="utf-8", errors="replace").read().strip())
        else:
            print("(不存在)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
