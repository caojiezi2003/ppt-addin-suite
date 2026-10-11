# -*- coding: utf-8 -*-
"""
非阻塞式编译体检。

核心观察：
    如果「编译错误」对话框在模块**注入时**就弹出来（而不必等 Application.Run），
    那么定位问题就不需要调用 Run —— 也就不存在被模态对话框卡死的问题。

因此本工具：
    注入模块 -> 只等几秒 -> 枚举 VBA 的 #32770 对话框抓文本 -> 用 PostMessage 点掉 -> 退出。
    全程不调 Application.Run，永远不会阻塞。

用法：
    python test/probe_inject.py --sanity          # 自检：故意注入坏模块，验证工具灵敏
    python test/probe_inject.py image-export
    python test/probe_inject.py image-replace
    python test/probe_inject.py all               # 全部插件 + 自检驱动
    python test/probe_inject.py <文件> [更多文件]  # 直接指定 .bas 文件
"""

import ctypes
import os
import sys
import time

import win32com.client as com

HERE = os.path.dirname(os.path.abspath(__file__))
SUITE = os.path.dirname(HERE)
SRC = os.path.join(SUITE, "src")

VBE_CT_STDMODULE = 1
VBE_CT_MSFORM = 3
WM_CLOSE = 0x0010
BM_CLICK = 0x00F5

user32 = ctypes.windll.user32
CHILD = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
WND = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)


def strip_attr(code):
    return "\r\n".join(
        l for l in code.replace("\r\n", "\n").split("\n")
        if not l.strip().lower().startswith("attribute "))


def read_src(path):
    with open(path, encoding="utf-8") as f:
        return strip_attr(f.read())


def collect_dialogs():
    """返回 [(标题, [子控件文本...], [按钮hwnd...])]"""
    found = []

    def enum_child(hwnd, lp):
        buf = ctypes.create_unicode_buffer(2048)
        user32.GetWindowTextW(hwnd, buf, 2048)
        cls = ctypes.create_unicode_buffer(64)
        user32.GetClassNameW(hwnd, cls, 64)
        if buf.value.strip():
            found[-1]["kids"].append(buf.value.strip())
            if cls.value == "Button" and buf.value.strip() in ("确定", "OK"):
                found[-1]["buttons"].append(hwnd)
        return True

    def cb(hwnd, lp):
        if user32.IsWindowVisible(hwnd):
            cls = ctypes.create_unicode_buffer(64)
            user32.GetClassNameW(hwnd, cls, 64)
            if cls.value == "#32770":
                buf = ctypes.create_unicode_buffer(512)
                user32.GetWindowTextW(hwnd, buf, 512)
                if "Visual Basic" in buf.value:
                    found.append({"hwnd": hwnd, "title": buf.value,
                                  "kids": [], "buttons": []})
                    user32.EnumChildWindows(hwnd, CHILD(enum_child), 0)
        return True

    user32.EnumWindows(WND(cb), 0)
    return found


def dismiss(dialogs):
    for d in dialogs:
        for b in d["buttons"]:
            user32.PostMessageW(b, BM_CLICK, 0, 0)
        if not d["buttons"]:
            user32.PostMessageW(d["hwnd"], WM_CLOSE, 0, 0)


def inject(vbp, name, code, ct=VBE_CT_STDMODULE):
    comp = vbp.VBComponents.Add(ct)
    comp.Name = name
    cm = comp.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(code.replace("\r\n", "\n").replace("\n", "\r\n"))
    return comp


def resolve_targets(arg):
    """返回 [(组件名, 代码, 组件类型)]"""
    if arg == "all":
        mods = []
        mods.append(("m_Core", read_src(os.path.join(SRC, "_core",
                                                     "m_Core.bas")),
                     VBE_CT_STDMODULE))
        for pid in ("image-export", "slide-export", "image-replace"):
            d = os.path.join(SRC, pid)
            for f in sorted(os.listdir(d)):
                if f.lower().endswith(".bas"):
                    mods.append((os.path.splitext(f)[0],
                                 read_src(os.path.join(d, f)),
                                 VBE_CT_STDMODULE))
                elif f.lower().endswith(".code.vba"):
                    mods.append((f[:-len(".code.vba")],
                                 read_src(os.path.join(d, f)),
                                 VBE_CT_MSFORM))
        mods.append(("modSelfTest", read_src(os.path.join(HERE,
                                                          "modSelfTest.bas")),
                     VBE_CT_STDMODULE))
        return mods

    if os.path.isdir(os.path.join(SRC, arg)):
        d = os.path.join(SRC, arg)
        mods = [("m_Core", read_src(os.path.join(SRC, "_core", "m_Core.bas")),
                 VBE_CT_STDMODULE)]
        for f in sorted(os.listdir(d)):
            if f.lower().endswith(".bas"):
                mods.append((os.path.splitext(f)[0],
                             read_src(os.path.join(d, f)), VBE_CT_STDMODULE))
            elif f.lower().endswith(".code.vba"):
                mods.append((f[:-len(".code.vba")],
                             read_src(os.path.join(d, f)), VBE_CT_MSFORM))
        return mods

    mods = []
    for p in sys.argv[1:]:
        if os.path.isfile(p):
            mods.append((os.path.splitext(os.path.basename(p))[0],
                         read_src(p), VBE_CT_STDMODULE))
    return mods


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    arg = args[0] if args else "image-replace"

    if "--sanity" in sys.argv:
        mods = [("mCore", "Option Explicit\r\n"
                          "Public Sub Touch()\r\n"
                          "    Dim q As Long\r\n"
                          "    q = deliberately_undeclared_thing\r\n"
                          "End Sub\r\n", VBE_CT_STDMODULE)]
        print("自检模式：注入一个含未声明变量的模块，应当能抓到「变量未定义」")
    else:
        mods = resolve_targets(arg)
        if not mods:
            print("没有可注入的模块：%s" % arg)
            return 2

    print("将注入 %d 个组件：%s" % (len(mods), [m[0] for m in mods]))

    app = com.DispatchEx("PowerPoint.Application")
    try:
        app.Visible = True
        pres = app.Presentations.Add()
        vbp = pres.VBProject
        for name, code, ct in mods:
            inject(vbp, name, code, ct)
            print("  注入 %-20s %4d 行" % (name, code.count("\n") + 1))

        # 不再调用 Application.Run —— 只静静等待对话框自己冒出来
        all_dialogs = []
        for _ in range(12):
            time.sleep(1.0)
            ds = collect_dialogs()
            if ds:
                for d in ds:
                    print("  对话框：「%s」-> %s"
                          % (" / ".join(d["kids"]) or d["title"], "已点掉"))
                all_dialogs.extend(ds)
                dismiss(ds)
        time.sleep(1.0)
        ds = collect_dialogs()
        all_dialogs.extend(ds)
        dismiss(ds)

        print("=" * 62)
        if all_dialogs:
            print("发现 %d 个 VBA 对话框 —— 编译未通过 ❌" % len(all_dialogs))
            seen = set()
            for d in all_dialogs:
                key = tuple(d["kids"])
                if key in seen:
                    continue
                seen.add(key)
                print("   · %s" % " | ".join(d["kids"]))
            return 1
        print("全程未出现 VBA 对话框 ✅（注入阶段不报编译错误）")
        return 0
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(1.5)
        os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")


if __name__ == "__main__":
    sys.exit(main())
