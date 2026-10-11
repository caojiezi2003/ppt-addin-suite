# -*- coding: utf-8 -*-
"""
VBA 编译错误定位器（逐行级）。

痛点：
    `.ppam` 加载项的工程是「隐藏」的 —— 编译失败只给一句
    「隐藏的模块中的编译错误: xxx」，无行号无标识符。
    而普通演示文稿的模块是「可见」的，VBA 会弹「编译错误: 变量未定义」
    并把光标停在出错行 —— 但那个对话框是**模态**的，会把 Application.Run 卡死，
    于是「点掉对话框」和「等待 Run 返回」必须分处两个线程。

本工具的解法：
    主线程  —— 不断枚举 VBA 的 #32770 对话框、抓文本、PostMessage 点「确定」
    工作线程 —— 另起一个 COM 连接去调 Application.Run，被模态框卡住也没关系
    Run 返回后 —— 从 VBE 的 ActiveCodePane.GetSelection() 读回出错行号并打印上下文

用法：
    python test/probe_vbe.py --sanity
    python test/probe_vbe.py image-replace
    python test/probe_vbe.py image-replace selftest_replace.vba
    python test/probe_vbe.py all
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
VBE_CT_MSFORM = 3
BM_CLICK = 0x00F5

user32 = ctypes.windll.user32
CHILD = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
WND = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

PROBE_TAIL = ("\r\nPublic Function ProbeRun() As String\r\n"
              "    ProbeRun = \"compiled-ok\"\r\n"
              "End Function\r\n")


def strip_attr(code):
    return "\r\n".join(
        l for l in code.replace("\r\n", "\n").split("\n")
        if not l.strip().lower().startswith("attribute "))


def read_src(path):
    with open(path, encoding="utf-8") as f:
        return strip_attr(f.read())


# ------------------------------------------------------------------ 对话框

stop_flag = threading.Event()
seen_dialogs = []
dialog_hits = {"n": 0}

WM_COMMAND = 0x0111
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_CLOSE = 0x0010
VK_RETURN = 0x0D
IDOK = 1


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]


class _INPUTunion(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("u", _INPUTunion)]


def press_enter():
    """向当前前台窗口发一次回车 —— 对话框的默认按钮就是「确定」"""
    down = INPUT(type=1, u=_INPUTunion(
        ki=KEYBDINPUT(wVk=VK_RETURN, wScan=0, dwFlags=0, time=0,
                      dwExtraInfo=None)))
    up = INPUT(type=1, u=_INPUTunion(
        ki=KEYBDINPUT(wVk=VK_RETURN, wScan=0, dwFlags=2, time=0,
                      dwExtraInfo=None)))
    user32.SendInput(1, ctypes.byref(down), ctypes.sizeof(INPUT))
    time.sleep(0.05)
    user32.SendInput(1, ctypes.byref(up), ctypes.sizeof(INPUT))


def dismiss(d):
    """多管齐下：BM_CLICK / WM_COMMAND(IDOK) / 回车键 / 关窗"""
    user32.SetForegroundWindow(d["hwnd"])
    time.sleep(0.15)
    for b in d["buttons"]:
        user32.PostMessageW(b, BM_CLICK, 0, 0)
    user32.PostMessageW(d["hwnd"], WM_COMMAND, IDOK, 0)
    press_enter()
    time.sleep(0.15)
    if user32.IsWindow(d["hwnd"]):
        user32.PostMessageW(d["hwnd"], WM_CLOSE, 0, 0)


def collect_dialogs():
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


def watcher(timeout=60.0):
    """主线程里的观察者：抓对话框文本 + 点掉"""
    t0 = time.time()
    reported = set()
    while time.time() - t0 < timeout:
        for d in collect_dialogs():
            dialog_hits["n"] += 1
            seen_dialogs.append(d)
            key = tuple(d["kids"])
            if key not in reported:
                reported.add(key)
                print("  [对话框] %s" % (" | ".join(d["kids"])))
            dismiss(d)
        if stop_flag.is_set():
            break
        time.sleep(0.4)
    time.sleep(0.8)
    for d in collect_dialogs():
        dialog_hits["n"] += 1
        seen_dialogs.append(d)
        dismiss(d)


# ------------------------------------------------------------------ 注入

def inject(vbp, name, code, ct=VBE_CT_STDMODULE):
    comp = vbp.VBComponents.Add(ct)
    comp.Name = name
    cm = comp.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(code.replace("\r\n", "\n").replace("\n", "\r\n"))
    return comp


def resolve_targets(arg, extra):
    """返回 [(组件名, 代码, 类型)]；第一个标准模块会追加 ProbeRun"""
    mods = [("m_Core", read_src(os.path.join(SRC, "_core", "m_Core.bas")),
             VBE_CT_STDMODULE)]
    pids = (["image-export", "slide-export", "image-replace"]
            if arg == "all" else [arg])
    for pid in pids:
        d = os.path.join(SRC, pid)
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            if f.lower().endswith(".bas"):
                mods.append((os.path.splitext(f)[0], read_src(os.path.join(d, f)),
                             VBE_CT_STDMODULE))
            elif f.lower().endswith(".code.vba"):
                mods.append((f[:-len(".code.vba")], read_src(os.path.join(d, f)),
                             VBE_CT_MSFORM))
    if arg == "all":
        mods.append(("modSelfTest", read_src(os.path.join(HERE, "modSelfTest.bas")),
                     VBE_CT_STDMODULE))

    if extra:
        p = os.path.join(HERE, extra)
        if os.path.exists(p):
            with open(p, encoding="utf-8") as f:
                patch = f.read()
            nm, code, ct = mods[-1]
            mods[-1] = (nm, code + "\r\n" +
                        "\r\n".join(patch.replace("\r\n", "\n").split("\n")), ct)
            print("  已追加注入代码：%s -> %s" % (extra, nm))
        else:
            print("  注入代码文件不存在：%s" % p)
    return mods


def main():
    argv = sys.argv[1:]
    flags = [a for a in argv if a.startswith("--")]
    args = [a for a in argv if not a.startswith("--")]
    sanity = "--sanity" in flags

    if sanity:
        mods = [("mCore",
                 "Option Explicit\r\nPublic Function ProbeRun() As String\r\n"
                 "    Dim q As Long\r\n"
                 "    q = deliberately_undeclared_thing\r\n"
                 "    ProbeRun = \"compiled-ok\"\r\n"
                 "End Function\r\n", VBE_CT_STDMODULE)]
        target = "mCore"
    else:
        arg = args[0] if args else "image-replace"
        extra = args[1] if len(args) > 1 else None
        mods = resolve_targets(arg, extra)
        # 找到第一个标准模块作为探针宿主
        target = None
        for i, (n, c, ct) in enumerate(mods):
            if ct == VBE_CT_STDMODULE:
                target = n
                mods[i] = (n, c + PROBE_TAIL, ct)
                break

    print("目标过程：%s.ProbeRun" % target)
    print("将注入 %d 个组件：%s" % (len(mods), [m[0] for m in mods]))

    app = com.DispatchEx("PowerPoint.Application")
    result = {"done": False, "err": None}

    try:
        app.Visible = True
        pres = app.Presentations.Add()
        vbp = pres.VBProject
        for name, code, ct in mods:
            inject(vbp, name, code, ct)
            print("  注入 %-20s %4d 行" % (name, code.count("\n") + 1))

        # ---- 工作线程：另起 COM 连接去 Run（会被模态框卡住）----
        def worker():
            pythoncom.CoInitialize()
            try:
                app2 = com.DispatchEx("PowerPoint.Application")
                app2._FlagAsMethod("Run")
                app2.Run("%s.ProbeRun" % target)
                result["done"] = True
            except Exception as e:
                result["err"] = str(e)[:200]
                result["done"] = True
            finally:
                pythoncom.CoUninitialize()

        th = threading.Thread(target=worker, daemon=True)
        th.start()

        # ---- 主线程：抓对话框并点掉 ----
        watcher(timeout=45.0)
        stop_flag.set()
        th.join(timeout=10)

        print("=" * 64)
        print("对话框捕获次数：%d" % dialog_hits["n"])
        if result["done"] and result["err"] is None:
            print("编译通过 ✅  ->  ProbeRun 正常返回")
            return 0
        print("编译失败 ❌  ->  %s" % (result["err"] or "(线程未返回)"))

        # ---- 读回出错行 ----
        try:
            cp = app.VBE.ActiveCodePane
            cm = cp.CodeModule
            sel = cp.GetSelection()          # (startLine, startCol, endLine, endCol)
            ln = int(sel[0])
            print("出错位置：%s  第 %d 行（共 %d 行）"
                  % (cm.Name, ln, cm.CountOfLines))
            for k in range(max(1, ln - 4), min(ln + 4, cm.CountOfLines) + 1):
                mark = "   <<< 出错行" if k == ln else ""
                print("   %5d | %s%s" % (k, cm.Lines(k, 1), mark))
        except Exception as e:
            print("无法从 VBE 读回光标：%s" % e)
        return 1
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(1.5)
        os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")


if __name__ == "__main__":
    sys.exit(main())
