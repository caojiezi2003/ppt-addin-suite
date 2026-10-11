# -*- coding: utf-8 -*-
"""为《使用说明》截取真实界面截图。

用法：
    python _doc_shots.py setup          # 启动 PPT + 打开示例文稿（保留进程）
    python _doc_shots.py shot-ribbon    # 功能区总览
    python _doc_shots.py shot-export    # 导出图片窗体
    python _doc_shots.py shot-slidesel  # 缩略图多选
    python _doc_shots.py shot-slidesave # 另存对话框
    python _doc_shots.py shot-selimg    # 选中图片
    python _doc_shots.py shot-replace   # 选择新图对话框
    python _doc_shots.py teardown       # 关闭 PPT
"""
import ctypes
import os
import sys
import time

ctypes.windll.shcore.SetProcessDpiAwareness(2)

import win32con
import win32gui
import win32ui
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "output")
SHOTS = os.path.abspath(os.path.join(HERE, "..", "docs", "screenshots"))
DECK = os.path.join(OUT, "_demo_deck.pptx")
os.makedirs(SHOTS, exist_ok=True)

WM_CLOSE = 0x0010


# ---------------------------------------------------------------- 基础工具
def capture(hwnd, path):
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = r - l, b - t
    if w <= 0 or h <= 0:
        raise RuntimeError("bad rect %r" % ((l, t, r, b),))
    hdc = win32gui.GetWindowDC(hwnd)
    mfc = win32ui.CreateDCFromHandle(hdc)
    save = mfc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfc, w, h)
    save.SelectObject(bmp)
    ok = ctypes.windll.user32.PrintWindow(hwnd, save.GetSafeHdc(), 2)
    info = bmp.GetInfo()
    img = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]),
                           bmp.GetBitmapBits(True), "raw", "BGRX", 0, 1)
    win32gui.DeleteObject(bmp.GetHandle())
    save.DeleteDC()
    mfc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hdc)
    img.save(path)
    return ok, (w, h)


def enum_visible():
    wins = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h):
            wins.append(h)
        return True

    win32gui.EnumWindows(cb, None)
    return wins


def find(pred, timeout=12.0):
    end = time.time() + timeout
    while time.time() < end:
        for h in enum_visible():
            if pred(h):
                return h
        time.sleep(0.2)
    return None


def by_title(substr):
    return find(lambda h: substr in win32gui.GetWindowText(h))


def by_class(cls, title_part=None):
    def pred(h):
        if win32gui.GetClassName(h) != cls:
            return False
        return title_part is None or title_part in win32gui.GetWindowText(h)
    return find(pred)


def close_dialog(hwnd):
    win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)


# ---------------------------------------------------------------- PowerPoint
def ppt_hwnd():
    return by_title(" - PowerPoint")


def get_app():
    import win32com.client as com
    return com.GetActiveObject("PowerPoint.Application")


def get_pres(app):
    return app.Presentations(1)


def setup():
    os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")
    time.sleep(1.5)
    import win32com.client as com
    app = com.DispatchEx("PowerPoint.Application")
    app.Visible = True
    app.WindowState = 1  # 最大化，保证 Ribbon 布局稳定
    pres = app.Presentations.Open(DECK, False, False, True)
    time.sleep(1.0)
    hwnd = ppt_hwnd()
    print("ppt hwnd:", hwnd, win32gui.GetWindowRect(hwnd))


def teardown():
    os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")


# ---------------------------------------------------------------- 场景
def activate_tab():
    """点击「增强插件」标签。"""
    _click_at(ppt_hwnd(), 219, 62)
    time.sleep(0.8)


def _click_at(hwnd, x, y):
    import win32api
    l, t, _, _ = win32gui.GetWindowRect(hwnd)
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(0.3)
    win32api.SetCursorPos((l + x, t + y))
    time.sleep(0.15)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.06)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def shot_ribbon():
    activate_tab()
    hwnd = ppt_hwnd()
    capture(hwnd, os.path.join(SHOTS, "_raw_ribbon_full.png"))
    print("ok _raw_ribbon_full.png")


def shot_export():
    """点「导出为图片」大按钮 → 抓模态窗体 → 关闭。"""
    activate_tab()
    hwnd = ppt_hwnd()
    # 大按钮：第一个分组内，约 (60, 165)
    _click_at(hwnd, 22, 115)
    form = by_class("ThunderDFrame")
    if not form:
        print("!! 窗体未出现")
        return
    time.sleep(1.2)  # 等界面完全绘制
    capture(form, os.path.join(SHOTS, "_raw_export_form.png"))
    print("ok _raw_export_form.png", win32gui.GetWindowRect(form))
    close_dialog(form)
    time.sleep(0.8)


def shot_slidesel():
    """缩略图窗格选中第 2、4 页。"""
    app = get_app()
    pres = get_pres(app)
    # 先激活缩略图窗格，再选中第 2、4 页
    app.ActiveWindow.Panes(1).Activate()
    time.sleep(0.3)
    pres.Slides.Range(ArrayOf(2, 4)).Select()
    time.sleep(0.5)
    hwnd = ppt_hwnd()
    capture(hwnd, os.path.join(SHOTS, "_raw_slidesel.png"))
    print("ok _raw_slidesel.png")


def ArrayOf(*idx):
    import win32com.client
    return win32com.client.VARIANT(pythoncom_vt_array_int(), list(idx))


def pythoncom_vt_array_int():
    import pythoncom
    return pythoncom.VT_ARRAY | pythoncom.VT_I4


def shot_slidesave():
    app = get_app()
    pres = get_pres(app)
    app.ActiveWindow.Panes(1).Activate()
    time.sleep(0.3)
    pres.Slides.Range(ArrayOf(2, 4)).Select()
    time.sleep(0.4)
    hwnd = ppt_hwnd()
    # 缩略图窗格获得焦点后点功能区按钮
    _click_at(hwnd, 100, 115)  # 「导出选中幻灯片」大按钮
    dlg = by_title("张幻灯片另存为")
    if not dlg:
        print("!! 另存对话框未出现")
        return
    time.sleep(1.0)
    capture(dlg, os.path.join(SHOTS, "_raw_slidesave.png"))
    print("ok _raw_slidesave.png")
    close_dialog(dlg)
    time.sleep(0.6)


def shot_selimg():
    """进入第 2 页并选中图片形状。"""
    app = get_app()
    pres = get_pres(app)
    pres.Windows(1).View.GotoSlide(2)
    time.sleep(0.5)
    for shp in pres.Slides(2).Shapes:
        if shp.Type == 13:  # msoPicture
            shp.Select()
            break
    time.sleep(0.5)
    hwnd = ppt_hwnd()
    capture(hwnd, os.path.join(SHOTS, "_raw_selimg.png"))
    print("ok _raw_selimg.png")


def shot_replace():
    _click_at(ppt_hwnd(), 182, 115)  # 「原位替换图片」大按钮
    dlg = by_title("选择用于替换的新图片")
    if not dlg:
        print("!! 替换对话框未出现")
        return
    time.sleep(1.0)
    capture(dlg, os.path.join(SHOTS, "_raw_replace_pick.png"))
    print("ok _raw_replace_pick.png")
    close_dialog(dlg)
    time.sleep(0.6)




# ---------------------------------------------------------------- 替换完整流程
def _child_controls(hwnd, out=None):
    out = out if out is not None else []
    def cb(h, _):
        out.append(h)
        return True
    win32gui.EnumChildWindows(hwnd, cb, None)
    return out


def _set_filename(dlg, path):
    """在打开文件对话框的文件名框里填入完整路径。"""
    for h in _child_controls(dlg):
        if win32gui.GetClassName(h) in ("Edit", "ComboBoxEx32"):
            txt = win32gui.GetWindowText(h)
            # 文件名编辑框通常是空的或只有占位；有路径的那就直接改
            if win32gui.GetClassName(h) == "Edit":
                win32gui.SendMessage(h, 0x000C, 0, path)  # WM_SETTEXT
                return True
    return False


def _click_ok(dlg):
    import win32con
    for h in _child_controls(dlg):
        if win32gui.GetClassName(h) == "Button":
            if win32gui.GetDlgCtrlID(h) == 1:  # IDOK
                win32gui.SendMessage(h, 0x00F5, 0, 0)  # BM_CLICK
                return True
    return False


def replace_run():
    """完整走一遍真实替换：选中图片 -> 点按钮 -> 选新图 -> 打开 -> 抓完成提示。"""
    import win32con
    app = get_app()
    pres = get_pres(app)
    pres.Windows(1).View.GotoSlide(2)
    time.sleep(0.4)
    for shp in pres.Slides(2).Shapes:
        if shp.Type == 13:
            shp.Select()
            break
    time.sleep(0.4)
    _click_at(ppt_hwnd(), 182, 115)
    dlg = by_title("选择用于替换的新图片")
    if not dlg:
        print("!! 替换对话框未出现")
        return
    time.sleep(0.8)
    newimg = os.path.join(OUT, "_demo_new.jpg")
    _set_filename(dlg, newimg)
    time.sleep(0.4)
    capture(dlg, os.path.join(SHOTS, "_raw_replace_pick.png"))
    print("ok _raw_replace_pick.png (已填入新图路径)")
    _click_ok(dlg)
    done = find(lambda h: "完成" in win32gui.GetWindowText(h), timeout=20)
    if not done:
        print("!! 完成提示未出现")
        return
    time.sleep(0.6)
    capture(done, os.path.join(SHOTS, "_raw_replace_done.png"))
    print("ok _raw_replace_done.png")
    close_dialog(done)
    time.sleep(0.6)


def export_slide2():
    """把第 2 页分别按 FHD 与 4K 导出，供文档对比分辨率。"""
    app = get_app()
    pres = get_pres(app)
    w, h = 960, 540
    out1 = os.path.join(SHOTS, "_gen_before.png")
    pres.Slides(2).Export(out1, "PNG", 1920, 1080)
    print("ok _gen_before.png")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "setup"
    fn = globals().get(cmd)
    if callable(fn):
        fn()
    else:
        print("unknown cmd:", cmd)
