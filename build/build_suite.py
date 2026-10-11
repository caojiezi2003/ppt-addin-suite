# -*- coding: utf-8 -*-
"""
PPT 增强插件套件 —— 构建脚本

设计目标：新增插件 = 在 src/ 下放一个目录（含 plugin.json + 源码），重跑本脚本。
          构建器不需要改动。功能区 XML 由各插件清单自动合成。

目录约定：
    suite.json              套件元信息（名称 / 功能区标签）
    src/<插件>/plugin.json   插件清单（分组、按钮、组件文件）
    src/<插件>/*.bas         标准模块
    src/<插件>/*.code.vba    窗体代码模块

流程：
    ① 扫描 src/*/plugin.json，按 order 排序
    ② 跨插件过程名冲突检查（Public 过程 / 常量 / 模块级变量）
    ③ 合成 customUI/customUI.xml（单一 tab，按插件分 group）
    ④ 阶段 A  准备 VBA 工程访问权限
    ⑤ 阶段 B  通过 PowerPoint COM 生成 VBA 工程并另存为 .ppam
    ⑥ 阶段 C  注入 customUI + 修正内容类型 + 校验（含回调存在性）
    ⑦ 阶段 D  安装到 %APPDATA%\\Microsoft\\AddIns 并登记注册表

用法：
    python build_suite.py                 # 完整构建 + 安装
    python build_suite.py --no-install    # 只构建
    python build_suite.py --register-only # 只把 dist 里的产物安装登记
    python build_suite.py --list          # 只列出扫描到的插件
    python build_suite.py --disable-legacy# 按 suite.json 的 legacy 清单清理旧加载项
"""

import os
import re
import sys
import json
import shutil
import zipfile
import argparse

try:
    import winreg
    import win32com.client as com
except ImportError:
    sys.stderr.write("缺少 pywin32，请先安装：pip install pywin32\n")
    sys.exit(2)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(ROOT, "src")
DIST_DIR = os.path.join(ROOT, "dist")
CUSTOMUI_DIR = os.path.join(ROOT, "customUI")

VBE_CT_STDMODULE = 1
VBE_CT_MSFORM = 3
PP_SAVEAS_ADDIN = 30
PP_SAVEAS_PPTM = 25

CT_PPTM = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
CT_PPAM = "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
REL_UI_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"

log_lines = []


def log(msg):
    print(msg)
    log_lines.append(str(msg))


def read_text(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


# 构建期占位符：源码里写 __TESTDIR__，注入 VBA 工程时替换为本机 test 目录的绝对路径。
# 这样 test/ 下的自检模块不必硬编码某台机器上的路径，换机或他人克隆后可直接跑 --selftest。
PLACEHOLDERS = {
    "__TESTDIR__": os.path.join(ROOT, "test"),
}


def read_vba_source(path):
    """读取待注入 VBA 工程的源码，并展开构建期占位符。"""
    code = read_text(path)
    for key, val in PLACEHOLDERS.items():
        code = code.replace(key, val)
    return code


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def check_encoding(name, code):
    """源码里不得出现「本机代码页（GBK）表示不了」的字符。

    为什么必须有这道闸门（本项目实际踩过的坑）：
        AddFromString 本身走 Unicode，注入这一步没问题；但 VBE **保存工程**时
        会按本机代码页（简中 = GBK / CP936）落盘，表示不了的字符会被
        **静默替换成 '?'**。也就是说：源码里写一个 ⚠ 或 emoji，构建日志一切
        正常、包校验也全绿，但生成出来的 .ppam 里那行已经变成 `?`。
        注释被毁还算小事；若出现在字符串字面量里（输出路径、功能区回调名、
        提示文案）就是实打实的功能缺陷，而且极难定位。

        本闸门把它从「静默损坏」变成「构建期明确报错」。
    """
    bad = {}
    for ch in set(code):
        if ord(ch) < 128:
            continue
        try:
            ch.encode("gbk")
        except UnicodeEncodeError:
            bad[ch] = code.count(ch)
    if bad:
        detail = "、".join("'%s'(U+%04X，%d 处)" % (ch, ord(ch), n)
                          for ch, n in sorted(bad.items(), key=lambda x: -x[1]))
        raise RuntimeError(
            "组件 %s 的源码含有本机代码页（GBK）无法表示的字符，"
            "保存进 .ppam 时会被静默替换成 '?'：%s\n"
            "  请改用 GBK 内的等价写法（例如  ⚠  →  「注意」）。"
            % (name, detail))


def xml_esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


# ================================================================ ① 扫描

def load_suite():
    return load_json(os.path.join(ROOT, "suite.json"))


def load_plugins():
    """扫描 src/ 下所有含 plugin.json 的目录，按 order 排序"""
    plugins = []
    if not os.path.isdir(SRC_DIR):
        raise RuntimeError("缺少 src 目录")
    for name in sorted(os.listdir(SRC_DIR)):
        d = os.path.join(SRC_DIR, name)
        manifest = os.path.join(d, "plugin.json")
        if os.path.isdir(d) and os.path.exists(manifest):
            m = load_json(manifest)
            m["_dir"] = d
            m["_dirname"] = name
            m.setdefault("order", 999)
            m.setdefault("buttons", [])
            m.setdefault("components", [])
            plugins.append(m)
    plugins.sort(key=lambda p: (p["order"], p["_dirname"]))
    return plugins


# ================================================================ ② 冲突检查

PUB_DECL = re.compile(
    r"^\s*Public\s+(Sub|Function|Property\s+(?:Get|Let|Set)|Const|Enum|Type)\s+([A-Za-z_]\w*)",
    re.I)
PUB_VAR = re.compile(
    r"^\s*Public\s+(?!Sub|Function|Property|Const|Enum|Type|Declare|Event)"
    r"([A-Za-z_]\w*)\s", re.I)


def scan_public_members(path):
    """提取一个源码文件里所有 Public 过程 / 常量 / 模块级变量名"""
    names = []
    for line in read_text(path).splitlines():
        s = line.strip()
        if s.startswith("'") or s.startswith("Rem "):
            continue
        m = PUB_DECL.match(line)
        if m:
            names.append(m.group(2))
            continue
        m = PUB_VAR.match(line)
        if m:
            names.append(m.group(1))
    return names


def check_conflicts(plugins):
    """跨插件检查 Public 成员重名 —— 重名会导致 VBA 工程编译失败"""
    owner = {}
    conflicts = []
    for p in plugins:
        for comp in p["components"]:
            path = os.path.join(p["_dir"], comp["file"])
            if not os.path.exists(path):
                conflicts.append("插件 %s 声明的文件不存在：%s" % (p["id"], comp["file"]))
                continue
            for n in scan_public_members(path):
                key = n.lower()
                if key in owner and owner[key] != p["id"]:
                    conflicts.append(
                        "Public 成员重名：%s  同时出现在 %s 与 %s"
                        % (n, owner[key], p["id"]))
                else:
                    owner[key] = p["id"]
    return conflicts


# ================================================================ ③ 合成功能区

def gen_customui(suite, plugins):
    tab = suite["tab"]
    out = []
    out.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    out.append('<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">')
    out.append('  <ribbon>')
    out.append('    <tabs>')
    out.append('      <tab id="%s" label="%s" insertAfterMso="%s">'
               % (xml_esc(tab["id"]), xml_esc(tab["label"]),
                  xml_esc(tab.get("insertAfterMso", "TabInsert"))))
    for p in plugins:
        g = p.get("group")
        if not g:
            continue
        attrs = 'id="%s" label="%s"' % (xml_esc(g["id"]), xml_esc(g["label"]))
        if g.get("imageMso"):
            attrs += ' imageMso="%s"' % xml_esc(g["imageMso"])
        out.append('        <group %s>' % attrs)
        for b in p["buttons"]:
            battrs = ['id="%s"' % xml_esc(b["id"]),
                      'label="%s"' % xml_esc(b["label"]),
                      'size="%s"' % xml_esc(b.get("size", "large")),
                      'onAction="%s"' % xml_esc(b["onAction"])]
            if b.get("imageMso"):
                battrs.append('imageMso="%s"' % xml_esc(b["imageMso"]))
            if b.get("screentip"):
                battrs.append('screentip="%s"' % xml_esc(b["screentip"]))
            if b.get("supertip"):
                battrs.append('supertip="%s"' % xml_esc(b["supertip"]))
            out.append('          <button %s/>' % "\n                  ".join(battrs))
        out.append('        </group>')
    out.append('      </tab>')
    out.append('    </tabs>')
    out.append('  </ribbon>')
    out.append('</customUI>')
    return "\n".join(out) + "\n"


def check_callbacks(plugins, xml):
    """customUI 里引用的回调，必须在源码里真实存在，否则点击无反应"""
    defined = set()
    for p in plugins:
        for comp in p["components"]:
            path = os.path.join(p["_dir"], comp["file"])
            if not os.path.exists(path):
                continue
            for line in read_text(path).splitlines():
                m = re.match(r"^\s*Public\s+(?:Sub|Function)\s+([A-Za-z_]\w*)", line, re.I)
                if m:
                    defined.add(m.group(1).lower())
    missing = []
    for name in re.findall(r'onAction="([^"]+)"', xml):
        if name.lower() not in defined:
            missing.append(name)
    return missing, sorted(defined)


# ================================================================ ④ 阶段 A

def ensure_vbom_access(office_ver):
    key_path = r"Software\Microsoft\Office\%s\PowerPoint\Security" % office_ver
    changed = False
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0,
                            winreg.KEY_READ | winreg.KEY_WRITE) as k:
        try:
            cur = winreg.QueryValueEx(k, "AccessVBOM")[0]
        except FileNotFoundError:
            cur = None
        if cur != 1:
            winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            changed = True
            log("[A] 已开启「信任对 VBA 工程对象模型的访问」AccessVBOM = 1")
        else:
            log("[A] AccessVBOM 已为 1，无需修改")
    return changed


def powerpoint_running():
    try:
        com.GetActiveObject("PowerPoint.Application")
        return True
    except Exception:
        return False


# ================================================================ ⑤ 阶段 B

def set_form_size(comp, w_pt, h_pt):
    def get(key):
        try:
            return comp.Properties(key)
        except Exception:
            return None

    def put(key, val):
        p = get(key)
        if p is not None:
            try:
                p.Value = val
                return True
            except Exception:
                pass
        try:
            comp.Designer.SetValue(key, val)
            return True
        except Exception:
            pass
        try:
            setattr(comp.Designer, key, val)
            return True
        except Exception:
            return False

    ok = put("Width", w_pt) and put("Height", h_pt)
    if not ok:
        log("[B]   警告：窗体 %s 设计期尺寸设置失败（回退默认值）" % comp.Name)
    return ok


def build_com(suite, plugins, patches=None):
    os.makedirs(DIST_DIR, exist_ok=True)
    addin_filename = suite["suite"]["addin_filename"]
    raw_path = os.path.join(DIST_DIR, addin_filename)
    if os.path.exists(raw_path):
        os.remove(raw_path)

    app = com.DispatchEx("PowerPoint.Application")
    try:
        app.Visible = True
    except Exception:
        pass

    pres = None
    try:
        pres = app.Presentations.Add()
        while pres.Slides.Count > 0:
            pres.Slides(1).Delete()

        try:
            vbp = pres.VBProject
        except Exception as e:
            raise RuntimeError(
                "无法访问 VBA 工程（AccessVBOM 未生效）。请完全退出 PowerPoint "
                "后重新运行本脚本。原始错误：%s" % e)

        for p in plugins:
            for comp_spec in p["components"]:
                path = os.path.join(p["_dir"], comp_spec["file"])
                code = read_vba_source(path)
                # AddFromString 走 Unicode；同时滤除 Attribute 头（组件名用 Name 设定）
                code = "\r\n".join(
                    l for l in code.replace("\r\n", "\n").split("\n")
                    if not l.strip().lower().startswith("attribute "))

                # 自检注入（仅 --selftest 构建）：按「组件名」追加测试代码。
                # 追加内容是同一模块内部的代码，因此可访问该模块的 Private 成员。
                patch = (patches or {}).get(comp_spec["name"])
                if patch:
                    code = code + "\r\n" + "\r\n".join(
                        patch.replace("\r\n", "\n").split("\n"))

                # 字符集闸门：代码页表示不了的字符会被 VBE 静默换成 '?'，必须挡在这里
                check_encoding(comp_spec["name"], code)

                if comp_spec["type"] == "module":
                    comp = vbp.VBComponents.Add(VBE_CT_STDMODULE)
                    comp.Name = comp_spec["name"]
                    cm = comp.CodeModule
                    if cm.CountOfLines > 0:
                        cm.DeleteLines(1, cm.CountOfLines)
                    cm.AddFromString(code)
                    log("[B]   + 模块 %-18s %4d 行  (%s)"
                        % (comp_spec["name"], cm.CountOfLines, p["id"]))
                elif comp_spec["type"] == "form":
                    comp = vbp.VBComponents.Add(VBE_CT_MSFORM)
                    comp.Name = comp_spec["name"]
                    set_form_size(comp, comp_spec.get("design_width_pt", 392),
                                  comp_spec.get("design_height_pt", 312))
                    body = code.replace("\r\n", "\n").replace("\n", "\r\n")
                    cm = comp.CodeModule
                    if cm.CountOfLines > 0:
                        cm.DeleteLines(1, cm.CountOfLines)
                    cm.AddFromString(body)
                    log("[B]   + 窗体 %-18s %4d 行  (%s)  设计尺寸 %sx%s pt"
                        % (comp_spec["name"], cm.CountOfLines, p["id"],
                           comp_spec.get("design_width_pt", "?"),
                           comp_spec.get("design_height_pt", "?")))
                else:
                    raise RuntimeError("未知组件类型：%s" % comp_spec["type"])

        try:
            pres.SaveAs(raw_path, PP_SAVEAS_ADDIN)
            log("[B] 已另存为加载项格式 (.ppam)")
        except Exception as e:
            log("[B] 直接另存为 .ppam 失败（%s），改为先存 .pptm 再改写内容类型" % e)
            raw_path = os.path.join(DIST_DIR, "PptEnhanceSuite.pptm")
            if os.path.exists(raw_path):
                os.remove(raw_path)
            pres.SaveAs(raw_path, PP_SAVEAS_PPTM)
            log("[B] 已另存为宏演示文稿 (.pptm)")
        return raw_path
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass


# ================================================================ ⑥ 阶段 C

def patch_opc(raw_path, out_path, ui_xml):
    with zipfile.ZipFile(raw_path, "r") as zin:
        names = zin.namelist()
        parts = {n: zin.read(n) for n in names}

    # 1) 内容类型 pptm -> ppam
    ct_name = "[Content_Types].xml"
    if ct_name not in parts:
        raise RuntimeError("包内缺少 [Content_Types].xml")
    ct = parts[ct_name].decode("utf-8")
    if CT_PPTM in ct:
        ct = ct.replace(CT_PPTM, CT_PPAM)
        log("[C] 内容类型已由 pptm 改写为 ppam")
    elif CT_PPAM in ct:
        log("[C] 内容类型已是 ppam")
    else:
        log("[C] 警告：未识别到主部件内容类型，保持原样")
    parts[ct_name] = ct.encode("utf-8")

    # 2) 包级关系：加入 customUI
    rels_name = "_rels/.rels"
    if rels_name not in parts:
        raise RuntimeError("包内缺少 _rels/.rels")
    rels = parts[rels_name].decode("utf-8")
    if REL_UI_TYPE not in rels:
        rid = "rIdEnhanceSuiteUI"
        existing = {t for t in rels.split('"') if t.startswith("rId")}
        n = 1
        while rid in existing:
            n += 1
            rid = "rIdEnhanceSuiteUI%d" % n
        rel = ('<Relationship Id="%s" Type="%s" Target="customUI/customUI.xml"/>'
               % (rid, REL_UI_TYPE))
        rels = rels.replace("</Relationships>", rel + "</Relationships>")
        log("[C] 已加入 customUI 关系 (%s)" % rid)
    else:
        log("[C] customUI 关系已存在")
    parts[rels_name] = rels.encode("utf-8")

    # 3) 写入 customUI.xml
    parts["customUI/customUI.xml"] = ui_xml.encode("utf-8")
    log("[C] 已写入 customUI/customUI.xml (%d 字节)" % len(ui_xml.encode("utf-8")))

    # 4) 重新打包（[Content_Types].xml 置首）
    if os.path.exists(out_path):
        os.remove(out_path)
    order = [ct_name] + [n for n in names if n != ct_name]
    if "customUI/customUI.xml" not in order:
        order.append("customUI/customUI.xml")

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in order:
            if n in parts:
                zout.writestr(n, parts[n])

    log("[C] 已重新打包 -> %s (%.1f KB)"
        % (os.path.basename(out_path), os.path.getsize(out_path) / 1024.0))


def verify_package(path, ui_xml, expect_callbacks):
    ok = True
    with zipfile.ZipFile(path, "r") as z:
        names = z.namelist()
        ct = z.read("[Content_Types].xml").decode("utf-8")
        rels = z.read("_rels/.rels").decode("utf-8")
        ui_in_pack = (z.read("customUI/customUI.xml").decode("utf-8")
                      if "customUI/customUI.xml" in names else "")

    checks = [
        ("主部件为 ppam 内容类型", CT_PPAM in ct),
        ("customUI/customUI.xml 存在", "customUI/customUI.xml" in names),
        ("customUI 关系存在", REL_UI_TYPE in rels),
        ("VBA 工程存在", any(n.lower().endswith("vbaproject.bin") for n in names)),
        ("包内 customUI 与源文件一致", ui_in_pack.strip() == ui_xml.strip()),
    ]
    for label, passed in checks:
        log("[C] %s 校验：%s" % (label, "通过" if passed else "失败"))
        ok = ok and passed

    missing = [c for c in expect_callbacks if c not in ui_in_pack]
    log("[C] 回调存在性校验：%s" % ("通过" if not missing else "缺失 %s" % missing))
    ok = ok and not missing
    return ok


# ================================================================ ⑦ 阶段 D

def register_addin_for_startup(addin_path, office_ver, reg_name):
    key_path = r"Software\Microsoft\Office\%s\PowerPoint\AddIns\%s" % (office_ver, reg_name)
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0,
                            winreg.KEY_READ | winreg.KEY_WRITE) as k:
        winreg.SetValueEx(k, "Path", 0, winreg.REG_SZ, addin_path)
        winreg.SetValueEx(k, "AutoLoad", 0, winreg.REG_DWORD, 1)
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as k:
        path_back = winreg.QueryValueEx(k, "Path")[0]
        auto_back = int(winreg.QueryValueEx(k, "AutoLoad")[0])
    ok = (os.path.normcase(path_back) == os.path.normcase(addin_path) and auto_back == 1)
    log("[D] 注册表登记 HKCU\\...\\%s\\PowerPoint\\AddIns\\%s" % (office_ver, reg_name))
    log("[D]   Path=%s" % path_back)
    log("[D]   AutoLoad=%s  -> %s" % (auto_back, "通过" if ok else "校验失败"))
    return ok


def install(src_path, suite):
    s = suite["suite"]
    office_ver = s.get("office_ver", "16.0")
    addins_dir = os.path.join(os.environ["APPDATA"], "Microsoft", "AddIns")
    os.makedirs(addins_dir, exist_ok=True)
    dst = os.path.join(addins_dir, s["addin_filename"])
    shutil.copy2(src_path, dst)
    log("[D] 已复制到 %s" % dst)

    reg_ok = register_addin_for_startup(dst, office_ver, s["addin_reg_name"])

    if powerpoint_running():
        try:
            app = com.GetActiveObject("PowerPoint.Application")
            added = app.AddIns.Add(dst)
            added.Loaded = True
            log("[D] 已在当前运行的 PowerPoint 会话中即时加载")
        except Exception as e:
            log("[D] 即时加载调用返回：%s（注册表登记已生效，下次启动仍会加载）" % e)
    else:
        log("[D] PowerPoint 未运行，跳过即时加载；下次启动将自动加载")
    return dst, reg_ok


def list_legacy_installed(exclude_names):
    """列出 %APPDATA%\\Microsoft\\AddIns 下除本套件外的其它加载项（供切换时清理）"""
    addins_dir = os.path.join(os.environ["APPDATA"], "Microsoft", "AddIns")
    out = []
    if os.path.isdir(addins_dir):
        for f in sorted(os.listdir(addins_dir)):
            if f.lower().endswith((".ppam", ".ppa")) and f not in exclude_names:
                out.append(os.path.join(addins_dir, f))
    return out


def disable_legacy(suite):
    """按 suite.json 的 legacy 清单，解除并移除被本套件取代的旧加载项。

    设计上刻意「只认清单、不做扫描」：
    清理对象必须显式写在 suite.json 的 legacy 数组里。绝不去遍历注册表或
    AddIns 目录「猜」哪些是旧的 —— 那样会误伤 iSlide 之类与本套件无关的
    第三方加载项。要清理什么，是可审阅、可 diff 的声明。

    两个动作，作用域都被严格限死：
      ① 注册表：只删 HKCU\\...\\PowerPoint\\AddIns\\<清单里的名字> 这一个键
      ② 文件：  只删 %APPDATA%\\Microsoft\\AddIns\\<清单里的文件名> 这一个文件
         —— 按文件名精确匹配，不递归、不通配，绝不触碰用户的其它目录
    """
    s = suite["suite"]
    office_ver = s.get("office_ver", "16.0")
    addins_dir = os.path.join(os.environ["APPDATA"], "Microsoft", "AddIns")
    entries = suite.get("legacy") or []

    log("=" * 66)
    log("增强插件 · 清理被取代的旧加载项")
    log("=" * 66)

    if not entries:
        log("[L] suite.json 里没有 legacy 清单，无可清理项")
        return 0

    if powerpoint_running():
        log("[L] 中止：请先完全退出 PowerPoint。")
        log("      加载项文件在 PowerPoint 运行期间被占用，删不掉；")
        log("      注册表登记也应在关闭状态下修改，避免被回写覆盖。")
        return 3

    log("[L] PowerPoint 未运行，可以安全清理")
    log("[L] 清理范围（仅以下 %d 条，不做任何扫描）：" % len(entries))
    for e in entries:
        log("      - %s" % (e.get("note") or e.get("registry", "?")))

    removed_reg, removed_file = 0, 0
    problems = []

    for e in entries:
        reg = e.get("registry")
        log("")
        log("[L] --- %s" % (e.get("note") or reg or "?"))

        # ---- ① 解除注册表登记 ----
        if reg:
            key_path = r"Software\Microsoft\Office\%s\PowerPoint\AddIns\%s" % (office_ver, reg)
            path_back = "(无 Path 值)"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as k:
                    try:
                        path_back = winreg.QueryValueEx(k, "Path")[0]
                    except FileNotFoundError:
                        pass
            except FileNotFoundError:
                log("[L]   注册表登记不存在，跳过（%s）" % reg)
                reg = None
            except OSError as ex:
                problems.append("读取注册表项 %s 失败：%s" % (reg, ex))
                reg = None

            if reg:
                try:
                    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
                    log("[L]   已删除注册表登记：HKCU\\...\\PowerPoint\\AddIns\\%s" % reg)
                    log("[L]     原指向：%s" % path_back)
                    removed_reg += 1
                except OSError as ex:
                    problems.append("删除注册表项 %s 失败：%s" % (reg, ex))

        # ---- ② 移除 AddIns 目录内的旧文件 ----
        for fn in e.get("files", []):
            fp = os.path.join(addins_dir, fn)
            if not os.path.exists(fp):
                log("[L]   AddIns 内无 %s，跳过" % fn)
                continue
            try:
                os.remove(fp)
                log("[L]   已删除文件：%s" % fp)
                removed_file += 1
            except OSError as ex:
                problems.append("删除文件 %s 失败：%s" % (fp, ex))

    log("")
    log("-" * 66)
    log("[L] 清理完成：注销注册表 %d 项、删除文件 %d 个" % (removed_reg, removed_file))

    leftover = list_legacy_installed({s["addin_filename"]})
    if leftover:
        log("[L] 提示：AddIns 目录下仍有其它加载项（不在本次清单内，未做处理）：")
        for p in leftover:
            log("      - %s" % p)
    else:
        log("[L] AddIns 目录下已无其它加载项")

    if problems:
        log("[L] 存在问题：")
        for p in problems:
            log("[L]   [!] %s" % p)
        rc = 1
    else:
        rc = 0
    log("[L] 重启 PowerPoint 后生效")

    with open(os.path.join(DIST_DIR, "disable-legacy-report.txt"), "w",
              encoding="utf-8") as f:
        f.write("\n".join(log_lines))
    return rc


# ================================================================ main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-install", action="store_true", help="只构建，不安装")
    ap.add_argument("--register-only", action="store_true",
                    help="不重新构建，只把 dist 中已有产物安装并登记")
    ap.add_argument("--list", action="store_true", help="只列出扫描到的插件")
    ap.add_argument("--disable-legacy", action="store_true",
                    help="按 suite.json 的 legacy 清单清理被取代的旧加载项"
                         "（解除注册表登记 + 删除 AddIns 内旧文件，需先退出 PowerPoint）")
    ap.add_argument("--selftest", action="store_true",
                    help="注入自检代码，产出 PptEnhanceSuite-selftest.ppam（不安装）")
    args = ap.parse_args()

    suite = load_suite()
    if args.selftest:
        # 用独立的文件名 / 标签页 id，避免与已注册的正式加载项在同一 PowerPoint
        # 实例中重名冲突（正式版会被注册表自动加载）
        suite["suite"] = dict(suite["suite"],
                              addin_filename="PptEnhanceSuiteTest.ppam")
        suite["tab"] = dict(suite["tab"], id="tabEnhanceTest",
                            label="[测试] 增强插件")
        args.no_install = True
    s = suite["suite"]

    # 清理模式：与构建流程完全无关，处理完直接返回
    if args.disable_legacy:
        return disable_legacy(suite)

    log("=" * 66)
    log("%s · 构建  v%s" % (s["title"], s["version"]))
    log("=" * 66)

    plugins = load_plugins()
    log("[①] 扫描到 %d 个插件：" % len(plugins))
    for p in plugins:
        g = p.get("group")
        log("      order=%-4s %-16s %-12s 组件 %d 个，按钮 %d 个"
            % (p["order"], p["id"], ("组「%s」" % g["label"]) if g else "(仅模块)",
               len(p["components"]), len(p["buttons"])))

    if args.list:
        return 0

    # 自检注入：把测试驱动模块追加为临时插件，并按组件名追加测试代码
    patches = {}
    if args.selftest:
        patches["frmImageExport"] = read_vba_source(
            os.path.join(ROOT, "test", "selftest.vba"))
        patches["m_SlideExport"] = read_vba_source(
            os.path.join(ROOT, "test", "selftest_slide.vba"))
        patches["m_ImageReplace"] = read_vba_source(
            os.path.join(ROOT, "test", "selftest_replace.vba"))
        plugins.append({
            "id": "_test", "order": 9999, "title": "自检驱动", "group": None,
            "buttons": [],
            "components": [{"type": "module", "name": "modSelfTest",
                            "file": "modSelfTest.bas"}],
            "_dir": os.path.join(ROOT, "test"), "_dirname": "_test",
        })
        log("[*] 已追加自检驱动 modSelfTest；注入 SelfTest() -> frmImageExport、"
            "SlideExport_SelfTest() -> m_SlideExport、"
            "ImageReplace_SelfTest() -> m_ImageReplace")

    # ② 冲突检查
    conflicts = check_conflicts(plugins)
    if conflicts:
        log("[②] 命名冲突检查：发现 %d 处问题，构建中止" % len(conflicts))
        for c in conflicts:
            log("      [!] %s" % c)
        return 3
    log("[②] 命名冲突检查：通过（无跨插件重名）")

    # ③ 合成 customUI
    ui_xml = gen_customui(suite, plugins)
    os.makedirs(CUSTOMUI_DIR, exist_ok=True)
    with open(os.path.join(CUSTOMUI_DIR, "customUI.xml"), "w",
              encoding="utf-8", newline="\n") as f:
        f.write(ui_xml)
    missing, defined = check_callbacks(plugins, ui_xml)
    log("[③] 已合成 customUI/customUI.xml")
    log("      源码中可用回调 %d 个：%s" % (len(defined), ", ".join(defined)))
    if missing:
        log("[③] 回调缺失：%s —— 构建中止" % missing)
        return 3
    log("[③] 回调引用检查：通过")

    # 阶段 D（只登记模式）
    if args.register_only:
        src = os.path.join(DIST_DIR, s["addin_filename"])
        if not os.path.exists(src):
            log("[!] 未找到 %s，请先完整构建一次" % src)
            return 4
        dst, reg_ok = install(src, suite)
        log("[D] 安装完成：%s" % dst)
        with open(os.path.join(DIST_DIR, "install-report.txt"), "w",
                  encoding="utf-8") as f:
            f.write("\n".join(log_lines))
        return 0 if reg_ok else 1

    # 阶段 A
    changed = ensure_vbom_access(s.get("office_ver", "16.0"))
    if powerpoint_running():
        log("[A] 检测到 PowerPoint 正在运行")
        if changed:
            log("[A] 错误：刚修改了 AccessVBOM，必须完全退出 PowerPoint 后重试")
            return 3
    else:
        log("[A] PowerPoint 未在运行")

    # 阶段 B
    raw = build_com(suite, plugins, patches=patches)

    # 阶段 C
    final = os.path.join(DIST_DIR, s["addin_filename"])
    patch_opc(raw, final, ui_xml)
    if raw != final and os.path.exists(raw):
        try:
            os.remove(raw)
        except OSError:
            pass

    callbacks = re.findall(r'onAction="([^"]+)"', ui_xml)
    ok = verify_package(final, ui_xml, callbacks)
    log("[C] 包校验总体：%s" % ("通过" if ok else "存在失败项"))

    # 阶段 D
    if not args.no_install:
        dst, reg_ok = install(final, suite)
        log("[D] 安装完成：%s" % dst)
        if not reg_ok:
            log("[D] 警告：注册表登记校验未通过，加载项可能不会随启动自动加载")
        legacy = list_legacy_installed({s["addin_filename"]})
        if legacy:
            log("[D] 提示：AddIns 目录下还有其它加载项，可能与本套件出现重复入口：")
            for p in legacy:
                log("      - %s" % p)
            log("      确认本套件工作正常后，可移除它们（或运行 --disable-legacy）")
    else:
        log("[D] 已跳过安装")

    log("")
    log("产物：%s" % final)
    with open(os.path.join(DIST_DIR, "build-report.txt"), "w",
              encoding="utf-8") as f:
        f.write("\n".join(log_lines))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
