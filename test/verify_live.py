# -*- coding: utf-8 -*-
"""
阶段 2 实时验证：在真实 PowerPoint 会话中确认
  · 套件加载项被加载（Loaded=True）
  · VBA 工程结构包含预期的模块与窗体
  · 宏可被 Application.Run 实际调用

安全约定：若检测到已有 PowerPoint 会话，直接跳过，不打扰用户正在编辑的文稿。
"""

import os
import sys

import win32com.client as com

HERE = os.path.dirname(os.path.abspath(__file__))
SUITE = os.path.dirname(HERE)
ADDIN = os.path.join(SUITE, "dist", "PptEnhanceSuite.ppam")
PROJ = "PptEnhanceSuite"

EXPECT = ["m_Core", "m_ImageExport", "frmImageExport", "m_SlideExport",
          "m_ImageReplace"]


def main():
    try:
        com.GetActiveObject("PowerPoint.Application")
        print("SKIP: 检测到已有 PowerPoint 会话，跳过验证以免打扰。")
        return 0
    except Exception:
        pass

    app = com.DispatchEx("PowerPoint.Application")
    rc = 0
    try:
        app.Visible = True
        added = app.AddIns.Add(ADDIN)
        added.Loaded = True
        print("ADDIN: name=%s  loaded=%s" % (added.Name, bool(added.Loaded)))
        if not added.Loaded:
            rc = 1

        proj = None
        target = None
        try:
            vbe = app.VBE
            n = vbe.VBProjects.Count
            print("VBE : VBProjects.Count = %d" % n)
            for i in range(1, n + 1):
                try:
                    p = vbe.VBProjects.Item(i)
                except Exception as e:
                    print("VBE :   工程[%d] 不可访问：%s" % (i, e))
                    continue
                try:
                    cn = [p.VBComponents.Item(k + 1).Name
                          for k in range(p.VBComponents.Count)]
                except Exception as e:
                    # 受密码保护的工程读不到组件列表，跳过即可
                    print("VBE :   工程 %-26s 组件列表不可读（受保护）" % p.Name)
                    continue
                print("VBE :   工程 %-26s 组件 %s" % (p.Name, sorted(cn)))
                if "m_SlideExport" in cn:
                    proj = p
                    target = cn
        except Exception as e:
            print("VBE : 无法访问 VBE：%s" % e)

        if target is None:
            print("VBE : .ppam 加载项的 VBA 工程默认受保护，无法从对象模型枚举组件"
                  "（这是预期行为）；模块清单改由构建日志与包内解析核对，"
                  "此处以宏执行结果为准")
        else:
            print("VBE : 定位到套件工程 %s" % proj.Name)
            for want in EXPECT:
                ok = want in target
                if not ok:
                    rc = 1
                print("VBE :   包含 %-16s %s" % (want, "OK" if ok else "缺失"))

        app._FlagAsMethod("Run")
        try:
            print("RUN : FmtNum(288) -> %r"
                  % app.Run("'%s.ppam'!m_ImageExport.FmtNum" % PROJ, 288.0))
        except Exception as e:
            print("RUN : FmtNum 失败：%s" % e)
            rc = 1
        try:
            print("RUN : ParseSlideRangeToList('1-3,5',10) -> %r"
                  % app.Run("'%s.ppam'!m_ImageExport.ParseSlideRangeToList" % PROJ,
                            "1-3,5", 10))
        except Exception as e:
            print("RUN : ParseSlideRangeToList 失败：%s" % e)
            rc = 1

        # 组合（编组）内替换属于 m_ImageReplace；这里做一次无害的代码路径冒烟：
        # ImageReplace_Describe 不带参数会立刻返回 "(Nothing)"，不触碰任何文档。
        # 这一步真正验证的是「m_ImageReplace 被加载并且能编译、能执行」——
        # 只靠包内解析或模块清单都证明不了这一点。
        try:
            print("RUN : ImageReplace_Describe() -> %r"
                  % app.Run("'%s.ppam'!m_ImageReplace.ImageReplace_Describe" % PROJ))
        except Exception as e:
            print("RUN : ImageReplace_Describe 失败：%s" % e)
            rc = 1
    finally:
        try:
            app.Quit()
        except Exception:
            pass
    return rc


if __name__ == "__main__":
    sys.exit(main())
