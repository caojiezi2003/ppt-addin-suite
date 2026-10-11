Attribute VB_Name = "m_Core"
Option Explicit

' ============================================================
'  PPT 增强插件套件 — 公共核心层
'
'  本模块不承载任何具体功能，只提供：
'    · 套件标识（标题 / 版本）
'    · 统一的错误出口
'    · 供各插件共用的通用 Ribbon 回调
'
'  设计原则：每个插件自带自己的 Ribbon 回调（Ribbon_<插件>_<动作>），
'  不做集中分发 —— 这样新增插件只要放一个目录，不必改动任何已有代码。
' ============================================================

Public Const SUITE_TITLE As String = "增强插件"
Public Const SUITE_VERSION As String = "1.0.0"

' 统一错误出口。各插件在 Ribbon 回调的 EH 段调用：
'     ReportError "导出图片", Err.Number, Err.Description
Public Sub ReportError(ByVal ctx As String, ByVal num As Long, ByVal desc As String)
    Dim msg As String
    On Error Resume Next
    msg = "【" & ctx & "】运行出错" & vbCrLf & "错误 " & num & "：" & desc
    Debug.Print msg
    MsgBox msg, vbCritical, SUITE_TITLE
    On Error GoTo 0
End Sub

' 通用 Ribbon 回调：按钮写 getEnabled="RibbonGetEnabled" 即恒为可用
Public Sub RibbonGetEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub
