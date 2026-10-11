Attribute VB_Name = "m_SlideExport"
Option Explicit

' ============================================================
'  选页另存 —— PPT 增强插件套件 / 插件
'
'  在左侧缩略图窗格中选中若干幻灯片（单击选一页、Ctrl 多选不连续的页、
'  Shift 选连续区间），把选中的页另存为一份新的 .pptx。
'  版式 / 母版 / 页面尺寸 / 主题 / 备注 / 动画 / 切换全部原样保留。
'
'  实现方式：
'    整本另存为临时副本 → 从后往前删掉未选中的页 → 再另存为用户指定的文件。
'    这样 100% 保留原版式，避免“复制到新文稿”造成的版式丢失。
'
'  套件规范化命名：
'    模块        m_SlideExport
'    Ribbon 回调 Ribbon_SlideExport_Selected
' ============================================================

' ------------------------------------------------------------
'  Ribbon 回调入口
' ------------------------------------------------------------
Public Sub Ribbon_SlideExport_Selected(control As IRibbonControl)
    On Error GoTo EH
    SlideExport_Run
    Exit Sub
EH:
    ReportError "选页另存", Err.Number, Err.Description
End Sub

' ------------------------------------------------------------
'  交互主流程：取选中页 → 问路径 → 导出
' ------------------------------------------------------------
Private Sub SlideExport_Run()
    Dim srcRange As SlideRange
    Dim srcPres  As Presentation
    Dim fd       As FileDialog
    Dim selIdx() As Long
    Dim i        As Long
    Dim outPath  As String
    Dim errMsg   As String

    ' 1) 取当前缩略图窗格中选中的幻灯片
    On Error Resume Next
    Set srcRange = ActiveWindow.Selection.SlideRange
    On Error GoTo 0

    If srcRange Is Nothing Then
        MsgBox "请先打开一份演示文稿，并在左侧缩略图窗格中选中至少一个幻灯片。" & vbCrLf & _
               "（单击选中一页；按住 Ctrl 可多选不连续的页；Shift 选连续区间）", _
               vbExclamation, "未选中幻灯片"
        Exit Sub
    End If
    If srcRange.Count = 0 Then Exit Sub

    Set srcPres = ActivePresentation
    If srcPres Is Nothing Then Exit Sub

    ' 2) 记下选中页的原始序号（删页时据此判定保留）
    ReDim selIdx(1 To srcRange.Count)
    For i = 1 To srcRange.Count
        selIdx(i) = srcRange(i).SlideIndex
    Next i

    ' 3) 询问输出路径
    Set fd = Application.FileDialog(2)          ' 2 = msoFileDialogSaveAs
    fd.Title = "将选中的 " & srcRange.Count & " 张幻灯片另存为"
    fd.InitialFileName = SlideExport_SuggestName()
    If fd.Show <> -1 Then Exit Sub
    outPath = CStr(fd.SelectedItems(1))
    If LCase$(Right$(outPath, 5)) <> ".pptx" Then outPath = outPath & ".pptx"

    ' 4) 执行导出
    If SlideExport_Export(srcPres, selIdx, outPath, errMsg) Then
        MsgBox "已导出 " & srcRange.Count & " 张幻灯片：" & vbCrLf & outPath, _
               vbInformation, "导出完成"
    Else
        MsgBox errMsg, vbCritical, "导出失败"
    End If
End Sub

' ------------------------------------------------------------
'  核心导出：把 srcPres 中序号属于 selIdx 的幻灯片另存为 outPath
'
'    srcPres   源演示文稿
'    selIdx    要保留的幻灯片序号（1 基，对应 SlideIndex）
'    outPath   输出 .pptx 全路径
'    errMsg    失败时返回原因
'    返回      True 成功 / False 失败
'
'  说明：与界面无关的纯逻辑，便于自动化验证。
' ------------------------------------------------------------
Private Function SlideExport_Export(ByVal srcPres As Presentation, _
                                    ByRef selIdx() As Long, _
                                    ByVal outPath As String, _
                                    ByRef errMsg As String) As Boolean
    Dim newPres As Presentation
    Dim fso     As Object
    Dim tmpPath As String
    Dim s       As Long
    Dim kept    As Long

    On Error GoTo EH

    Set fso = CreateObject("Scripting.FileSystemObject")
    ' GetTempName() 保证唯一，避免同秒内多次导出互相覆盖临时文件
    tmpPath = fso.BuildPath(fso.GetSpecialFolder(2), _
                            "PPTX_" & Replace(fso.GetTempName(), ".tmp", "") & ".pptx")

    ' 整本另存为临时副本 —— 保留全部版式 / 母版 / 页面尺寸 / 主题
    srcPres.SaveCopyAs tmpPath

    ' 【修复】原写法 Presentations.Open(tmpPath, 0) 中的 0 落在第 2 个参数 ReadOnly 上，
    '         WithWindow 仍取默认值 msoTrue —— 导出过程中会闪出一个临时文稿窗口。
    '         Open 的签名是 Open(FileName, ReadOnly, Untitled, WithWindow)，
    '         此处把第 4 个参数显式设为 msoFalse（无窗口打开）。
    Set newPres = Presentations.Open(tmpPath, , , msoFalse)

    ' 从后往前删，删掉不在保留集合里的页（倒序可保证删除不影响未处理的序号）
    For s = newPres.Slides.Count To 1 Step -1
        If Not SlideExport_IsInArray(newPres.Slides(s).SlideIndex, selIdx) Then
            newPres.Slides(s).Delete
        End If
    Next s

    kept = newPres.Slides.Count
    If kept = 0 Then
        errMsg = "选中的幻灯片在副本中不存在，未生成文件。"
        newPres.Close
        On Error Resume Next: Kill tmpPath: On Error GoTo 0
        Exit Function
    End If

    ' 24 = ppSaveAsOpenXMLPresentation
    newPres.SaveAs outPath, 24
    newPres.Close
    On Error Resume Next: Kill tmpPath: On Error GoTo 0

    SlideExport_Export = True
    Exit Function

EH:
    errMsg = "导出失败：" & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not newPres Is Nothing Then newPres.Close
    If Len(tmpPath) > 0 Then Kill tmpPath
    On Error GoTo 0
    SlideExport_Export = False
End Function

' ------------------------------------------------------------
'  工具
' ------------------------------------------------------------
Private Function SlideExport_IsInArray(ByVal v As Long, ByRef arr() As Long) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = v Then
            SlideExport_IsInArray = True
            Exit Function
        End If
    Next i
    SlideExport_IsInArray = False
End Function

' 默认文件名：<原文件名>_selected.pptx
Private Function SlideExport_SuggestName() As String
    Dim base As String
    Dim fso  As Object
    On Error Resume Next
    base = ActivePresentation.FullName
    On Error GoTo 0
    Set fso = CreateObject("Scripting.FileSystemObject")
    If base = "" Then
        SlideExport_SuggestName = "selected_slides.pptx"
    Else
        SlideExport_SuggestName = fso.GetBaseName(base) & "_selected.pptx"
    End If
End Function
