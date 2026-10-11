Attribute VB_Name = "modSelfTest"
Option Explicit

' 仅用于自动化测试的驱动模块，不进入正式加载项（仅 --selftest 构建时注入）。
' 全部结果写入文件，避免 COM 调用被中断时丢失结果。

Private Const RPT As String = "__TESTDIR__\vbresult.txt"
Private Const OUTDIR As String = "__TESTDIR__\output"
' 供「选页另存」测试使用的 5 页样张与导出目标（由 run_suite_test.py 准备）
Private Const SLIDE_DECK As String = "__TESTDIR__\output\_deck5.pptx"
Private Const SLIDE_OUT As String = "__TESTDIR__\output\_selected.pptx"
' 供「图片替换」测试使用的带全套属性的样张、新图片与替换结果（由 run_suite_test.py 准备）
Private Const REPLACE_DECK As String = "__TESTDIR__\output\_replace_in.pptx"
Private Const REPLACE_IMG As String = "__TESTDIR__\output\_replace_new.png"
Private Const REPLACE_OUT As String = "__TESTDIR__\output\_replace_out.pptx"
' 供「图片位于编组内」这一场景使用的样张与结果（目标图片在一个带旋转+翻转的组合里）
Private Const GROUP_DECK As String = "__TESTDIR__\output\_group_in.pptx"
Private Const GROUP_OUT As String = "__TESTDIR__\output\_group_out.pptx"

Public Sub LogLine(ByVal s As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open RPT For Append Access Write As #f
    Print #f, s
    Close #f
    On Error GoTo 0
End Sub

Public Sub ClearLog()
    On Error Resume Next
    Kill RPT
    On Error GoTo 0
End Sub

Public Sub CloseTestForm()
    On Error Resume Next
    Dim i As Long
    For i = VBA.UserForms.Count To 1 Step -1
        Unload VBA.UserForms(i - 1)
    Next i
    On Error GoTo 0
End Sub

' 把窗体显示出来供截图核对排版
Public Sub ShowFormForShot()
    On Error Resume Next
    m_ImageExport.gQuiet = True
    m_ImageExport.gDebug = True
    Dim f As Object
    Set f = VBA.UserForms.Add("frmImageExport")
    f.Show vbModeless
    DoEvents
End Sub

Public Function TestExport(ByVal outDir As String, ByVal firstSlide As Long, _
                           ByVal pxW As Long, ByVal pxH As Long, _
                           ByVal fmtIdx As Long, ByVal prefix As String) As String
    Dim flags() As Boolean, ok As Long, fail As String
    Dim n As Long
    On Error GoTo EH
    n = Application.ActivePresentation.Slides.Count
    ReDim flags(1 To n)
    flags(firstSlide) = True
    m_ImageExport.ExportImages flags, outDir, prefix, fmtIdx, pxW, pxH, ok, fail
    TestExport = "OK=" & ok & ";FAIL=" & Replace(fail, vbCrLf, " ") & _
                 ";FILE=" & outDir & "\" & prefix & "_" & Format$(firstSlide, "000") & "." & _
                 m_ImageExport.FormatExt(fmtIdx)
    Exit Function
EH:
    TestExport = "ERROR: " & Err.Number & " " & Err.Description
End Function

Public Function TestDpi(ByVal dpi As Double) As String
    Dim wIn As Double, hIn As Double, w As Long, h As Long
    On Error GoTo EH
    m_ImageExport.GetSlideSizeInches Application.ActivePresentation, wIn, hIn
    m_ImageExport.DpiToPixels dpi, wIn, hIn, w, h
    TestDpi = Format$(wIn, "0.000") & "x" & Format$(hIn, "0.000") & "in -> " & w & "x" & h
    Exit Function
EH:
    TestDpi = "ERROR: " & Err.Number & " " & Err.Description
End Function

Public Function TestRange(ByVal spec As String) As String
    Dim s As String
    On Error GoTo EH
    s = m_ImageExport.ParseSlideRangeToList(spec, Application.ActivePresentation.Slides.Count)
    TestRange = spec & " -> [" & s & "]"
    Exit Function
EH:
    TestRange = "ERROR: " & Err.Number & " " & Err.Description
End Function

' ---- 总入口：逐步写日志，任何一步失败都不影响后续 ----
Public Sub RunAllTests()
    On Error Resume Next

    ClearLog
    LogLine "START slides=" & Application.ActivePresentation.Slides.Count
    LogLine "MODULE_COMPILE=OK"

    LogLine "DPI96=" & TestDpi(96#)
    LogLine "DPI150=" & TestDpi(150#)
    LogLine "DPI300=" & TestDpi(300#)
    LogLine "DPI600=" & TestDpi(600#)
    LogLine "RANGE_A=" & TestRange("1-3")
    LogLine "RANGE_B=" & TestRange("1,3")
    LogLine "RANGE_C=" & TestRange("2-1")
    LogLine "RANGE_D=" & TestRange("1-99")
    LogLine "RANGE_E=" & TestRange("abc")
    LogLine "ASPECT=" & m_ImageExport.AspectText(3840, 2160)
    LogLine "SAFENAME=" & m_ImageExport.SafeFileName("a/b:c*d?e")

    ' --- 窗体 ---
    m_ImageExport.gQuiet = True
    m_ImageExport.gLastError = ""
    m_ImageExport.gDebug = True
    Dim f As Object
    LogLine "  FORM.add enter"
    Set f = VBA.UserForms.Add("frmImageExport")
    LogLine "  FORM.add returned"
    LogLine "FORM_CREATED initerr=[" & m_ImageExport.gLastError & "]"
    LogLine "SELFTEST=" & f.SelfTest()
    LogLine "FORM_ERR=[" & m_ImageExport.gLastError & "]"

    ' --- 实际导出 ---
    LogLine "EXP_4K_PNG=" & TestExport(OUTDIR, 1, 3840, 2160, 0, "T4K")
    LogLine "EXP_FHD_JPG=" & TestExport(OUTDIR, 2, 1920, 1080, 1, "TFHD")
    LogLine "EXP_300_PNG=" & TestExport(OUTDIR, 3, 4000, 2250, 0, "T300")

    ' --- 插件②「选页另存」：验证无窗口打开 + 按序号保留页面 ---
    LogLine "SLIDE_EXPORT=" & m_SlideExport.SlideExport_SelfTest(SLIDE_DECK, SLIDE_OUT)

    ' --- 插件③「图片替换」：验证位置/尺寸/裁剪/旋转/透明度/图层/元数据全保真 ---
    LogLine "IMG_REPLACE=" & m_ImageReplace.ImageReplace_SelfTest( _
                REPLACE_DECK, REPLACE_IMG, REPLACE_OUT)

    ' --- 插件③ 组合场景：图片位于「编组」内时的替换 ---
    '     验证组合自身的结构/几何/旋转/翻转/图层，以及图片属性都保住了。
    LogLine "IMG_GROUP=" & m_ImageReplace.ImageReplace_GroupSelfTest( _
                GROUP_DECK, REPLACE_IMG, GROUP_OUT)

    CloseTestForm
    LogLine "END"
End Sub
