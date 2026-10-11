Attribute VB_Name = "m_ImageExport"
Option Explicit

' ============================================================
'  PPT 图片导出 — 核心引擎模块
'
'  要解决的问题：
'    PowerPoint 的「另存为 → PNG/JPG」只能按内置的固定分辨率导出
'    （默认 96 DPI，受注册表 ExportBitmapResolution 全局控制，
'     且存在 maxdpi = sqrt(1e8 / (宽英寸 * 高英寸)) 的硬上限），
'    无法在导出时指定像素尺寸或 DPI。
'
'  解决原理：
'    Slide.Export(路径, 格式, 宽像素, 高像素) 可逐页指定输出像素，
'    完全绕开注册表限制。像素 = 物理尺寸(英寸) x DPI，
'    因此「自定义分辨率」与「自定义像素尺寸」在数学上是同一件事，
'    本模块同时接受两种输入并互相换算。
' ============================================================

Public Const APP_TITLE As String = "PPT 图片导出"
Public Const APP_VERSION As String = "1.0.0"

Private Const CFG_APP As String = "PptImageExport"

' 大图预警阈值（像素总数）：超过约 4000 万像素时提示用户
Private Const BIG_PIXEL_WARN As Double = 40000000#

' 运行时状态
Public gForm As Object      ' 由窗体在初始化时赋值，用于进度回调
Public gCancel As Boolean   ' 取消标志
Public gQuiet As Boolean    ' 静默模式：自动化测试时置 True，弹窗改为记录到 gLastError
Public gLastError As String
Public gDebug As Boolean    ' 调试开关：置 True 时把关键步骤写入 %TEMP%\PptImageExport-debug.log

' 调试日志（仅 gDebug = True 时写）
Public Sub Dbg(ByVal s As String)
    Dim f As Integer
    On Error Resume Next
    If Not gDebug Then Exit Sub
    f = FreeFile
    Open Environ$("TEMP") & "\PptImageExport-debug.log" For Append Access Write As #f
    Print #f, Format$(Now, "hh:nn:ss") & "  " & s
    Close #f
    On Error GoTo 0
End Sub


' ============================================================
'  一、功能区回调
' ============================================================

Public Sub Ribbon_ImageExport_Export(control As IRibbonControl)
    ShowExportDialog
End Sub

Public Sub ShowExportDialog()
    On Error GoTo EH
    If Application.Presentations.Count = 0 Then
        MsgBox "请先打开一个演示文稿。", vbExclamation, APP_TITLE
        Exit Sub
    End If
    Dim f As Object
    Set f = VBA.UserForms.Add("frmImageExport")
    f.Show vbModal
    Exit Sub
EH:
    gLastError = "打开导出面板失败：" & Err.Number & " " & Err.Description
    If Not gQuiet Then MsgBox gLastError, vbCritical, APP_TITLE
End Sub


' ============================================================
'  二、尺寸 / 分辨率换算
' ============================================================

' 幻灯片物理尺寸（英寸）。PageSetup.SlideWidth 单位是磅，72 磅 = 1 英寸
Public Sub GetSlideSizeInches(ByVal pres As Presentation, ByRef wIn As Double, ByRef hIn As Double)
    wIn = pres.PageSetup.SlideWidth / 72#
    hIn = pres.PageSetup.SlideHeight / 72#
    If wIn <= 0 Then wIn = 13.333
    If hIn <= 0 Then hIn = 7.5
End Sub

' DPI -> 像素
Public Sub DpiToPixels(ByVal dpi As Double, ByVal wIn As Double, ByVal hIn As Double, _
                       ByRef pxW As Long, ByRef pxH As Long)
    If dpi <= 0 Then dpi = 96
    ' 用 Int(x+0.5) 做四舍五入，避免 CLng 的银行家舍入在浮点噪声下出现 +1 偏差
    pxW = Int(dpi * wIn + 0.5)
    pxH = Int(dpi * hIn + 0.5)
    If pxW < 1 Then pxW = 1
    If pxH < 1 Then pxH = 1
End Sub

' 像素 -> 等效 DPI
Public Function PixelsToDpi(ByVal pxW As Long, ByVal wIn As Double) As Double
    If wIn > 0 Then PixelsToDpi = pxW / wIn Else PixelsToDpi = 0
End Function

' 长边像素 -> 按比例算出宽高
Public Sub LongEdgeToPixels(ByVal longEdge As Long, ByVal wIn As Double, ByVal hIn As Double, _
                            ByRef pxW As Long, ByRef pxH As Long)
    If longEdge < 16 Then longEdge = 16
    If wIn >= hIn Then
        pxW = longEdge
        pxH = Int(longEdge * hIn / wIn + 0.5)
    Else
        pxH = longEdge
        pxW = Int(longEdge * wIn / hIn + 0.5)
    End If
    If pxH < 1 Then pxH = 1
    If pxW < 1 Then pxW = 1
End Sub

' 按宽算高（保持比例）
Public Function HeightFromWidth(ByVal pxW As Long, ByVal wIn As Double, ByVal hIn As Double) As Long
    If wIn <= 0 Then wIn = 1
    HeightFromWidth = Int(pxW * hIn / wIn + 0.5)
    If HeightFromWidth < 1 Then HeightFromWidth = 1
End Function

' 按高算宽（保持比例）
Public Function WidthFromHeight(ByVal pxH As Long, ByVal wIn As Double, ByVal hIn As Double) As Long
    If hIn <= 0 Then hIn = 1
    WidthFromHeight = Int(pxH * wIn / hIn + 0.5)
    If WidthFromHeight < 1 Then WidthFromHeight = 1
End Function

' 数字转文本：去掉 VBA Format 对整数值多出来的小数点（288 -> "288" 而非 "288."）
Public Function FmtNum(ByVal v As Double) As String
    Dim s As String
    s = Format$(v, "0.###")
    Do While Len(s) > 0 And Right$(s, 1) = "."
        s = Left$(s, Len(s) - 1)
    Loop
    FmtNum = s
End Function

' 宽高比的简洁文本表示
Public Function AspectText(ByVal w As Long, ByVal h As Long) As String
    Dim a As Long, b As Long, t As Long, g As Long, rw As Long, rh As Long
    If w <= 0 Or h <= 0 Then AspectText = "-": Exit Function
    a = w: b = h
    Do While b <> 0
        t = a Mod b
        a = b
        b = t
    Loop
    g = a
    If g < 1 Then g = 1
    rw = w \ g: rh = h \ g
    If rw <= 50 And rh <= 50 Then
        AspectText = rw & ":" & rh
    Else
        AspectText = Format$(CDbl(w) / CDbl(h), "0.000") & ":1"
    End If
End Function

' 格式索引 -> 导出过滤器名
Public Function FormatFilter(ByVal fmtIdx As Long) As String
    Select Case fmtIdx
        Case 0: FormatFilter = "PNG"
        Case 1: FormatFilter = "JPG"
        Case 2: FormatFilter = "TIF"
        Case 3: FormatFilter = "BMP"
        Case 4: FormatFilter = "GIF"
        Case Else: FormatFilter = "PNG"
    End Select
End Function

' 格式索引 -> 扩展名
Public Function FormatExt(ByVal fmtIdx As Long) As String
    Select Case fmtIdx
        Case 0: FormatExt = "png"
        Case 1: FormatExt = "jpg"
        Case 2: FormatExt = "tif"
        Case 3: FormatExt = "bmp"
        Case 4: FormatExt = "gif"
        Case Else: FormatExt = "png"
    End Select
End Function


' ============================================================
'  三、导出范围解析
' ============================================================

' 返回当前视图所在幻灯片的序号（1 起）。取不到时返回 0
Public Function CurrentSlideIndex(ByVal pres As Presentation) As Long
    Dim s As Slide
    On Error Resume Next
    Set s = Application.ActiveWindow.View.Slide
    If Err.Number = 0 And Not s Is Nothing Then
        CurrentSlideIndex = s.SlideIndex
    End If
    Err.Clear
    On Error GoTo 0
End Function

' 取左侧缩略图中选中的幻灯片序号，写入 arr（1 起）。返回数量
Public Function SelectedSlideIndices(ByRef arr() As Long) As Long
    Dim sr As SlideRange, i As Long, n As Long
    On Error Resume Next
    Set sr = Application.ActiveWindow.Selection.SlideRange
    If Err.Number <> 0 Or sr Is Nothing Then
        Err.Clear
        On Error GoTo 0
        ReDim arr(0 To 0)
        SelectedSlideIndices = 0
        Exit Function
    End If
    n = sr.Count
    Err.Clear
    On Error GoTo 0
    If n < 1 Then
        ReDim arr(0 To 0)
        SelectedSlideIndices = 0
        Exit Function
    End If
    ReDim arr(1 To n)
    For i = 1 To n
        On Error Resume Next
        arr(i) = sr.Item(i).SlideIndex
        If Err.Number <> 0 Then arr(i) = 0
        Err.Clear
        On Error GoTo 0
    Next i
    SelectedSlideIndices = n
End Function

' 解析页码范围表达式，返回逗号分隔的页码串（如 "1,2,3"）；无匹配返回 ""
' 注意：不要让 VBA 函数直接返回数组（Boolean() 等），在部分宿主中会死锁
Public Function ParseSlideRangeToList(ByVal spec As String, ByVal maxIndex As Long) As String
    Dim used() As Boolean
    Dim n As Long, i As Long
    Dim out As String
    Dim parts() As String
    Dim tok As String, s1 As String, s2 As String
    Dim dashPos As Long
    Dim a As Long, b As Long, k As Long, t As Long

    If maxIndex < 1 Then maxIndex = 1
    n = maxIndex
    ReDim used(1 To n)

    spec = Replace(spec, "，", ",")
    spec = Replace(spec, "、", ",")
    spec = Replace(spec, "；", ",")
    spec = Replace(spec, ";", ",")
    spec = Replace(spec, "－", "-")
    spec = Replace(spec, "—", "-")
    spec = Replace(spec, "–", "-")
    spec = Replace(spec, "~", "-")
    spec = Replace(spec, "～", "-")
    spec = Replace(spec, " ", "")
    spec = Replace(spec, vbTab, "")

    parts = Split(spec, ",")
    For i = LBound(parts) To UBound(parts)
        tok = Trim$(parts(i))
        If Len(tok) > 0 Then
            If IsNumeric(tok) Then
                a = CLng(tok)
                If a >= 1 And a <= n Then used(a) = True
            Else
                dashPos = InStr(1, tok, "-")
                If dashPos > 1 Then
                    s1 = Left$(tok, dashPos - 1)
                    s2 = Mid$(tok, dashPos + 1)
                    s2 = Replace(s2, "-", "")
                    If IsNumeric(s1) And IsNumeric(s2) Then
                        a = CLng(s1): b = CLng(s2)
                        If a > b Then t = a: a = b: b = t
                        If a < 1 Then a = 1
                        If b > n Then b = n
                        For k = a To b
                            used(k) = True
                        Next k
                    End If
                End If
            End If
        End If
    Next i

    For i = 1 To n
        If used(i) Then
            If Len(out) > 0 Then out = out & ","
            out = out & CStr(i)
        End If
    Next i

    ParseSlideRangeToList = out
End Function

' 依据范围模式填充标志数组（ByRef 传出），返回选中页数
'   mode: 0=当前页  1=全部页  2=选中页  3=指定页码范围
Public Function ResolveSelection(ByVal pres As Presentation, ByVal mode As Long, _
                                 ByVal spec As String, _
                                 ByRef flags() As Boolean, _
                                 ByRef errMsg As String) As Long
    Dim n As Long, i As Long, cnt As Long
    Dim arr() As Long, c As Long, cur As Long
    Dim list As String, toks() As String

    n = pres.Slides.Count
    errMsg = ""
    If n < 1 Then n = 1
    ReDim flags(1 To n)

    Select Case mode
        Case 0
            cur = CurrentSlideIndex(pres)
            If cur < 1 Or cur > n Then cur = 1
            flags(cur) = True

        Case 1
            For i = 1 To n
                flags(i) = True
            Next i

        Case 2
            c = SelectedSlideIndices(arr)
            If c = 0 Then
                errMsg = "当前没有选中任何幻灯片。请先在左侧缩略图面板中选中要导出的页面。"
            Else
                For i = 1 To c
                    If arr(i) >= 1 And arr(i) <= n Then flags(arr(i)) = True
                Next i
            End If

        Case 3
            If Len(Trim$(spec)) = 0 Then
                errMsg = "请填写页码范围，例如：3-7, 12, 20-25"
            Else
                list = ParseSlideRangeToList(spec, n)
                If Len(list) > 0 Then
                    toks = Split(list, ",")
                    For i = LBound(toks) To UBound(toks)
                        flags(CLng(toks(i))) = True
                    Next i
                End If
            End If
    End Select

    cnt = 0
    For i = 1 To n
        If flags(i) Then cnt = cnt + 1
    Next i
    ResolveSelection = cnt
End Function

' 统计标志数组中为 True 的数量
Public Function CountSelected(ByRef flags() As Boolean) As Long
    Dim i As Long, c As Long
    On Error Resume Next
    For i = LBound(flags) To UBound(flags)
        If flags(i) Then c = c + 1
    Next i
    On Error GoTo 0
    CountSelected = c
End Function


' ============================================================
'  四、核心导出
' ============================================================

' 执行导出。返回是否全部成功；okCount 为成功张数；failMsg 记录失败明细
Public Function ExportImages(ByRef flags() As Boolean, ByVal outDir As String, _
                             ByVal prefix As String, ByVal fmtIdx As Long, _
                             ByVal pxW As Long, ByVal pxH As Long, _
                             ByRef okCount As Long, ByRef failMsg As String) As Boolean

    Dim pres As Presentation
    Dim filter As String, ext As String
    Dim i As Long, total As Long, done As Long
    Dim path As String, baseName As String

    Set pres = Application.ActivePresentation

    filter = FormatFilter(fmtIdx)
    ext = FormatExt(fmtIdx)

    If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"
    baseName = SafeFileName(prefix)
    If Len(baseName) = 0 Then baseName = "Slide"

    total = CountSelected(flags)
    okCount = 0
    failMsg = ""
    gCancel = False
    done = 0

    For i = LBound(flags) To UBound(flags)
        If gCancel Then Exit For
        If flags(i) Then
            done = done + 1
            path = outDir & baseName & "_" & Format$(i, "000") & "." & ext

            On Error Resume Next
            pres.Slides(i).Export path, filter, pxW, pxH
            If Err.Number <> 0 Then
                failMsg = failMsg & "  第 " & i & " 页：" & Err.Description & vbCrLf
                Err.Clear
            Else
                okCount = okCount + 1
            End If
            On Error GoTo 0

            NotifyProgress done, total, i
            DoEvents
        End If
    Next i

    ExportImages = (Len(failMsg) = 0)
End Function

Private Sub NotifyProgress(ByVal done As Long, ByVal total As Long, ByVal slideIdx As Long)
    On Error Resume Next
    If Not gForm Is Nothing Then gForm.SetProgress done, total, slideIdx
    On Error GoTo 0
End Sub


' ============================================================
'  五、文件与目录工具
' ============================================================

' 去掉文件名中的非法字符
Public Function SafeFileName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbCr, vbLf, vbTab)
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, CStr(bad(i)), "_")
    Next i
    s = Trim$(s)
    Do While Len(s) > 0
        If Right$(s, 1) = "." Or Right$(s, 1) = "_" Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    SafeFileName = s
End Function

' 逐级创建目录（MkDir 不支持多级），返回目录是否可用。
'
' 注意：存在性判断刻意**不用 VBA 内置的 Dir 函数**。
'      Dir 是宿主内置函数，在 .ppam 隐藏工程里保存后会被 VBE 规范化成小写
'      dir，重新加载时有被当成未声明标识符的风险 —— 而加载项报出来的只有
'      一句无法定位的「隐藏的模块中的编译错误」。
'      改用 Scripting.FileSystemObject（后期绑定），彻底规避这个问题。
Public Function EnsureDir(ByVal p As String) As Boolean
    Dim parts() As String, i As Long, cur As String
    Dim fso As Object
    On Error Resume Next

    p = Trim$(p)
    If Len(p) = 0 Then Exit Function
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Len(p) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function

    If fso.FolderExists(p) Then
        EnsureDir = True
        Exit Function
    End If

    If Left$(p, 2) = "\\" Then
        MkDir p
        EnsureDir = fso.FolderExists(p)
        Err.Clear
        Exit Function
    End If

    If InStr(p, ":") > 0 Then
        cur = Left$(p, 3)                 ' 例如 "C:\"
        parts = Split(Mid$(p, 4), "\")
    Else
        cur = ""
        parts = Split(p, "\")
    End If

    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            cur = cur & parts(i) & "\"
            If fso.FolderExists(cur) = False Then MkDir cur
        End If
    Next i

    EnsureDir = fso.FolderExists(p)
    Err.Clear
End Function

' 目录选择对话框。返回空字符串表示用户取消
Public Function PickFolder(ByVal startDir As String) As String
    Dim fd As Object, sh As Object, f As Object

    On Error Resume Next
    Set fd = Application.FileDialog(4)     ' msoFileDialogFolderPicker
    If Err.Number = 0 And Not fd Is Nothing Then
        On Error GoTo 0
        With fd
            .Title = "选择图片导出目录"
            If Len(startDir) > 0 Then .InitialFileName = startDir
            If .Show = -1 Then
                PickFolder = .SelectedItems(1)
            End If
        End With
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    ' 回退方案：Shell 对象（无 Office 组件依赖）
    On Error Resume Next
    Set sh = CreateObject("Shell.Application")
    Set f = sh.BrowseForFolder(0&, "选择图片导出目录", 0&)
    If Not f Is Nothing Then PickFolder = f.Self.Path
    Err.Clear
    On Error GoTo 0
End Function

' 在资源管理器中打开目录
Public Sub OpenFolder(ByVal p As String)
    On Error Resume Next
    If Len(p) = 0 Then Exit Sub
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    Shell "explorer.exe """ & p & """", vbNormalFocus
    On Error GoTo 0
End Sub

' 默认导出目录：演示文稿所在目录\导出图片
Public Function DefaultOutDir(ByVal pres As Presentation) As String
    Dim p As String
    On Error Resume Next
    p = pres.Path
    On Error GoTo 0
    If Len(p) = 0 Then
        p = Environ$("USERPROFILE") & "\Pictures"
    End If
    DefaultOutDir = p & "\导出图片"
End Function


' ============================================================
'  六、设置持久化（HKCU\Software\VB and VBA Program Settings\...）
' ============================================================

Public Sub SaveCfg(ByVal key As String, ByVal value As String)
    On Error Resume Next
    SaveSetting CFG_APP, "Config", key, value
    On Error GoTo 0
End Sub

Public Function LoadCfg(ByVal key As String, ByVal dflt As String) As String
    Dim v As String
    On Error Resume Next
    v = GetSetting(CFG_APP, "Config", key, dflt)
    If Err.Number <> 0 Then v = dflt
    Err.Clear
    On Error GoTo 0
    LoadCfg = v
End Function


' ============================================================
'  七、即时预览文案
' ============================================================

Public Function PreviewText(ByVal pxW As Long, ByVal pxH As Long, _
                            ByVal wIn As Double, ByVal hIn As Double, _
                            ByVal pageCount As Long, ByVal fmtIdx As Long) As String
    Dim s As String
    Dim dpi As Double
    dpi = PixelsToDpi(pxW, wIn)
    s = "输出 " & pxW & " × " & pxH & " px" & "   ·   " & _
        "比例 " & AspectText(pxW, pxH) & "   ·   " & _
        "等效 " & FmtNum(dpi) & " DPI" & "   ·   " & _
        FormatFilter(fmtIdx) & "   ·   " & _
        "共 " & pageCount & " 页"
    If pxW * CDbl(pxH) > BIG_PIXEL_WARN Then
        s = s & vbCrLf & "提示：单页超过 " & Format$(BIG_PIXEL_WARN / 1000000#, "0") & _
            " 百万像素，导出耗时较长且占用较多内存。"
    End If
    PreviewText = s
End Function
