Option Explicit

' ============================================================
'  frmImageExport — 导出设置面板（UserForm 代码模块）
'
'  说明：
'    控件在 UserForm_Initialize 中运行时创建，因此窗体不含
'    设计期控件，也就不需要 .frx 二进制资源，便于版本管理。
'    坐标单位通过 CalibrateUnits 在运行时实测校准（point 或 twip
'    都能自适应），布局一律以「磅」为单位书写。
' ============================================================

' ---------- 布局常量（单位：磅） ----------
Private Const CW As Double = 380      ' 客户区宽度
Private Const CH As Double = 283      ' 客户区高度
Private Const M As Double = 10        ' 左边距

' ---------- 状态 ----------
Private mUP As Double           ' 每磅对应的控件坐标单位数（运行时校准）
Private mUpdating As Boolean    ' 抑制事件递归
Private mBusy As Boolean        ' 导出进行中
Private mPres As Presentation
Private mWIn As Double
Private mHIn As Double
Private mPresetKeys() As String
Private mPresetCount As Long

' ---------- 控件引用 ----------
' 需要响应事件的控件必须声明为 WithEvents：
'   运行时用 Controls.Add 创建的控件不会按名字自动绑定事件，
'   只有挂到 WithEvents 变量上，<变量名>_<事件> 处理器才会被调用。
' 不需要事件的 Label 仍用 Object。
Private WithEvents optCur As MSForms.OptionButton
Private WithEvents optAll As MSForms.OptionButton
Private WithEvents optSel As MSForms.OptionButton
Private WithEvents optSpec As MSForms.OptionButton
Private WithEvents txtSpec As MSForms.TextBox
Private WithEvents cboPreset As MSForms.ComboBox
Private WithEvents txtDpi As MSForms.TextBox
Private WithEvents txtW As MSForms.TextBox
Private WithEvents txtH As MSForms.TextBox
Private WithEvents chkRatio As MSForms.CheckBox
Private WithEvents cboFmt As MSForms.ComboBox
Private WithEvents txtPrefix As MSForms.TextBox
Private WithEvents btnBrowse As MSForms.CommandButton
Private WithEvents btnExport As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton

Private lblRange As Object
Private lblSize As Object, lblPreset As Object
Private lblDpi As Object
Private lblW As Object, lblX As Object, lblH As Object
Private lblPx As Object
Private lblPreview As Object
Private lblSec3 As Object, lblFmtName As Object
Private lblDir As Object, txtDir As Object
Private lblPrefix As Object, lblHint As Object
Private lblStatus As Object


' ============================================================
'  初始化
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo EH

    Set mPres = Application.ActivePresentation
    m_ImageExport.GetSlideSizeInches mPres, mWIn, mHIn
    Set m_ImageExport.gForm = Me
    m_ImageExport.gCancel = False

    mUpdating = True
    mBusy = False

    Dbg "init: begin"
    Me.Caption = m_ImageExport.APP_TITLE & "  v" & m_ImageExport.APP_VERSION
    Me.Font.Name = "Microsoft YaHei UI"
    Me.Font.Size = 9
    Me.StartUpPosition = 1                  ' 相对宿主窗口居中
    Dbg "init: form props ok"

    CalibrateUnits
    Dbg "init: mUP=" & mUP
    BuildControls
    Dbg "init: controls built"
    BuildPresets
    Dbg "init: presets built"
    RestoreSettings
    Dbg "init: settings restored"
    mUpdating = False

    UpdateEnableState
    Dbg "init: enable state ok"
    RefreshPreview
    Dbg "init: preview ok"

    On Error Resume Next
    btnExport.SetFocus
    On Error GoTo 0
    Exit Sub

EH:
    m_ImageExport.gLastError = "界面初始化失败：" & Err.Number & " " & Err.Description
    Dbg "init: ERROR " & Err.Number & " " & Err.Description
    If Not m_ImageExport.gQuiet Then
        MsgBox m_ImageExport.gLastError, vbCritical, m_ImageExport.APP_TITLE
    End If
End Sub

' 实测控件坐标单位：用固定字号文本的 AutoSize 宽度反推每磅单位数
Private Sub CalibrateUnits()
    Dim t As Object, w As Double
    On Error Resume Next

    Set t = Me.Controls.Add("Forms.Label.1", "zzProbe")
    With t
        .Font.Name = "Microsoft YaHei UI"
        .Font.Size = 9
        .Font.Bold = False
        .AutoSize = True
        .Caption = "MMMMMMMMMM"
    End With
    w = t.Width
    Me.Controls.Remove "zzProbe"
    Err.Clear
    On Error GoTo 0

    ' 9pt 下 10 个大写 M 约 78 磅宽（含内边距）
    If w > 0 Then
        mUP = w / 78#
    Else
        mUP = 20#
    End If

    ' 归一化：控件单位只可能是「磅」(约 1) 或「twip」(约 20)
    If mUP < 3# Then
        mUP = 1#
    Else
        mUP = 20#
    End If
End Sub

' 磅 -> 控件坐标单位
Private Function P(ByVal pt As Double) As Long
    P = CLng(pt * mUP)
End Function

' 注意：绝不能在 UserForm_Initialize 中写 Me.InsideWidth / Me.InsideHeight
'     或 Me.Width / Me.Height —— 实测会导致 VBA 直接死锁。
'     窗体外形尺寸在构建阶段通过 VBE Designer 的设计期属性设定。


' ============================================================
'  控件构建
' ============================================================

Private Function AddCtl(ByVal progId As String, ByVal nm As String, _
                        ByVal l As Double, ByVal t As Double, _
                        ByVal w As Double, ByVal h As Double) As Object
    Dim c As Object
    Set c = Me.Controls.Add(progId, nm)
    c.Left = P(l)
    c.Top = P(t)
    c.Width = P(w)
    c.Height = P(h)
    On Error Resume Next
    c.Font.Name = "Microsoft YaHei UI"
    c.Font.Size = 9
    On Error GoTo 0
    Set AddCtl = c
End Function

Private Sub Sect(ByVal c As Object, ByVal txt As String)
    On Error Resume Next
    c.Caption = txt
    c.Font.Bold = True
    c.Font.Size = 10
    c.ForeColor = RGB(21, 79, 129)
    On Error GoTo 0
End Sub

Private Sub BuildControls()
    ' --- 1 导出范围 ---
    Set lblRange = AddCtl("Forms.Label.1", "lblRange", M, 8, 160, 13)
    Sect lblRange, "1  导出范围"

    Set optCur = AddCtl("Forms.OptionButton.1", "optCur", M, 25, 58, 16)
    optCur.Caption = "当前页"
    optCur.Value = True

    Set optAll = AddCtl("Forms.OptionButton.1", "optAll", 70, 25, 58, 16)
    optAll.Caption = "全部页"

    Set optSel = AddCtl("Forms.OptionButton.1", "optSel", 130, 25, 58, 16)
    optSel.Caption = "选中页"

    Set optSpec = AddCtl("Forms.OptionButton.1", "optSpec", 190, 25, 60, 16)
    optSpec.Caption = "页码范围"

    Set txtSpec = AddCtl("Forms.TextBox.1", "txtSpec", 252, 25, 118, 16)
    txtSpec.ControlTipText = "示例：3-7, 12, 20-25"
    Dbg "ctl: range row ok"

    ' --- 2 尺寸与分辨率 ---
    Set lblSize = AddCtl("Forms.Label.1", "lblSize", M, 46, 160, 13)
    Sect lblSize, "2  尺寸与分辨率"

    Set lblPreset = AddCtl("Forms.Label.1", "lblPreset", M, 63, 52, 17)
    lblPreset.Caption = "预设档位"

    Set cboPreset = AddCtl("Forms.ComboBox.1", "cboPreset", 64, 63, 175, 17)
    cboPreset.Style = 2                     ' 只读下拉
    cboPreset.ControlTipText = "选择常用尺寸档位，可再手动微调"

    Set lblDpi = AddCtl("Forms.Label.1", "lblDpi", 243, 63, 20, 17)
    lblDpi.Caption = "DPI"

    Set txtDpi = AddCtl("Forms.TextBox.1", "txtDpi", 265, 63, 45, 17)
    txtDpi.ControlTipText = "填 DPI 会自动换算像素：像素 = 英寸 x DPI"

    Set lblW = AddCtl("Forms.Label.1", "lblW", M, 86, 14, 17)
    lblW.Caption = "宽"

    Set txtW = AddCtl("Forms.TextBox.1", "txtW", 26, 86, 45, 17)
    txtW.ControlTipText = "输出宽度（像素）"

    Set lblX = AddCtl("Forms.Label.1", "lblX", 74, 86, 8, 17)
    lblX.Caption = "×"

    Set lblH = AddCtl("Forms.Label.1", "lblH", 84, 86, 14, 17)
    lblH.Caption = "高"

    Set txtH = AddCtl("Forms.TextBox.1", "txtH", 98, 86, 45, 17)
    txtH.ControlTipText = "输出高度（像素）"

    Set lblPx = AddCtl("Forms.Label.1", "lblPx", 145, 86, 12, 17)
    lblPx.Caption = "px"

    Set chkRatio = AddCtl("Forms.CheckBox.1", "chkRatio", 160, 86, 110, 17)
    chkRatio.Caption = "锁定宽高比"
    chkRatio.Value = True
    chkRatio.ControlTipText = "勾选后改宽或高会自动按幻灯片比例换算另一边"

    Set lblPreview = AddCtl("Forms.Label.1", "lblPreview", M, 110, 360, 30)
    lblPreview.ForeColor = RGB(0, 80, 160)
    lblPreview.WordWrap = True
    lblPreview.Caption = ""
    Dbg "ctl: size row ok"

    ' --- 3 输出设置 ---
    Set lblSec3 = AddCtl("Forms.Label.1", "lblSec3", M, 147, 160, 13)
    Sect lblSec3, "3  输出设置"

    Set lblFmtName = AddCtl("Forms.Label.1", "lblFmtName", M, 164, 45, 17)
    lblFmtName.Caption = "图片格式"

    Set cboFmt = AddCtl("Forms.ComboBox.1", "cboFmt", 57, 164, 85, 17)
    cboFmt.Style = 2
    cboFmt.AddItem "PNG"
    cboFmt.AddItem "JPG"
    cboFmt.AddItem "TIFF"
    cboFmt.AddItem "BMP"
    cboFmt.ListIndex = 0

    Set lblDir = AddCtl("Forms.Label.1", "lblDir", M, 187, 40, 17)
    lblDir.Caption = "输出目录"

    Set txtDir = AddCtl("Forms.TextBox.1", "txtDir", 52, 187, 262, 17)

    Set btnBrowse = AddCtl("Forms.CommandButton.1", "btnBrowse", 318, 187, 52, 17)
    btnBrowse.Caption = "浏览..."

    Set lblPrefix = AddCtl("Forms.Label.1", "lblPrefix", M, 210, 45, 17)
    lblPrefix.Caption = "文件名前缀"

    Set txtPrefix = AddCtl("Forms.TextBox.1", "txtPrefix", 57, 210, 85, 17)
    txtPrefix.ControlTipText = "文件名规则：前缀_页码.扩展名"

    Set lblHint = AddCtl("Forms.Label.1", "lblHint", 147, 210, 223, 17)
    lblHint.ForeColor = RGB(110, 110, 110)
    lblHint.Caption = ""

    Set lblStatus = AddCtl("Forms.Label.1", "lblStatus", M, 233, 360, 14)
    lblStatus.ForeColor = RGB(0, 100, 0)
    lblStatus.Caption = ""

    Set btnExport = AddCtl("Forms.CommandButton.1", "btnExport", M, 253, 92, 20)
    btnExport.Caption = "开始导出"
    btnExport.Font.Bold = True

    Set btnCancel = AddCtl("Forms.CommandButton.1", "btnCancel", 110, 253, 72, 20)
    btnCancel.Caption = "取消"
    btnCancel.Cancel = True
    Dbg "ctl: all controls ok"
End Sub

Private Sub BuildPresets()
    Dim i As Long
    mPresetCount = 10
    ReDim mPresetKeys(0 To mPresetCount - 1)

    mPresetKeys(0) = "DPI|96"
    mPresetKeys(1) = "EDGE|1920"
    mPresetKeys(2) = "EDGE|2560"
    mPresetKeys(3) = "EDGE|3840"
    mPresetKeys(4) = "EDGE|5120"
    mPresetKeys(5) = "EDGE|7680"
    mPresetKeys(6) = "DPI|150"
    mPresetKeys(7) = "DPI|300"
    mPresetKeys(8) = "DPI|600"
    mPresetKeys(9) = "CUSTOM|"

    cboPreset.AddItem "幻灯片原始尺寸（96 DPI）"
    cboPreset.AddItem "FHD 1080P（长边 1920 px）"
    cboPreset.AddItem "2K（长边 2560 px）"
    cboPreset.AddItem "4K（长边 3840 px）"
    cboPreset.AddItem "5K（长边 5120 px）"
    cboPreset.AddItem "8K（长边 7680 px）"
    cboPreset.AddItem "印刷级 150 DPI"
    cboPreset.AddItem "印刷级 300 DPI"
    cboPreset.AddItem "印刷级 600 DPI"
    cboPreset.AddItem "自定义"
    cboPreset.ListIndex = 3
End Sub


' ============================================================
'  数值与预览
' ============================================================

' DPI 显示：去掉整数的多余小数点
Private Function FmtDpi(ByVal d As Double) As String
    FmtDpi = m_ImageExport.FmtNum(d)
End Function

Private Function ParseNum(ByVal s As String) As Double
    Dim v As Double
    s = Trim$(s)
    s = Replace(s, " ", "")
    s = Replace(s, ",", "")
    s = Replace(s, "，", "")
    s = Replace(s, "px", "")
    s = Replace(s, "PX", "")
    s = Replace(s, "dpi", "")
    s = Replace(s, "DPI", "")
    If Len(s) = 0 Then ParseNum = 0: Exit Function
    On Error Resume Next
    v = CDbl(s)
    If Err.Number <> 0 Then v = 0
    Err.Clear
    On Error GoTo 0
    ParseNum = v
End Function

Private Function FmtIndex() As Long
    If cboFmt.ListIndex < 0 Then FmtIndex = 0 Else FmtIndex = cboFmt.ListIndex
End Function

Private Function CurrentRangeMode() As Long
    If optSpec.Value = True Then
        CurrentRangeMode = 3
    ElseIf optSel.Value = True Then
        CurrentRangeMode = 2
    ElseIf optAll.Value = True Then
        CurrentRangeMode = 1
    Else
        CurrentRangeMode = 0
    End If
End Function

Private Sub RefreshPreview()
    Dim w As Long, h As Long, cnt As Long
    Dim flags() As Boolean, errMsg As String
    Dim pfx As String

    w = CLng(ParseNum(txtW.Text))
    h = CLng(ParseNum(txtH.Text))
    If w < 1 Then w = 1
    If h < 1 Then h = 1

    cnt = m_ImageExport.ResolveSelection(mPres, CurrentRangeMode(), txtSpec.Text, flags, errMsg)

    lblPreview.Caption = m_ImageExport.PreviewText(w, h, mWIn, mHIn, cnt, FmtIndex())

    pfx = m_ImageExport.SafeFileName(txtPrefix.Text)
    If Len(pfx) = 0 Then pfx = "Slide"
    lblHint.Caption = "命名规则：" & pfx & "_001." & m_ImageExport.FormatExt(FmtIndex())
End Sub

Private Sub UpdateEnableState()
    On Error Resume Next
    Dim on_ As Boolean
    on_ = (optSpec.Value = True)
    txtSpec.Enabled = on_
    If on_ Then
        txtSpec.BackColor = RGB(255, 255, 255)
    Else
        txtSpec.BackColor = RGB(240, 240, 240)
    End If
    On Error GoTo 0
End Sub

Private Sub SetBusyUi()
    On Error Resume Next
    Dim e As Boolean
    e = Not mBusy
    btnExport.Enabled = e
    btnBrowse.Enabled = e
    cboPreset.Enabled = e
    cboFmt.Enabled = e
    txtDpi.Enabled = e
    txtW.Enabled = e
    txtH.Enabled = e
    txtDir.Enabled = e
    txtPrefix.Enabled = e
    chkRatio.Enabled = e
    optCur.Enabled = e
    optAll.Enabled = e
    optSel.Enabled = e
    optSpec.Enabled = e
    If mBusy Then
        btnCancel.Caption = "中止"
    Else
        btnCancel.Caption = "取消"
    End If
    On Error GoTo 0
End Sub

' 供引擎回调，更新进度
Public Sub SetProgress(ByVal done As Long, ByVal total As Long, ByVal slideIdx As Long)
    On Error Resume Next
    lblStatus.Caption = "正在导出  " & done & " / " & total & "   （第 " & slideIdx & " 页）"
    On Error GoTo 0
End Sub


' ============================================================
'  事件处理
' ============================================================

Private Sub RangeChanged()
    If mUpdating Then Exit Sub
    UpdateEnableState
    RefreshPreview
End Sub

Private Sub optCur_Click()
    RangeChanged
End Sub

Private Sub optAll_Click()
    RangeChanged
End Sub

Private Sub optSel_Click()
    RangeChanged
End Sub

Private Sub optSpec_Click()
    RangeChanged
End Sub

Private Sub txtSpec_Change()
    If mUpdating Then Exit Sub
    RefreshPreview
End Sub

Private Sub cboPreset_Change()
    If mUpdating Then Exit Sub
    ApplyPreset
End Sub

Private Sub ApplyPreset()
    Dim i As Long, kind As String, val As String, pos As Long
    Dim w As Long, h As Long, d As Double

    i = cboPreset.ListIndex
    If i < 0 Then Exit Sub

    kind = mPresetKeys(i)
    pos = InStr(1, kind, "|")
    val = Mid$(kind, pos + 1)
    kind = Left$(kind, pos - 1)

    mUpdating = True
    Select Case kind
        Case "DPI"
            d = ParseNum(val)
            If d <= 0 Then d = 96
            m_ImageExport.DpiToPixels d, mWIn, mHIn, w, h
            txtDpi.Text = FmtDpi(d)
            txtW.Text = CStr(w)
            txtH.Text = CStr(h)
        Case "EDGE"
            m_ImageExport.LongEdgeToPixels CLng(ParseNum(val)), mWIn, mHIn, w, h
            txtW.Text = CStr(w)
            txtH.Text = CStr(h)
            txtDpi.Text = FmtDpi(m_ImageExport.PixelsToDpi(w, mWIn))
        Case Else
            ' CUSTOM：保留当前数值
    End Select
    mUpdating = False

    RefreshPreview
End Sub

Private Sub SetPresetToCustom()
    Dim i As Long
    For i = 0 To mPresetCount - 1
        If Left$(mPresetKeys(i), 7) = "CUSTOM|" Then
            If cboPreset.ListIndex <> i Then
                mUpdating = True
                cboPreset.ListIndex = i
                mUpdating = False
            End If
            Exit For
        End If
    Next i
End Sub

Private Sub SyncDpiFromPixels()
    Dim w As Long, h As Long, d As Double
    w = CLng(ParseNum(txtW.Text))
    h = CLng(ParseNum(txtH.Text))
    If w < 1 Or h < 1 Then Exit Sub
    d = m_ImageExport.PixelsToDpi(w, mWIn)
    txtDpi.Text = FmtDpi(d)
End Sub

Private Sub txtDpi_Change()
    Dim d As Double, w As Long, h As Long
    If mUpdating Then Exit Sub
    d = ParseNum(txtDpi.Text)
    If d <= 0 Then Exit Sub

    m_ImageExport.DpiToPixels d, mWIn, mHIn, w, h
    mUpdating = True
    txtW.Text = CStr(w)
    txtH.Text = CStr(h)
    SetPresetToCustom
    mUpdating = False
    RefreshPreview
End Sub

Private Sub txtW_Change()
    Dim w As Long
    If mUpdating Then Exit Sub
    w = CLng(ParseNum(txtW.Text))
    If w < 1 Then Exit Sub

    mUpdating = True
    If chkRatio.Value = True Then
        txtH.Text = CStr(m_ImageExport.HeightFromWidth(w, mWIn, mHIn))
    End If
    SyncDpiFromPixels
    SetPresetToCustom
    mUpdating = False
    RefreshPreview
End Sub

Private Sub txtH_Change()
    Dim h As Long, w As Long
    If mUpdating Then Exit Sub
    h = CLng(ParseNum(txtH.Text))
    If h < 1 Then Exit Sub

    mUpdating = True
    If chkRatio.Value = True Then
        txtW.Text = CStr(m_ImageExport.WidthFromHeight(h, mWIn, mHIn))
    End If
    SyncDpiFromPixels
    SetPresetToCustom
    mUpdating = False
    RefreshPreview
End Sub

Private Sub chkRatio_Click()
    If mUpdating Then Exit Sub
    If chkRatio.Value = True Then
        Dim w As Long
        w = CLng(ParseNum(txtW.Text))
        If w > 0 Then
            mUpdating = True
            txtH.Text = CStr(m_ImageExport.HeightFromWidth(w, mWIn, mHIn))
            mUpdating = False
        End If
    End If
    RefreshPreview
End Sub

Private Sub cboFmt_Change()
    If mUpdating Then Exit Sub
    RefreshPreview
End Sub

Private Sub txtPrefix_Change()
    If mUpdating Then Exit Sub
    RefreshPreview
End Sub

Private Sub btnBrowse_Click()
    Dim p As String
    p = m_ImageExport.PickFolder(txtDir.Text)
    If Len(p) > 0 Then txtDir.Text = p
End Sub


' ============================================================
'  导出主流程
' ============================================================

Private Sub btnExport_Click()
    Dim flags() As Boolean, errMsg As String
    Dim outDir As String, prefix As String
    Dim w As Long, h As Long, cnt As Long, okCount As Long
    Dim failMsg As String
    Dim msg As String
    Dim px As Double

    If mBusy Then Exit Sub

    outDir = Trim$(txtDir.Text)
    If Len(outDir) = 0 Then
        MsgBox "请先选择输出目录。", vbExclamation, m_ImageExport.APP_TITLE
        Exit Sub
    End If
    If Not m_ImageExport.EnsureDir(outDir) Then
        MsgBox "无法创建或访问输出目录：" & vbCrLf & outDir, vbExclamation, m_ImageExport.APP_TITLE
        Exit Sub
    End If

    w = CLng(ParseNum(txtW.Text))
    h = CLng(ParseNum(txtH.Text))
    If w < 1 Or h < 1 Then
        MsgBox "请填写有效的像素尺寸。", vbExclamation, m_ImageExport.APP_TITLE
        Exit Sub
    End If
    If w > 20000 Or h > 20000 Then
        MsgBox "单边像素不建议超过 20000，请下调尺寸后重试。", vbExclamation, m_ImageExport.APP_TITLE
        Exit Sub
    End If

    If chkRatio.Value = True Then
        mUpdating = True
        h = m_ImageExport.HeightFromWidth(w, mWIn, mHIn)
        txtH.Text = CStr(h)
        mUpdating = False
    End If

    cnt = m_ImageExport.ResolveSelection(mPres, CurrentRangeMode(), txtSpec.Text, flags, errMsg)
    If Len(errMsg) > 0 Then
        MsgBox errMsg, vbExclamation, m_ImageExport.APP_TITLE
        Exit Sub
    End If

    cnt = m_ImageExport.CountSelected(flags)
    If cnt = 0 Then
        MsgBox "没有匹配到任何幻灯片，请检查导出范围。", vbExclamation, m_ImageExport.APP_TITLE
        Exit Sub
    End If

    px = CDbl(w) * CDbl(h)
    If px > 40000000# Then
        If MsgBox("单页输出约 " & Format$(px / 1000000#, "0") & " 百万像素，" & vbCrLf & _
                  "导出耗时较长且占用较多内存。" & vbCrLf & vbCrLf & "是否继续？", _
                  vbQuestion + vbYesNo, m_ImageExport.APP_TITLE) <> vbYes Then
            Exit Sub
        End If
    End If

    prefix = m_ImageExport.SafeFileName(txtPrefix.Text)
    If Len(prefix) = 0 Then prefix = "Slide"

    SaveSettings outDir, prefix, w, h

    mBusy = True
    m_ImageExport.gCancel = False
    SetBusyUi
    lblStatus.Caption = "正在导出  0 / " & cnt & " ..."
    DoEvents

    m_ImageExport.ExportImages flags, outDir, prefix, FmtIndex(), w, h, okCount, failMsg

    mBusy = False
    SetBusyUi

    msg = "导出完成。" & vbCrLf & vbCrLf & _
          "成功：" & okCount & " / " & cnt & " 张" & vbCrLf & _
          "尺寸：" & w & " × " & h & " px" & vbCrLf & _
          "目录：" & outDir
    If Len(failMsg) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "以下页面导出失败：" & vbCrLf & failMsg
    End If
    If m_ImageExport.gCancel Then
        msg = msg & vbCrLf & vbCrLf & "（已按中止，未完成剩余页面）"
    End If

    lblStatus.Caption = "完成  " & okCount & " / " & cnt & " 张"

    If MsgBox(msg & vbCrLf & vbCrLf & "是否打开输出目录？", _
              IIf(Len(failMsg) > 0, vbExclamation, vbInformation) + vbYesNo, _
              m_ImageExport.APP_TITLE) = vbYes Then
        m_ImageExport.OpenFolder outDir
    End If
End Sub

Private Sub btnCancel_Click()
    If mBusy Then
        m_ImageExport.gCancel = True
        lblStatus.Caption = "正在中止，请稍候..."
        DoEvents
    Else
        Unload Me
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mBusy Then
        Cancel = True
        m_ImageExport.gCancel = True
    End If
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    Set m_ImageExport.gForm = Nothing
    On Error GoTo 0
End Sub


' ============================================================
'  设置持久化
' ============================================================

Private Sub SaveSettings(ByVal dir As String, ByVal prefix As String, ByVal w As Long, ByVal h As Long)
    m_ImageExport.SaveCfg "Dir", dir
    m_ImageExport.SaveCfg "Prefix", prefix
    m_ImageExport.SaveCfg "W", CStr(w)
    m_ImageExport.SaveCfg "H", CStr(h)
    m_ImageExport.SaveCfg "Fmt", CStr(FmtIndex())
    m_ImageExport.SaveCfg "Ratio", IIf(chkRatio.Value = True, "1", "0")
    m_ImageExport.SaveCfg "Range", CStr(CurrentRangeMode())
    m_ImageExport.SaveCfg "Spec", txtSpec.Text
End Sub

Private Sub RestoreSettings()
    Dim fi As Long, rm As Long, w As Long, h As Long

    mUpdating = True

    txtDir.Text = m_ImageExport.LoadCfg("Dir", m_ImageExport.DefaultOutDir(mPres))
    txtPrefix.Text = m_ImageExport.LoadCfg("Prefix", "Slide")
    txtSpec.Text = m_ImageExport.LoadCfg("Spec", "")
    chkRatio.Value = (m_ImageExport.LoadCfg("Ratio", "1") = "1")

    fi = CLng(Val(m_ImageExport.LoadCfg("Fmt", "0")))
    If fi < 0 Or fi > cboFmt.ListCount - 1 Then fi = 0
    cboFmt.ListIndex = fi

    rm = CLng(Val(m_ImageExport.LoadCfg("Range", "0")))
    Select Case rm
        Case 1: optAll.Value = True
        Case 2: optSel.Value = True
        Case 3: optSpec.Value = True
        Case Else: optCur.Value = True
    End Select

    w = CLng(Val(m_ImageExport.LoadCfg("W", "0")))
    h = CLng(Val(m_ImageExport.LoadCfg("H", "0")))
    If w < 1 Or h < 1 Then
        m_ImageExport.LongEdgeToPixels 3840, mWIn, mHIn, w, h
        cboPreset.ListIndex = 3
    Else
        cboPreset.ListIndex = 9
    End If
    txtW.Text = CStr(w)
    txtH.Text = CStr(h)
    txtDpi.Text = FmtDpi(m_ImageExport.PixelsToDpi(w, mWIn))
    Dbg "restore: done w=" & w & " h=" & h

    mUpdating = False
End Sub
