
' ===== 自检代码（仅测试构建注入 frmImageExport，正式加载项不含） =====
Public Function SelfTest() As String
    Dim r As String
    On Error GoTo EH

    r = "UP=" & mUP & "|"
    r = r & "UNIT=" & IIf(mUP = 1, "point", "twip") & "|"
    r = r & "INSIDE=" & Me.InsideWidth & "x" & Me.InsideHeight & "|"
    r = r & "OUTER=" & Me.Width & "x" & Me.Height & "|"
    r = r & "SLIDES=" & mPres.Slides.Count & "|"

    mUpdating = False

    ' 1) 动态控件按名事件绑定：改宽应触发 txtW_Change 并自动换算高
    txtW.Text = "1234"
    DoEvents
    r = r & "W_CHANGE=" & IIf(Abs(CLng(Val(txtH.Text)) - 694) <= 1, "OK", "FAIL(" & txtH.Text & ")") & "|"

    ' 2) 预设下拉变更事件
    cboPreset.ListIndex = 3
    DoEvents
    r = r & "PRESET4K=" & IIf(InStr(1, lblPreview.Caption, "3840") > 0, "OK", "FAIL") & "|"

    ' 3) 格式下拉变更事件
    cboFmt.ListIndex = 1
    DoEvents
    r = r & "FMT_JPG=" & IIf(InStr(1, lblPreview.Caption, "JPG") > 0, "OK", "FAIL") & "|"
    cboFmt.ListIndex = 0
    DoEvents

    ' 4) 范围单选事件
    optAll.Value = True
    DoEvents
    r = r & "RANGE_ALL=" & IIf(InStr(1, lblPreview.Caption, "共 " & mPres.Slides.Count & " 页") > 0, "OK", "FAIL") & "|"

    ' 5) 页码范围
    optSpec.Value = True
    txtSpec.Text = "1,3"
    DoEvents
    r = r & "SPEC=" & lblPreview.Caption & "|"
    r = r & "SPEC_ENABLED=" & IIf(txtSpec.Enabled, "OK", "FAIL") & "|"

    ' 6) DPI 换算联动
    optCur.Value = True
    txtDpi.Text = "300"
    DoEvents
    r = r & "DPI300=" & txtW.Text & "x" & txtH.Text & "(" & lblPreview.Caption & ")"

    ' 复原
    txtDpi.Text = "150"
    DoEvents
    SelfTest = r
    Exit Function
EH:
    SelfTest = "ERROR: " & Err.Number & " " & Err.Description
End Function
