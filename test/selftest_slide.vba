
' ===== 自检代码（仅 --selftest 构建时注入 m_SlideExport，正式加载项不含） =====
' 验证「选页另存」核心逻辑：
'   · 无窗口打开生效（WithWindow:=msoFalse 不产生额外窗口）
'   · 按指定序号保留页面，导出页数正确
'   · 导出结果可被重新打开
Public Function SlideExport_SelfTest(ByVal deckPath As String, _
                                     ByVal outPath As String) As String
    Dim deck      As Presentation
    Dim chk       As Presentation
    Dim selIdx()  As Long
    Dim winBefore As Long
    Dim winAfter  As Long
    Dim delta     As Long
    Dim ok        As Boolean
    Dim got       As Long
    Dim em        As String
    Dim r         As String

    On Error GoTo EH

    ' A) 无窗口打开：Application.Windows.Count 不应变化
    winBefore = Application.Windows.Count
    Set deck = Presentations.Open(deckPath, , , msoFalse)
    winAfter = Application.Windows.Count
    delta = winAfter - winBefore

    r = "NOWIN=" & IIf(delta = 0, "OK", "FAIL(+" & delta & ")") & "|"
    r = r & "DECK=" & deck.Slides.Count & "|"

    ' B) 选第 2 / 4 / 5 页 → 期望保留 3 页
    ReDim selIdx(1 To 3)
    selIdx(1) = 2
    selIdx(2) = 4
    selIdx(3) = 5

    em = ""
    ok = SlideExport_Export(deck, selIdx, outPath, em)
    r = r & "EXPORT=" & IIf(ok, "OK", "FAIL:" & em) & "|"

    If ok Then
        Set chk = Presentations.Open(outPath, , , msoFalse)
        got = chk.Slides.Count
        chk.Close
        r = r & "KEPT=" & got & "|"
        r = r & "COUNT_MATCH=" & IIf(got = 3, "OK", "FAIL")
    End If

    deck.Close
    SlideExport_SelfTest = r
    Exit Function

EH:
    On Error Resume Next
    If Not deck Is Nothing Then deck.Close
    SlideExport_SelfTest = "EXCEPTION=" & Err.Number & " " & Err.Description
End Function
