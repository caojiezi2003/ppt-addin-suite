
' ===== 自检代码（仅 --selftest 构建时注入 m_ImageReplace，正式加载项不含）=====
' 验证「图片替换」是否真的保住了全部属性：
'   逐项对比替换前后的 位置 / 尺寸 / 可见比例 / 裁剪分配 / 旋转 / 翻转 /
'   亮度 / 对比度 / 填充透明度 / 图层顺序 / 名称 / 替代文字 / 标题 /
'   超链接 / 动画 / 对象类型。
'   任一不一致都会在结果里以 DIFF(实际<>期望) 暴露出来，不靠肉眼判断。
'   注意裁剪比的是「比例」而不是「磅值」—— 换不同像素尺寸的图后磅值本来就该变，
'   详见「--- 裁剪 ---」那一段的说明。
'
' 另：把替换结果另存为 outPath，供 Python 侧解包做 OOXML 层面的独立核对。
'
' 诊断：整个过程的每一步都写进 ir-steps.txt / ir-trace.log，
'       若某个 COM 调用把宿主卡住，日志会恰好停在出问题的那一步。

' 注意：本文件是**追加到模块末尾**的，所以这里**不能出现任何模块级声明**
'      （Private/Public/Dim/Const ...）。VBA 规定模块级声明必须位于所有过程
'      之前，追加在过程之后的声明会被静默忽略，引用它的地方全部报
'      「变量未定义」—— 在 .ppam 里更只报一句无法定位的
'      「隐藏的模块中的编译错误」。需要常量就在过程内部用局部 Const。

' 日志路径硬编码在过程内部（避免模块级 Const 作用域带来的任何不确定性）
Private Sub IR_Step(ByVal s As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open "__TESTDIR__\ir-steps.txt" _
        For Append Access Write As #f
    Print #f, s
    Close #f
End Sub

' 无参入口：供 Application.Run 直接调用，排除"参数传递"这一变量。
Public Function ImageReplace_SelfTestAuto() As String
    Dim deck As String, img As String, out As String
    deck = "__TESTDIR__\output\_replace_in.pptx"
    img = "__TESTDIR__\output\_replace_new.png"
    out = "__TESTDIR__\output\_replace_out.pptx"
    IR_Step "auto-enter"
    ImageReplace_SelfTestAuto = ImageReplace_SelfTest(deck, img, out)
End Function

Public Function ImageReplace_SelfTest(ByVal deckPath As String, _
                                      ByVal newImgPath As String, _
                                      ByVal outPath As String) As String
    Dim deck   As Presentation
    Dim sld    As Slide
    Dim src    As Shape
    Dim aft    As Shape
    ' 基准值就地存放（刻意不用 UDT —— 见模块头说明）
    Dim bL As Single, bT As Single, bW As Single, bH As Single
    Dim bCL As Single, bCT As Single, bCR As Single, bCB As Single
    Dim bVisX As Single, bVisY As Single
    Dim bRot As Single, bBright As Single, bContrast As Single
    Dim bFillTrans As Single
    Dim bFlipH As Long, bFlipV As Long, bColorType As Long, bZ As Long
    Dim bName As String, bAlt As String, bTitle As String, bLink As String
    Dim bAnim As Boolean
    Dim em     As String
    Dim ok     As Boolean
    Dim r      As String
    Dim i      As Long

    Const IR_STEPS_FILE As String = "__TESTDIR__\ir-steps.txt"
    Const IR_TRACE_FILE As String = "__TESTDIR__\ir-trace.log"
    Const IR_TEST_TGT As String = "TGT_IMAGE"

    On Error Resume Next
    Kill IR_STEPS_FILE
    Kill IR_TRACE_FILE
    On Error GoTo EH

    IR_Step "enter"
    gIR_Debug = True
    gIR_TracePath = IR_TRACE_FILE

    IR_Step "open " & deckPath
    Set deck = Presentations.Open(deckPath, , , msoFalse)
    IR_Step "opened slides=" & deck.Slides.Count

    Set sld = deck.Slides(1)
    IR_Step "slide1 shapes=" & sld.Shapes.Count

    ' 定位测试目标图片
    Set src = Nothing
    For i = 1 To sld.Shapes.Count
        If sld.Shapes(i).Name = IR_TEST_TGT Then Set src = sld.Shapes(i)
    Next i
    IR_Step "found=" & IIf(src Is Nothing, "NO", "YES")
    If src Is Nothing Then
        deck.Close
        ImageReplace_SelfTest = "EXCEPTION=测试形状未找到：" & IR_TEST_TGT
        Exit Function
    End If

    ' 基准快照（用正式实现同一套逻辑，保证判据与实现同源）
    IR_Step "capture"
    If Not IR_Capture(src) Then
        deck.Close
        ImageReplace_SelfTest = "EXCEPTION=基准快照失败"
        Exit Function
    End If
    IR_Step "captured z=" & m_zOrder & " rot=" & m_rotation & _
            " cropL=" & m_cropL & " w=" & m_w

    bL = m_left
    bT = m_top
    bW = m_w
    bH = m_h
    bCL = m_cropL
    bCT = m_cropT
    bCR = m_cropR
    bCB = m_cropB
    bRot = m_rotation
    bFlipH = m_flipH
    bFlipV = m_flipV
    bBright = m_bright
    bContrast = m_contrast
    bColorType = m_colorType
    bFillTrans = m_fillTrans
    bZ = m_zOrder
    bName = m_shpName
    bAlt = m_shpAlt
    bTitle = m_shpTitle
    bLink = m_linkAddr
    bAnim = m_hasAnim

    ' 裁剪的正确不变量：可见比例（IR_Capture 已经把「未裁剪尺寸」量出来了，
    ' 直接相除即可，不需要再动源形状一次）
    If m_fullW > 0.0001 Then bVisX = m_w / m_fullW
    If m_fullH > 0.0001 Then bVisY = m_h / m_fullH
    IR_Step "baseline cropvis=" & IR_Fmt(bVisX) & "x" & IR_Fmt(bVisY) & _
            " cropLRTB=" & IR_Fmt(bCL) & "," & IR_Fmt(bCT) & "," & _
            IR_Fmt(bCR) & "," & IR_Fmt(bCB)

    ' 执行替换
    IR_Step "replace"
    em = ""
    ok = ImageReplace_ReplaceShape(src, newImgPath, em)
    IR_Step "replaced ok=" & ok & " em=[" & em & "]"

    r = "REPLACE=" & IIf(ok, "OK", "FAIL:" & em) & "|"
    If Not ok Then
        deck.Close
        ImageReplace_SelfTest = r
        Exit Function
    End If

    ' 按名称重新取回（原对象引用已随删除失效）
    IR_Step "refind"
    Set aft = Nothing
    For i = 1 To sld.Shapes.Count
        If sld.Shapes(i).Name = IR_TEST_TGT Then Set aft = sld.Shapes(i)
    Next i
    If aft Is Nothing Then
        IR_Step "refind FAILED"
        r = r & "REFIND=FAIL(替换后找不到同名形状)|"
        deck.Close
        ImageReplace_SelfTest = r
        Exit Function
    End If

    r = r & "REFIND=OK|"
    r = r & "TYPE=" & IIf(aft.Type = msoPicture, "PIC", "OTHER(" & aft.Type & ")") & "|"

    ' --- 几何 ---
    r = r & IR_CmpD("L", aft.Left, bL, 0.5)
    r = r & IR_CmpD("T", aft.Top, bT, 0.5)
    r = r & IR_CmpD("W", aft.Width, bW, 0.5)
    r = r & IR_CmpD("H", aft.Height, bH, 0.5)

    ' --- 裁剪 ---
    ' 判据口径说明（重要，别改回「直接比磅值」）：
    '     CropLeft 等虽然以「磅」暴露，但存进文件的是**比例**，分母是图片自身的原始尺寸。
    '     本插件的正确行为是「保住裁剪比例」，所以换上一张像素尺寸不同的新图之后，
    '     **磅值本来就该变**（实测 12 → 18）。直接比磅值会稳定误报 DIFF，
    '     这正是早前那版自检一直亮红灯的原因。
    '     改为比对两个与实现无关、且用户真正在意的量：
    '       ① 可见比例（裁剪后可见尺寸 / 未裁剪尺寸）—— 决定「看到图片的哪一块、多大」
    '       ② 四边分配比例（左 : 右 / 上 : 下）—— 决定裁剪怎么摊到各边
    '     绝对磅值由 Python 侧解包 OOXML 做第二重核对（那里的 srcRect 就是存储真值）。
    r = r & IR_CmpD("CROPX", IR_VisFracW(aft), bVisX, 0.002)
    r = r & IR_CmpD("CROPY", IR_VisFracH(aft), bVisY, 0.002)
    r = r & IR_CmpD("CROP_ALLOCX", IR_AllocRat(aft.PictureFormat.CropLeft, _
                                               aft.PictureFormat.CropRight), _
                    IR_AllocRat(bCL, bCR), 0.002)
    r = r & IR_CmpD("CROP_ALLOCY", IR_AllocRat(aft.PictureFormat.CropTop, _
                                               aft.PictureFormat.CropBottom), _
                    IR_AllocRat(bCT, bCB), 0.002)

    ' --- 变换 ---
    r = r & IR_CmpD("ROT", aft.Rotation, bRot, 0.01)
    r = r & IR_CmpL("FLIPH", aft.HorizontalFlip, bFlipH)
    r = r & IR_CmpL("FLIPV", aft.VerticalFlip, bFlipV)

    ' --- 图片属性 ---
    r = r & IR_CmpD("BRIGHT", aft.PictureFormat.Brightness, bBright, 0.01)
    r = r & IR_CmpD("CONTRAST", aft.PictureFormat.Contrast, bContrast, 0.01)
    r = r & IR_CmpL("COLORTYPE", aft.PictureFormat.ColorType, bColorType)
    r = r & IR_CmpD("FILLTRANS", aft.Fill.Transparency, bFillTrans, 0.01)

    ' --- 图层顺序（本次修复的核心之一）---
    r = r & IR_CmpL("ZORDER", aft.ZOrderPosition, bZ)

    ' --- 元数据 ---
    r = r & IR_CmpS("NAME", aft.Name, bName)
    r = r & IR_CmpS("ALT", aft.AlternativeText, bAlt)
    r = r & IR_CmpS("TITLE", aft.Title, bTitle)
    r = r & IR_CmpS("LINK", aft.ActionSettings(ppMouseClick).Hyperlink.Address, bLink)

    ' --- 动画 ---
    r = r & IR_CmpL("ANIM", aft.AnimationSettings.Animate, _
                    IIf(bAnim, msoTrue, msoFalse))

    ' 供 Python 侧解包核对
    IR_Step "save " & outPath
    deck.SaveAs outPath, 24
    IR_Step "saved"
    deck.Close
    IR_Step "closed"

    ImageReplace_SelfTest = r & "SAVED=OK"
    Exit Function

EH:
    IR_Step "EXCEPTION " & Err.Number & " " & Err.Description
    On Error Resume Next
    If Not deck Is Nothing Then deck.Close
    ImageReplace_SelfTest = "EXCEPTION=" & Err.Number & " " & Err.Description
End Function

Private Function IR_CmpD(ByVal label As String, ByVal got As Double, _
                         ByVal want As Double, ByVal tol As Double) As String
    IR_CmpD = label & "=" & IIf(Abs(got - want) <= tol, "OK", _
              "DIFF(" & IR_Fmt(got) & "<>" & IR_Fmt(want) & ")") & "|"
End Function

Private Function IR_CmpL(ByVal label As String, ByVal got As Long, _
                         ByVal want As Long) As String
    IR_CmpL = label & "=" & IIf(got = want, "OK", _
              "DIFF(" & got & "<>" & want & ")") & "|"
End Function

Private Function IR_CmpS(ByVal label As String, ByVal got As String, _
                         ByVal want As String) As String
    IR_CmpS = label & "=" & IIf(got = want, "OK", _
              "DIFF(" & got & "<>" & want & ")") & "|"
End Function

' ------------------------------------------------------------------
'  裁剪口径辅助过程
'
'  注意：本文件是**追加到模块末尾**的，所以下面只能出现「过程」，
'    不能出现任何模块级声明（Private/Public/Dim/Const ...），
'    否则会被静默忽略，引用处全部报「变量未定义」。
' ------------------------------------------------------------------

' 量出水平方向的「裁剪后可见比例」= 可见宽 / 未裁剪宽。
'
' 做法与 IR_Capture 里的「展开裁剪测完整矩形」完全一致：把裁剪临时清零，
' 读出的 Width 就是未裁剪显示宽；量完**立刻按原值写回**，保证状态不留痕迹
' （自检最后还要 SaveAs 供 Python 侧解包核对，必须干净）。
'
' 注意：写回一律先比对：每次多余的赋值都会再走一遍 PowerPoint 的浮点换算，
'   有机会在 EMU 末位抖一下（实测会造成 off.y 差 1 EMU）。
Private Function IR_VisFracW(ByRef shp As Shape) As Single
    Dim pf As Object
    Dim cL As Single, cR As Single
    Dim vw As Single, fw As Single

    On Error Resume Next
    Set pf = shp.PictureFormat
    If pf Is Nothing Then Exit Function

    vw = shp.Width
    cL = pf.CropLeft
    cR = pf.CropRight
    pf.CropLeft = 0
    pf.CropRight = 0
    fw = shp.Width

    If Abs(pf.CropLeft - cL) > 0.0001 Then pf.CropLeft = cL
    If Abs(pf.CropRight - cR) > 0.0001 Then pf.CropRight = cR
    If Abs(shp.Width - vw) > 0.0001 Then shp.Width = vw

    If fw > 0.0001 Then IR_VisFracW = vw / fw
    Err.Clear
End Function

' 垂直方向的「裁剪后可见比例」，写法同 IR_VisFracW。
Private Function IR_VisFracH(ByRef shp As Shape) As Single
    Dim pf As Object
    Dim cT As Single, cB As Single
    Dim vh As Single, fh As Single

    On Error Resume Next
    Set pf = shp.PictureFormat
    If pf Is Nothing Then Exit Function

    vh = shp.Height
    cT = pf.CropTop
    cB = pf.CropBottom
    pf.CropTop = 0
    pf.CropBottom = 0
    fh = shp.Height

    If Abs(pf.CropTop - cT) > 0.0001 Then pf.CropTop = cT
    If Abs(pf.CropBottom - cB) > 0.0001 Then pf.CropBottom = cB
    If Abs(shp.Height - vh) > 0.0001 Then shp.Height = vh

    If fh > 0.0001 Then IR_VisFracH = vh / fh
    Err.Clear
End Function

' 裁剪在两端之间的分配比例，例如 左:右 = 12:9 时返回 12/21 ≈ 0.5714。
' 用来核对裁剪是否按源形状的比例摊到各边，与「磅值该不该变」无关。
Private Function IR_AllocRat(ByVal a As Single, ByVal b As Single) As Single
    Dim t As Single
    On Error Resume Next
    t = a + b
    If t > 0.0001 Then IR_AllocRat = a / t
    Err.Clear
End Function

' ============================================================
'  组合（编组）内替换的自检
'
'  目的：证明图片位于组合里时，替换会同时保住两件事 ——
'    ① 组合自身的结构：成员数量、成员顺序、组合名、几何、旋转、翻转、图层
'    ② 图片自身的全部属性：沿用 ImageReplace_Describe 的口径
'  同时把结果另存为 outPath，供 Python 侧解包，从 OOXML 层面独立核对 grpSp。
'
'  注意：本段同样是**追加到模块末尾**的，只能出现过程，不能有任何模块级声明。
' ============================================================

' 在幻灯片顶层按名字找形状。
Private Function IR_FindTop(ByVal sld As Slide, ByVal nm As String) As Shape
    Dim i As Long
    On Error Resume Next
    For i = 1 To sld.Shapes.Count
        If sld.Shapes(i).Name = nm Then
            Set IR_FindTop = sld.Shapes(i)
            Err.Clear
            Exit Function
        End If
    Next i
    Err.Clear
End Function

' 在组合内递归按名字找形状（先扫本层，再下钻子组合）。
Private Function IR_FindInGroup(ByVal g As Shape, ByVal nm As String) As Shape
    Dim i As Long
    Dim it As Shape
    Dim hit As Shape
    On Error Resume Next
    If g Is Nothing Then Exit Function
    If g.Type <> msoGroup Then Exit Function
    For i = 1 To g.GroupItems.Count
        Set it = g.GroupItems.Item(i)
        If it.Name = nm Then
            Set IR_FindInGroup = it
            Err.Clear
            Exit Function
        End If
    Next i
    For i = 1 To g.GroupItems.Count
        Set it = g.GroupItems.Item(i)
        If it.Type = msoGroup Then
            Set hit = IR_FindInGroup(it, nm)
            If Not hit Is Nothing Then
                Set IR_FindInGroup = hit
                Err.Clear
                Exit Function
            End If
        End If
    Next i
    Err.Clear
End Function

' 组合的属性摘要（与图片的 ImageReplace_Describe 同一风格，便于逐字段比对）。
Private Function IR_GrpDesc(ByVal g As Shape) As String
    Dim r As String
    Dim i As Long
    On Error Resume Next
    If g Is Nothing Then
        IR_GrpDesc = "(Nothing)"
        Exit Function
    End If
    r = "NAME=" & g.Name
    r = r & "|TYPE=" & g.Type
    r = r & "|L=" & IR_Fmt(g.Left) & "|T=" & IR_Fmt(g.Top)
    r = r & "|W=" & IR_Fmt(g.Width) & "|H=" & IR_Fmt(g.Height)
    r = r & "|ROT=" & IR_Fmt(g.Rotation)
    r = r & "|FLIPH=" & g.HorizontalFlip & "|FLIPV=" & g.VerticalFlip
    r = r & "|Z=" & g.ZOrderPosition
    r = r & "|CNT=" & g.GroupItems.Count
    r = r & "|MEM="
    For i = 1 To g.GroupItems.Count
        If i > 1 Then r = r & ","
        r = r & g.GroupItems.Item(i).Name
    Next i
    Err.Clear
    IR_GrpDesc = r
End Function

' 把 "K=V|K=V" 形式的两个串逐字段比对，输出 K=OK / K=DIFF(前<>后)。
'
' skip 里列出的字段直接标 SKIP 不判——用于「这个字段本来就该变」的情形。
' 典型就是 CROP：新图的像素尺寸与旧图不同，裁剪的**磅值**必然要按比例换算，
' 照抄磅值反而是 bug。裁剪真正的不变量是**比例**，那由 Python 侧解包
' OOXML 比对 srcRect（crop.l/t/r/b）负责，口径更硬。
Private Function IR_CmpPairs(ByVal label As String, _
                             ByVal before As String, _
                             ByVal after As String, _
                             Optional ByVal skip As String = "") As String
    Dim pb() As String, pa() As String
    Dim i As Long
    Dim r As String
    Dim k As String, v1 As String, v2 As String
    Dim p1 As Long, p2 As Long

    pb = Split(before, "|")
    pa = Split(after, "|")
    If UBound(pb) <> UBound(pa) Then
        IR_CmpPairs = label & "=FIELDCOUNT(" & (UBound(pb) + 1) & "<>" & _
                      (UBound(pa) + 1) & ")|"
        Exit Function
    End If
    r = ""
    For i = 0 To UBound(pb)
        p1 = InStr(pb(i), "=")
        p2 = InStr(pa(i), "=")
        If p1 = 0 Or p2 = 0 Then
            r = r & label & ".BADFMT|"
        Else
            k = Left$(pb(i), p1 - 1)
            v1 = Mid$(pb(i), p1 + 1)
            v2 = Mid$(pa(i), p2 + 1)
            If Mid$(pa(i), 1, p2 - 1) <> k Then
                r = r & label & ".KEYDIFF|"
            ElseIf Len(skip) > 0 And InStr("," & skip & ",", "," & k & ",") > 0 Then
                r = r & label & "." & k & "=SKIP|"
            ElseIf v1 = v2 Then
                r = r & label & "." & k & "=OK|"
            Else
                r = r & label & "." & k & "=DIFF(" & v1 & "<>" & v2 & ")|"
            End If
        End If
    Next i
    IR_CmpPairs = r
End Function

' 无参入口：供 Application.Run 直接调用，排除参数传递这一变量。
Public Function ImageReplace_GroupSelfTestAuto() As String
    Const ROOT As String = "__TESTDIR__/output/"
    IR_Step "grp-auto-enter"
    ImageReplace_GroupSelfTestAuto = ImageReplace_GroupSelfTest( _
        ROOT & "_group_in.pptx", ROOT & "_replace_new.png", ROOT & "_group_out.pptx")
End Function

' 组合内替换自检主流程。
Public Function ImageReplace_GroupSelfTest(ByVal deckPath As String, _
                                           ByVal newImgPath As String, _
                                           ByVal outPath As String) As String
    Dim deck As Presentation
    Dim sld As Slide
    Dim grp As Shape
    Dim img As Shape
    Dim gB As String, iB As String
    Dim gA As String, iA As String
    Dim em As String
    Dim ok As Boolean
    Dim r As String
    Dim inGroupBefore As Boolean

    Const IR_TRACE_FILE As String = "__TESTDIR__/ir-grp-trace.log"

    On Error Resume Next
    Kill IR_TRACE_FILE
    Kill outPath
    On Error GoTo EH

    IR_Step "grp-enter"
    gIR_Debug = True
    gIR_TracePath = IR_TRACE_FILE

    ' 这里有窗口方式打开：后面的选区解析自检需要真实的 ActiveWindow /
    ' Selection —— 无窗口打开的演示文稿没有窗口，Shape.Select 会失败。
    Set deck = Presentations.Open(deckPath)
    IR_Step "grp-opened"
    If deck Is Nothing Then
        ImageReplace_GroupSelfTest = "GRP=OPEN_FAIL|"
        Exit Function
    End If
    Set sld = deck.Slides(1)

    Set grp = IR_FindTop(sld, "GRP_TARGET")
    If grp Is Nothing Then
        ImageReplace_GroupSelfTest = "GRP=NOT_FOUND|"
        GoTo FIN
    End If
    Set img = IR_FindInGroup(grp, "TGT_IMAGE")
    If img Is Nothing Then
        ImageReplace_GroupSelfTest = "GRP=IMG_NOT_FOUND|"
        GoTo FIN
    End If

    ' 前提确认：图片确实在组合里，否则本用例没有意义
    inGroupBefore = Not (IR_ParentGroup(img) Is Nothing)

    gB = IR_GrpDesc(grp)
    iB = ImageReplace_Describe(img)

    ' 替换。直接对组内图片调用即可 —— 函数内部会自动走组合流程。
    ok = ImageReplace_ReplaceShape(img, newImgPath, em)
    IR_Step "grp-replaced ok=" & ok

    ' 替换过程中组合被拆开又重建，旧引用已失效，按名字重新获取。
    Set grp = IR_FindTop(sld, "GRP_TARGET")
    Set img = IR_FindInGroup(grp, "TGT_IMAGE")

    gA = IR_GrpDesc(grp)
    iA = ImageReplace_Describe(img)

    r = "GRP.PRECOND=" & IIf(inGroupBefore, "OK", "NOT_INGROUP") & "|"
    r = r & "GRP.REPLACE=" & IIf(ok, "OK", "FAIL(" & em & ")") & "|"
    r = r & IR_CmpPairs("GRP", gB, gA)
    r = r & IR_CmpPairs("GIMG", iB, iA, "CROP")

    ' 另存供 Python 侧解包核对（先保存，再做选区测试，避免污染产物）
    deck.SaveAs outPath, 24
    r = r & "GRP.SAVED=OK|"
    IR_Step "grp-saved"

    ' 选区解析自检：覆盖用户实际操作方式 —— 单击组合 / 进入组合点选图片
    r = r & IR_SelTest(deck, sld, newImgPath)
    IR_Step "grp-sel-done"

    ImageReplace_GroupSelfTest = r
    gIR_Debug = False
    Exit Function

FIN:
    gIR_Debug = False
    On Error Resume Next
    deck.Close
    Err.Clear
    Exit Function

EH:
    IR_Step "grp-EH " & Err.Number & " " & Err.Description
    ImageReplace_GroupSelfTest = "GRP=EXCEPTION(" & Err.Number & " " & _
                                 Err.Description & ")|"
    gIR_Debug = False
End Function

' 选区解析自检 —— 覆盖用户真实的操作方式。
'
' 为什么单独测这一段：
'     用户报的问题就是「图片在编组内时无法识别」，报错是「当前类型 6」。
'     类型 6 来自 Selection.ShapeRange(1).Type —— 也就是说用户是**单击了整个组合**。
'     IR_ResolveTarget 的职责正是把这种选区翻译成「组合里的那张图片」，
'     所以它必须被真实地测到，而不能只测 ImageReplace_ReplaceShape。
Private Function IR_SelTest(ByVal deck As Presentation, _
                            ByVal sld As Slide, _
                            ByVal imgPath As String) As String
    Dim win As DocumentWindow
    Dim shp As Shape
    Dim em As String
    Dim ok As Boolean
    Dim r As String
    Dim grp As Shape
    Dim it As Shape
    Dim t1 As Shape, t2 As Shape, tg As Shape
    Dim got As String
    Dim i As Long

    On Error Resume Next
    Set grp = IR_FindTop(sld, "GRP_TARGET")
    If grp Is Nothing Then
        IR_SelTest = "SEL=NO_GROUP|"
        Err.Clear
        Exit Function
    End If
    If deck.Windows.Count = 0 Then
        IR_SelTest = "SEL=NO_WINDOW|"
        Err.Clear
        Exit Function
    End If
    Set win = deck.Windows(1)
    win.Activate

    r = "SEL.CNT=" & grp.GroupItems.Count & "|"

    ' ---- A. 单击整个组合（组内只有一张图片）----
    grp.Select
    ok = IR_ResolveTarget(win.Selection, shp, em)
    got = "?"
    If Not shp Is Nothing Then got = shp.Name
    r = r & "SEL.WHOLE=" & IIf(ok, "OK:" & got, "REJECT") & "|"

    ' ---- B. 双击进入组合后点选内部各个形状 ----
    ' 选中的是图片 → 应解析为那张图片；选中的不是图片（如矩形）→ 应拒绝。
    ' 这正是用户真实操作路径：手动点选组内指定的那张图片。
    For i = 1 To grp.GroupItems.Count
        Set it = grp.GroupItems.Item(i)
        it.Select
        win.Activate
        ok = IR_ResolveTarget(win.Selection, shp, em)
        got = "-"
        If Not shp Is Nothing Then got = shp.Name
        r = r & "SEL.IN." & it.Name & "=" & IIf(ok, "OK:" & got, "REJECT") & "|"
    Next i

    ' ---- C. 含两张图片的组合：应拒绝，并要求先进组合点选 ----
    Set t1 = sld.Shapes.AddPicture(imgPath, msoFalse, msoTrue, 20, 500, 60, 60)
    Set t2 = sld.Shapes.AddPicture(imgPath, msoFalse, msoTrue, 100, 500, 60, 60)
    Set tg = sld.Shapes.Range(Array(t1.Name, t2.Name)).Group()
    tg.Select
    ok = IR_ResolveTarget(win.Selection, shp, em)
    r = r & "SEL.MULTI=" & IIf(ok, "WRONG_RESOLVED", "REJECTED") & "|"
    tg.Delete

    Err.Clear
    IR_SelTest = r
End Function

