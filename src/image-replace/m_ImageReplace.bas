Attribute VB_Name = "m_ImageReplace"
Option Explicit

' ============================================================
'  图片替换 —— PPT 增强插件套件 / 插件
'
'  用另一张图片原位替换当前选中的图片。
'
'  历史实现（本套件之前的版本）的缺陷：
'      shp.Delete                                  ' 原形状被整个删除
'      Set shp = sld.Shapes.AddPicture(...)        ' 新形状追加到 Z 序最顶层
'      shp.PictureFormat.CropLeft = cropL          ' 只还原了裁剪参数
'
'  「删除 + 重新插入」这个做法，注定了只能还原「位置 + 尺寸 + 裁剪」三样，
'  其余属性全部丢失：
'      · 图层顺序（新图永远跑到最顶层，遮挡关系被破坏）
'      · 透明度 / 亮度 / 对比度 / 透明色
'      · 阴影 / 发光 / 柔化边缘 / 边框 / 艺术效果（PickUp/Apply 整块格式）
'      · 旋转 / 水平翻转 / 垂直翻转
'      · 名称 / 替代文字 / 标题 / 超链接
'      · 动画
'
'  本实现改为  「快照 → 迁移 → 还原」  三步：
'      ① 快照   记录全部可还原属性；并临时展开裁剪测得
'               「未裁剪完整矩形」（原实现的思路，实测最准），随后立刻恢复
'      ② 迁移   原形状 PickUp / PickupAnimation → 先插新图（失败不丢原图）
'               → 删原图 → 新形状 Apply / ApplyAnimation
'      ③ 还原   几何 → 裁剪 → 变换 → 图片属性 → 元数据 → 图层顺序 → 位置终校
'
'  图片位于「组合（编组）」内时，走另一条等价路径（见下方「③ 组合支持」）：
'      快照组合 → 变换归零 → 解组 → 【复用上面这套替换逻辑】 → 重组 → 复原组合属性
'      这样可以精确选中组合里的某一张图片替换，同时把组合本身完整保留。
'
'  套件规范化命名：
'      模块        m_ImageReplace
'      Ribbon 回调 Ribbon_ImageReplace_Replace
'      主流程      ImageReplace_Run
'      核心        ImageReplace_ReplaceShape  （与界面无关，便于自动化验证）
'      属性摘要    ImageReplace_Describe
' ============================================================

' ============================================================
'  模块级声明
'
'  ★ 放置位置是硬性要求（本插件踩过的坑，务必遵守）★
'     VBA 规定：**模块级声明必须全部位于任何过程之前**。
'     写在过程之后的 Private / Public 声明不会报语法错，
'     而是被**静默忽略** —— 于是所有引用它的过程都变成
'     「变量未定义」编译错误。
'     在 .ppam 加载项里更棘手：只报一句含糊的
'     「隐藏的模块中的编译错误: <模块名>」，既不指行也不指名字。
'     因此：本模块新增的任何模块级变量 / 常量，一律加在这一段里。
'
'  说明：属性快照刻意**不使用用户定义类型（UDT）**，而用一组扁平变量。
'        扁平变量不经过任何复合类型系统，最不容易触发宿主怪癖。
' ============================================================

' ---- ① 属性快照：几何（裁剪后的可见矩形）----
Private m_left As Single
Private m_top As Single
Private m_w As Single
Private m_h As Single
' ---- 未裁剪完整矩形（展开裁剪后测得）----
Private m_fullLeft As Single
Private m_fullTop As Single
Private m_fullW As Single
Private m_fullH As Single
' ---- 裁剪（磅）----
Private m_cropL As Single
Private m_cropT As Single
Private m_cropR As Single
Private m_cropB As Single
' ---- 变换 ----
Private m_rotation As Single
Private m_flipH As Long
Private m_flipV As Long
' ---- 图片属性 ----
Private m_bright As Single
Private m_contrast As Single
Private m_transColor As Long
Private m_transBg As Long
Private m_colorType As Long
Private m_fillTrans As Single
' ---- 图层 ----
Private m_zOrder As Long
' ---- 元数据 ----
Private m_shpName As String
Private m_shpAlt As String
Private m_shpTitle As String
Private m_hasLink As Boolean
Private m_linkAddr As String
Private m_linkSub As String
' ---- 动画 ----
Private m_hasAnim As Boolean
' ---- 快照是否有效 ----
Private m_ok As Boolean

' ---- 组合（编组）支持：链式处理所需的临时状态 ----
' m_names   「临时名 → 原名」映射；解组会让原本分属不同组合的同名形状撞车，
'            所以先把整棵树改成全局唯一名，事后再按这张表还原。
' m_chain   祖先组合链的快照（含每层的名字/位置/旋转翻转/Z 序/成员），
'            索引 1 = 最外层，末项 = 图片的直接父组合。重组时逐层回放。
Private m_names As Object
Private m_chain As Collection
Private m_renameIdx As Long

' ---- 步骤追踪（排查用，默认关闭，无任何副作用）----
' 置 gIR_Debug = True 后，替换过程的每一步都会追加写入日志文件
' （路径取 gIR_TracePath；留空则回退到 %TEMP%\PptImageReplace-trace.log）。
' 若某个 COM 调用把宿主卡住，日志会恰好停在出问题的那一步。
Public gIR_Debug As Boolean
Public gIR_TracePath As String
Private m_tracePath As String

' ------------------------------------------------------------
'  Ribbon 回调入口
' ------------------------------------------------------------
Public Sub Ribbon_ImageReplace_Replace(control As IRibbonControl)
    On Error GoTo EH
    ImageReplace_Run
    Exit Sub
EH:
    ReportError "图片替换", Err.Number, Err.Description
End Sub

' ------------------------------------------------------------
'  交互主流程：校验选区 → 选新图 → 替换
' ------------------------------------------------------------
Private Sub ImageReplace_Run()
    Dim sel        As Selection
    Dim shp        As Shape
    Dim newImgPath As String
    Dim errMsg     As String
    Dim inGroup    As Boolean

    On Error GoTo EH

    If ActiveWindow Is Nothing Then
        MsgBox "请先打开一份演示文稿。", vbExclamation, "图片替换"
        Exit Sub
    End If

    On Error Resume Next
    Set sel = ActiveWindow.Selection
    On Error GoTo EH
    If sel Is Nothing Then
        MsgBox "请先在幻灯片上选中一张图片。", vbExclamation, "图片替换"
        Exit Sub
    End If

    If sel.Type <> ppSelectionShapes Then
        MsgBox "请先在幻灯片上选中一张图片，再点「原位替换图片」。", _
               vbExclamation, "图片替换"
        Exit Sub
    End If

    ' 解析选区：直接选中的图片、双击进入组合后点选的图片、或内含唯一图片的组合
    If Not IR_ResolveTarget(sel, shp, errMsg) Then
        MsgBox errMsg, vbExclamation, "图片替换"
        Exit Sub
    End If
    inGroup = Not (IR_ParentGroup(shp) Is Nothing)

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "选择用于替换的新图片"
        .Filters.Clear
        .Filters.Add "图片文件", "*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff;*.gif", 1
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        newImgPath = CStr(.SelectedItems(1))
    End With

    If Not ImageReplace_ReplaceShape(shp, newImgPath, errMsg) Then
        MsgBox errMsg, vbCritical, "替换失败"
        Exit Sub
    End If

    If inGroup Then
        MsgBox "图片已在组合内原位替换。" & vbCrLf & vbCrLf & _
               "组合的结构、位置与旋转翻转，以及图片的位置、尺寸、裁剪、" & vbCrLf & _
               "亮度对比度、透明度、图层顺序、名称与超链接、动画均已保留。", _
               vbInformation, "完成"
    Else
        MsgBox "图片已原位替换。" & vbCrLf & vbCrLf & _
               "位置、尺寸、裁剪、旋转与翻转、亮度对比度、透明度、" & vbCrLf & _
               "阴影 / 边框 / 艺术效果、图层顺序、名称与超链接、动画均已保留。", _
               vbInformation, "完成"
    End If
    Exit Sub

EH:
    ReportError "图片替换", Err.Number, Err.Description
End Sub

Private Sub IR_Trace(ByVal s As String)
    Dim f As Integer
    On Error Resume Next
    If Not gIR_Debug Then Exit Sub
    If Len(m_tracePath) = 0 Then
        m_tracePath = gIR_TracePath
        If Len(m_tracePath) = 0 Then
            m_tracePath = Environ$("TEMP") & "\PptImageReplace-trace.log"
        End If
    End If
    f = FreeFile
    Open m_tracePath For Append Access Write As #f
    Print #f, s
    Close #f
End Sub

Private Sub IR_ClearSnap()
    m_ok = False
    m_left = 0
    m_top = 0
    m_w = 0
    m_h = 0
    m_fullLeft = 0
    m_fullTop = 0
    m_fullW = 0
    m_fullH = 0
    m_cropL = 0
    m_cropT = 0
    m_cropR = 0
    m_cropB = 0
    m_rotation = 0
    m_flipH = msoFalse
    m_flipV = msoFalse
    m_bright = 0
    m_contrast = 0
    m_transColor = 0
    m_transBg = msoFalse
    m_colorType = msoPictureAutomatic
    m_fillTrans = 0
    m_zOrder = 1
    m_shpName = ""
    m_shpAlt = ""
    m_shpTitle = ""
    m_hasLink = False
    m_linkAddr = ""
    m_linkSub = ""
    m_hasAnim = False
End Sub

' 记录源形状的全部可还原属性。调用前后形状外观保持一致。
Private Function IR_Capture(ByVal shp As Shape) As Boolean
    Dim pf     As Object
    Dim cropL  As Single, cropT As Single, cropR As Single, cropB As Single
    Dim fl     As Single, ft As Single, fw As Single, fh As Single

    IR_ClearSnap
    IR_Capture = False
    IR_Trace "capture:enter"
    If shp Is Nothing Then Exit Function

    On Error Resume Next
    Set pf = shp.PictureFormat
    IR_Trace "capture:picformat"
    If pf Is Nothing Then
        Err.Clear
        Exit Function
    End If

    m_left = shp.Left
    m_top = shp.Top
    m_w = shp.Width
    m_h = shp.Height
    IR_Trace "capture:geometry"

    cropL = pf.CropLeft
    cropT = pf.CropTop
    cropR = pf.CropRight
    cropB = pf.CropBottom
    m_cropL = cropL
    m_cropT = cropT
    m_cropR = cropR
    m_cropB = cropB

    ' 临时展开裁剪，测出「未裁剪完整矩形」，随后立刻恢复，避免破坏原图
    IR_Trace "capture:expand-measure"
    pf.CropLeft = 0
    pf.CropTop = 0
    pf.CropRight = 0
    pf.CropBottom = 0
    fl = shp.Left
    ft = shp.Top
    fw = shp.Width
    fh = shp.Height
    pf.CropLeft = cropL
    pf.CropTop = cropT
    pf.CropRight = cropR
    pf.CropBottom = cropB

    ' 保护：测量异常时回退到数学换算（裁剪值以未裁剪显示尺寸为基准）
    If fw <= 0 Or fh <= 0 Then
        fl = m_left - cropL
        ft = m_top - cropT
        fw = m_w + cropL + cropR
        fh = m_h + cropT + cropB
    End If
    m_fullLeft = fl
    m_fullTop = ft
    m_fullW = fw
    m_fullH = fh
    IR_Trace "capture:expand-done " & IR_Fmt(fl) & "," & IR_Fmt(ft) & " " & _
             IR_Fmt(fw) & "x" & IR_Fmt(fh)

    IR_Trace "capture:transform"
    m_rotation = shp.Rotation
    ' 注意：PowerPoint 的成员名是 HorizontalFlip / VerticalFlip，
    '      没有 FlipHorizontal / FlipVertical（写成后者会编译报
    '      「方法和数据成员未找到」）。翻转本身也可用 Flip(msoFlipHorizontal) 方法。
    m_flipH = shp.HorizontalFlip
    m_flipV = shp.VerticalFlip
    m_zOrder = shp.ZOrderPosition
    m_shpName = shp.Name
    m_shpAlt = shp.AlternativeText
    m_shpTitle = shp.Title

    IR_Trace "capture:picprops"
    m_bright = pf.Brightness
    m_contrast = pf.Contrast
    m_transColor = pf.TransparencyColor
    m_transBg = pf.TransparentBackground
    m_colorType = pf.ColorType
    m_fillTrans = shp.Fill.Transparency

    IR_Trace "capture:meta"
    m_linkAddr = shp.ActionSettings(ppMouseClick).Hyperlink.Address
    m_linkSub = shp.ActionSettings(ppMouseClick).Hyperlink.SubAddress
    m_hasLink = (Len(m_linkAddr) > 0 Or Len(m_linkSub) > 0)

    IR_Trace "capture:anim"
    m_hasAnim = (shp.AnimationSettings.Animate = msoTrue)

    Err.Clear
    On Error GoTo 0
    m_ok = True
    IR_Capture = True
    IR_Trace "capture:done"
End Function

' ============================================================
'  ② 迁移 + ③ 还原
' ============================================================

' 核心替换。快照必须已经通过 IR_Capture 建立。
Private Function IR_Build(ByVal sld As Slide, ByVal srcShp As Shape, _
                          ByVal newImgPath As String, _
                          ByRef newShp As Shape, ByRef errMsg As String) As Boolean
    Dim styled As Boolean
    Dim animed As Boolean

    IR_Build = False
    Set newShp = Nothing
    IR_Trace "build:enter"
    On Error GoTo EH

    ' ②-1 抓取整块格式与动画（必须在原形状删除前）
    styled = False
    animed = False
    On Error Resume Next
    IR_Trace "build:pickup"
    srcShp.PickUp
    styled = (Err.Number = 0)
    Err.Clear
    If m_hasAnim Then
        IR_Trace "build:pickup-anim"
        srcShp.PickupAnimation
        animed = (Err.Number = 0)
        Err.Clear
    End If
    On Error GoTo EH
    IR_Trace "build:picked styled=" & styled & " animed=" & animed

    ' ②-2 先插入新图。先插后删的好处：插入失败时原图仍在，不会造成数据丢失。
    IR_Trace "build:addpicture"
    Set newShp = sld.Shapes.AddPicture( _
        FileName:=newImgPath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=m_fullLeft, Top:=m_fullTop, Width:=m_fullW, Height:=m_fullH)
    If newShp Is Nothing Then
        errMsg = "插入新图片失败（文件可能不是有效图片）：" & newImgPath
        Exit Function
    End If
    IR_Trace "build:added"

    ' ②-3 删除原图
    IR_Trace "build:delete-src"
    srcShp.Delete
    IR_Trace "build:src-deleted"

    ' ②-4 套用原格式与动画（Apply 会连带覆盖位置 / 尺寸 / 旋转，③ 中逐项写回）
    On Error Resume Next
    IR_Trace "build:apply"
    If styled Then newShp.Apply
    IR_Trace "build:apply-anim"
    If animed Then newShp.ApplyAnimation
    Err.Clear
    On Error GoTo EH
    IR_Trace "build:applied"

    ' ③ 逐项还原
    IR_Restore newShp

    IR_Build = True
    IR_Trace "build:done"
    Exit Function

EH:
    errMsg = "替换过程中出错：" & Err.Number & " " & Err.Description
    IR_Build = False
End Function

Private Sub IR_Restore(ByRef shp As Shape)
    Dim z     As Long
    Dim guard As Long

    On Error Resume Next

    ' --- 几何：先回到「未裁剪完整矩形」---
    IR_Trace "restore:geometry"
    shp.Left = m_fullLeft
    shp.Top = m_fullTop
    shp.Width = m_fullW
    shp.Height = m_fullH

    ' --- 裁剪：按「比例」复现，而不是照抄磅值（详见 IR_FixCrop）---
    ' 注意：IR_FixCrop 内部靠「套上种子裁剪 → 实测可见尺寸 → 按比例校正」工作，
    '      因此它必须在这一步之前拿到「未裁剪显示尺寸」——
    '      也就是上面刚设好的 shp.Width = m_fullW / shp.Height = m_fullH。
    '      千万别把上面那两行挪到 IR_FixCrop 之后。
    IR_Trace "restore:crop"
    IR_FixCrop shp

    ' --- 可见尺寸终校 ---
    ' 正常走到这里，裁剪比例已被 IR_FixCrop 校正到与源形状一致
    ' （即 1 - m_w / m_fullW），可见尺寸本应正好等于 m_w / m_h。
    ' 再显式赋一次是兜底：即使某次测量失败、比例没校正成功，
    ' 至少能保证「摆放尺寸」与原图完全一致。
    IR_Trace "restore:visible-size"
    If m_w > 0 Then shp.Width = m_w
    If m_h > 0 Then shp.Height = m_h

    ' --- 变换 ---
    ' 顺序有讲究：Shape.Flip 在翻转的同时会**把 Rotation 取反**
    ' （PowerPoint 为保持翻转后的外观不变而做的连带调整）。
    ' 因此必须「先翻转、后设旋转」；反过来写的话旋转值会被 Flip 改掉。
    IR_Trace "restore:transform"
    IR_SetFlip shp, m_flipH, m_flipV
    shp.Rotation = m_rotation

    ' --- 图片属性 ---
    ' 顺序同样有讲究：设置 ColorType 会把亮度 / 对比度重置回默认值，
    ' 所以 ColorType 必须第一个设，亮度 / 对比度排在它后面。
    IR_Trace "restore:picprops"
    shp.PictureFormat.ColorType = m_colorType
    shp.PictureFormat.Brightness = m_bright
    shp.PictureFormat.Contrast = m_contrast
    If m_transBg = msoTrue Then
        shp.PictureFormat.TransparencyColor = m_transColor
        shp.PictureFormat.TransparentBackground = msoTrue
    End If
    shp.Fill.Transparency = m_fillTrans

    ' --- 元数据 ---
    IR_Trace "restore:meta"
    If Len(m_shpName) > 0 Then shp.Name = m_shpName
    shp.AlternativeText = m_shpAlt
    If Len(m_shpTitle) > 0 Then shp.Title = m_shpTitle
    If m_hasLink Then
        shp.ActionSettings(ppMouseClick).Hyperlink.Address = m_linkAddr
        shp.ActionSettings(ppMouseClick).Hyperlink.SubAddress = m_linkSub
    End If

    ' --- 图层顺序：新形状必然在顶层，逐步下移回原位 ---
    IR_Trace "restore:zorder from " & shp.ZOrderPosition & " to " & m_zOrder
    z = m_zOrder
    guard = 0
    Do While shp.ZOrderPosition > z And guard < 1000
        shp.ZOrder msoSendBackward
        guard = guard + 1
    Loop
    guard = 0
    Do While shp.ZOrderPosition < z And guard < 1000
        shp.ZOrder msoBringForward
        guard = guard + 1
    Loop

    ' --- 位置终校 ---
    ' 旋转 / 翻转会改变外接框，前面按「未裁剪矩形」定位后可能偏移；
    ' 最后再对齐一次快照里读到的 Left/Top，保证与原图完全重合。
    IR_Trace "restore:final-pos"
    shp.Left = m_left
    shp.Top = m_top

    Err.Clear
    On Error GoTo 0
    IR_Trace "restore:done"
End Sub

' 按「比例」而不是「磅值」复现裁剪。
'
' 为什么不能照抄磅值：
'     PictureFormat.CropLeft 等虽然以「磅」为单位暴露，但存进文件的其实是
'     **比例** —— <a:srcRect l="8000"> 就是 8%，而 8% 的分母是「图片原始尺寸」。
'     于是同一磅值对不同像素尺寸的图折算出的比例完全不同：
'     实测 200px 的图（原始宽 150 磅）换成 300px 的图（原始宽 225 磅）后，
'     同样是 12 磅，8% 掉成了 5.33%。后果是：
'       · 文件里记录的裁剪比例变了 → 看到的图像区域变了
'       · 裁剪后的可见矩形跟着漂移（实测 275.2 → 290.1）
'     而用户想要的显然是「同样的可见区域比例 + 同样的摆放尺寸」。
'
' 做法（全程实测，不假设磅值与原始尺寸之间的换算关系）：
'     调用前 IR_Restore 已把裁剪清零、把显示尺寸设为快照里的未裁剪尺寸
'     m_fullW / m_fullH。于是：
'       ① 直接把源形状的裁剪磅值当「种子」套上去，量出它在新图上折算出的
'          可见宽 vw；
'       ② 「可见宽的减少量」与「裁剪磅值之和」严格成正比，所以只要按比例
'          把种子缩放即可命中目标可见宽：
'              缩放系数 = (1 - m_w / m_fullW) / (1 - vw / m_fullW)
'          校正后文件里记录的比例正好等于源形状的比例 (1 - m_w / m_fullW)。
'
'     这样既不依赖对 PowerPoint 内部换算的猜测，也不依赖新图的像素 / DPI 元数据。
'
' ★ 踩坑记录 ★
'     最初写成「先读 shp.Width，再反推新图原始尺寸」，结果读数是在**还没套
'     任何裁剪**的时候取的，vw 恒等于 m_fullW，判断 `vw < m_fullW - 0.01`
'     永远不成立，整段被跳过 —— 表现成「新图的裁剪彻底丢了」（OOXML 里
'     srcRect 直接消失）。教训：**想在裁剪后量尺寸，就必须先把裁剪套上去。**
Private Sub IR_FixCrop(ByRef shp As Shape)
    Dim pf   As Object
    Dim sumW As Single, sumH As Single
    Dim fw   As Single, fh   As Single
    Dim vw   As Single, vh   As Single
    Dim sc   As Single

    On Error Resume Next
    Set pf = shp.PictureFormat
    If pf Is Nothing Then Exit Sub

    ' ---- 水平方向 ----
    sumW = m_cropL + m_cropR
    fw = m_fullW
    If sumW > 0.0001 And fw > 0.0001 And m_w > 0.0001 And m_w < fw - 0.01 Then
        ' ① 套上源形状的裁剪磅值当种子，量出它在新图上折算出的可见宽
        pf.CropLeft = m_cropL
        pf.CropRight = m_cropR
        vw = shp.Width
        IR_Trace "fixcrop:h seed=" & IR_Fmt(sumW) & " vw=" & IR_Fmt(vw) & _
                 " fw=" & IR_Fmt(fw) & " target=" & IR_Fmt(m_w)
        ' ② 按比例把种子缩放到目标可见宽
        If vw > 0.0001 And vw < fw - 0.01 Then
            sc = (1 - m_w / fw) / (1 - vw / fw)
            IR_Trace "fixcrop:h scale=" & IR_Fmt(sc)
            If sc > 0.0001 Then
                pf.CropLeft = m_cropL * sc
                pf.CropRight = m_cropR * sc
            End If
        End If
    End If

    ' ---- 垂直方向 ----
    sumH = m_cropT + m_cropB
    fh = m_fullH
    If sumH > 0.0001 And fh > 0.0001 And m_h > 0.0001 And m_h < fh - 0.01 Then
        pf.CropTop = m_cropT
        pf.CropBottom = m_cropB
        vh = shp.Height
        IR_Trace "fixcrop:v seed=" & IR_Fmt(sumH) & " vh=" & IR_Fmt(vh) & _
                 " fh=" & IR_Fmt(fh) & " target=" & IR_Fmt(m_h)
        If vh > 0.0001 And vh < fh - 0.01 Then
            sc = (1 - m_h / fh) / (1 - vh / fh)
            IR_Trace "fixcrop:v scale=" & IR_Fmt(sc)
            If sc > 0.0001 Then
                pf.CropTop = m_cropT * sc
                pf.CropBottom = m_cropB * sc
            End If
        End If
    End If

    Err.Clear
    On Error GoTo 0
End Sub

' 文件是否存在。
'
' 为什么绕开 VBA 内置的 Dir：
'     Dir 是宿主内置函数，在本机 PowerPoint 的 .ppam 隐藏工程里保存后会被
'     VBE 规范化成小写 dir；重新加载后如果它被当成未声明标识符，就会抛
'     「隐藏的模块中的编译错误」，既不指行也不指名字，极难定位。
'     Scripting.FileSystemObject 是后期绑定调用，完全没有这个问题。
Private Function IR_FileExists(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    If Len(p) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    IR_FileExists = fso.FileExists(p)
    Err.Clear
    On Error GoTo 0
End Function

' 把形状的翻转状态设为指定值。
'
' 为什么不能直接赋值：
'     PowerPoint 的 Shape.HorizontalFlip / VerticalFlip 是**只读**属性，
'     写成 shp.HorizontalFlip = xxx 会编译报「不能给只读属性赋值」。
'     改变翻转只能用 Flip 方法，而 Flip 是**在现有状态上取反**（不是设值），
'     所以必须「先读当前值，与目标不一致才翻转一次」。
Private Sub IR_SetFlip(ByRef shp As Shape, ByVal wantH As Long, ByVal wantV As Long)
    On Error Resume Next
    If (shp.HorizontalFlip = msoTrue) <> (wantH = msoTrue) Then
        shp.Flip msoFlipHorizontal
    End If
    If (shp.VerticalFlip = msoTrue) <> (wantV = msoTrue) Then
        shp.Flip msoFlipVertical
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' ============================================================
'  对外核心（与界面无关，便于自动化验证）
' ============================================================

' ============================================================
'  ③ 组合（编组）支持
'
'  为什么必须单独走一套流程：
'      PowerPoint 的 GroupShapes 集合**没有 AddPicture 成员**
'      （用 IDispatch 名称表实测查无此名），所以无法在组合内部直接插入
'      一张新图；而 Fill.UserPicture 虽然不报错，却会把形状变成「图片
'      填充」并让裁剪失效（CropLeft 写进去再读回来是 -2147483648），
'      语义已经被破坏。
'
'      唯一可行的路径是：
'          快照组合 → 变换归零 → 解组 → 【复用普通替换逻辑】 → 重组 → 复原组合属性
'
'      好处：解组之后目标图片就是幻灯片上的一张普通图片，
'            几何 / 裁剪 / 亮度对比度 / 透明度 / 图层 / 元数据 / 动画
'            这些还原逻辑一行都不用改，全部沿用已经验证过的实现。
'
'  ★ 变换归零为什么必须 ★
'      解组时 PowerPoint 会把组合的旋转 / 翻转「烘焙」进每个子形状
'      （实测：组合旋转 20°，解组后子形状的 Rotation 也变成 20°），
'      重组之后组合自身的旋转就回不来了。
'      先把组合的变换临时归零，解组就不会烘焙；重组后再把变换设回去，
'      外观与数据表示双双无损 —— 无变换 / 旋转 / 翻转 / 旋转+翻转
'      四个场景实测全部完全一致。
'
'  ★ 顺序（每一步都有原因，不要随手改）★
'      重命名   从最外层整棵树一次性做完（避免解组后同名撞车）
'      归零     最外层 → 最内层（外层归零会改变内层的绝对坐标）
'      解组     最外层 → 最内层（成员逐层落到幻灯片上）
'      重组     最内层 → 最外层（先组好内层，外层才能把新组当成员）
'      每层复原 先翻转、后旋转（Flip 会把 Rotation 取反）
'
'  【注意】已知边界：嵌套组合（组里还有组）—— 实现了，但未经验证
'      上面的「组合链」循环（IR_BuildChain / m_chain）本来就是为嵌套组合写的：
'      Ungroup 只解一层，所以按链逐层下钻、再逐层重组。
'
'      但实测本环境的 PowerPoint **无法构造出嵌套组合**：
'          · Shapes.Range(...).Group()            → 展平
'          · 选中后 ExecuteMso("ObjectsGroup")     → 展平
'          · 存盘 → 重新打开 → 再组合               → 仍然展平
'      两条路径、跨存盘重开都试过，结果一致：组合操作会把内层组合摊平。
'      既然造不出用例，这条分支就无法被回归覆盖。
'
'      推论：若真的遇到嵌套组合（例如来自第三方工具生成的 .pptx，
'      OOXML 格式本身是支持 p:grpSp 套 p:grpSp 的），
'      重组阶段很可能把内层组合摊平 —— 图片本身的属性仍然会被完整保留，
'      但内层「组合」这个容器可能不再存在。目前按「尽力而为」处理，
'      未做拦截；真有需要再按上面的实测方法补用例。
' ============================================================

' 是否组合（编组）对象。
Private Function IR_IsGroup(ByVal shp As Shape) As Boolean
    On Error Resume Next
    If shp Is Nothing Then Exit Function
    IR_IsGroup = (shp.Type = msoGroup)
    Err.Clear
End Function

' 形状的直接父组合；不在组合里则返回 Nothing。
'
' ★ 为什么不能用 Shape.Parent（踩过的坑）★
'     实测：组合内形状的 Parent 返回的是**幻灯片**（Slide1），而不是组合，
'     对 Parent 取 .Type 会直接抛错。也就是说「用 Parent 是不是 Group
'     来判断形状属于哪个组合」永远得到「否」—— 后果是组内图片被当成
'     幻灯片上的普通图片处理，替换时会把组合整个拆散。
'     正确做法是使用专门为此设计的 Shape.ParentGroup：
'         · 形状在组合里 → 返回所属组合（Type = msoGroup）
'         · 形状在顶层   → 调用直接抛错（正好用来判定「不在组合里」）
Private Function IR_ParentGroup(ByVal shp As Shape) As Shape
    Dim g As Shape
    On Error Resume Next
    If shp Is Nothing Then Exit Function
    Set g = shp.ParentGroup
    If g Is Nothing Then Exit Function
    If g.Type = msoGroup Then Set IR_ParentGroup = g
    Err.Clear
End Function

' 递归收集 shp 子树里的全部图片（含自身）。顺序 = 深度优先遍历顺序。
Private Sub IR_CollectPictures(ByVal shp As Shape, ByRef pics As Collection)
    Dim i As Long
    On Error Resume Next
    If shp Is Nothing Then Exit Sub
    If shp.Type = msoPicture Then
        pics.Add shp
        Err.Clear
        Exit Sub
    End If
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            IR_CollectPictures shp.GroupItems.Item(i), pics
        Next i
    End If
    Err.Clear
End Sub

' 递归给 shp 及全部后代分配唯一临时名，并记录「临时名 → 原名」。
'
' 为什么一定要重命名：
'     解组之后所有形状会落到同一层，而 PowerPoint 允许不同组合里存在
'     同名形状。重名会让重组时的 ShapeRange.Range(名字) 取错对象。
'     先把整棵树改成全局唯一名，重组完成后再按表还原原名。
Private Sub IR_RenameTree(ByVal shp As Shape)
    Dim i As Long
    Dim tmp As String
    On Error Resume Next
    If shp Is Nothing Then Exit Sub
    m_renameIdx = m_renameIdx + 1
    tmp = "PIR_TMP_" & m_renameIdx
    m_names.Add tmp, shp.Name
    shp.Name = tmp
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            IR_RenameTree shp.GroupItems.Item(i)
        Next i
    End If
    Err.Clear
End Sub

' 快照一层组合的全部可还原属性与成员顺序。
' 成员名用 Chr(1) 连接，避免形状名里出现分隔符时被切错。
Private Function IR_SnapLayer(ByVal g As Shape) As Object
    Dim d As Object
    Dim s As String
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    d("name") = g.Name
    d("l") = g.Left
    d("t") = g.Top
    d("w") = g.Width
    d("h") = g.Height
    d("rot") = g.Rotation
    d("fh") = g.HorizontalFlip
    d("fv") = g.VerticalFlip
    d("z") = g.ZOrderPosition

    s = ""
    For i = 1 To g.GroupItems.Count
        If i > 1 Then s = s & Chr$(1)
        s = s & g.GroupItems.Item(i).Name
    Next i
    d("members") = s

    ' 组合自带动画时记一笔（AnimationSettings 是旧 API，取不到就当没有）
    d("anim") = False
    Err.Clear
    If g.AnimationSettings.Animate = msoTrue Then d("anim") = True
    Err.Clear

    Set IR_SnapLayer = d
End Function

' 把组合的旋转 / 翻转临时归零。
'
' Flip 是**取反**语义（不是设值），所以要先读当前状态再决定是否翻转；
' 而且 Flip 会连带把 Rotation 取反，因此翻转必须排在旋转之前。
Private Sub IR_ZeroTransform(ByRef g As Shape)
    On Error Resume Next
    If g Is Nothing Then Exit Sub
    If g.HorizontalFlip = msoTrue Then g.Flip msoFlipHorizontal
    If g.VerticalFlip = msoTrue Then g.Flip msoFlipVertical
    g.Rotation = 0
    Err.Clear
End Sub

' 把组合的变换恢复成快照值。同样「先翻转、后旋转」。
Private Sub IR_ApplyTransform(ByRef g As Shape, ByVal d As Object)
    On Error Resume Next
    If g Is Nothing Then Exit Sub
    If (g.HorizontalFlip = msoTrue) <> (d("fh") = msoTrue) Then g.Flip msoFlipHorizontal
    If (g.VerticalFlip = msoTrue) <> (d("fv") = msoTrue) Then g.Flip msoFlipVertical
    g.Rotation = d("rot")
    Err.Clear
End Sub

' 把 shp 挪到指定的 Z 序位置。
'
' ZOrderPosition 是**只读**属性，改 Z 序只能用 ZOrder 方法一格一格挪；
' guard 防止宿主在极端情况下让位置反复抖动时死循环。
Private Sub IR_ShiftZ(ByRef shp As Shape, ByVal target As Long)
    Dim guard As Long
    On Error Resume Next
    If shp Is Nothing Then Exit Sub
    If target < 1 Then Exit Sub
    guard = 0
    Do While shp.ZOrderPosition > target And guard < 2000
        shp.ZOrder msoSendBackward
        guard = guard + 1
    Loop
    guard = 0
    Do While shp.ZOrderPosition < target And guard < 2000
        shp.ZOrder msoBringForward
        guard = guard + 1
    Loop
    Err.Clear
End Sub

' 向上收集祖先组合链。返回集合里：索引 1 = 最外层，末项 = 图片的直接父组合。
Private Function IR_BuildChain(ByVal imgShp As Shape, ByRef chain As Collection) As Boolean
    Dim tmp As New Collection
    Dim g As Shape
    Dim i As Long

    Set chain = New Collection
    On Error Resume Next
    Set g = IR_ParentGroup(imgShp)
    Do While Not g Is Nothing
        tmp.Add g
        Set g = IR_ParentGroup(g)
    Loop
    ' tmp 里是「内 → 外」，反过来存成「外 → 内」
    For i = tmp.Count To 1 Step -1
        chain.Add tmp(i)
    Next i
    IR_BuildChain = (chain.Count > 0)
    Err.Clear
End Function

' 递归恢复子树内所有形状的原名（含新图：它继承了被替换图片的临时名）。
Private Sub IR_RestoreNames(ByVal shp As Shape)
    Dim i As Long
    Dim nm As String
    On Error Resume Next
    If shp Is Nothing Then Exit Sub
    nm = shp.Name
    If m_names.Exists(nm) Then shp.Name = m_names(nm)
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            IR_RestoreNames shp.GroupItems.Item(i)
        Next i
    End If
    Err.Clear
End Sub

' 从当前选区解析出「要替换的那张图片」。
'
' 支持四种情形：
'   ① 直接选中一张图片
'   ② 双击进入组合后选中了组内的某张图片 —— 这是最精确的表达方式。
'      PowerPoint 此时仍把 Selection.ShapeRange 指向**组合**（Type=6），
'      但 HasChildShapeRange = True，真正的子形状要从 ChildShapeRange 取。
'      实测：选中整个组合时 HasChildShapeRange=False、ChildShapeRange 直接
'      抛 Invalid request，所以这两个属性正是区分两种状态的依据。
'   ③ 选中整个组合，且组合内只有一张图片 —— 意图明确，直接采用
'   ④ 选中整个组合，组合内有多张图片 —— 无法判断意图，给出可操作提示
Private Function IR_ResolveTarget(ByVal sel As Selection, _
                                  ByRef shp As Shape, _
                                  ByRef errMsg As String) As Boolean
    Dim cand As Shape
    Dim pics As New Collection
    Dim n As Long
    Dim ct As Long
    Dim cn As String

    ' ★ 为什么 hasChild 用 Long 而不是 Boolean（踩过的坑，别再改回去）★
    '     PowerPoint 经 COM 把 HasChildShapeRange 这类「三态」属性以
    '     VARIANT_TRUE 返回，本环境实测其原始值是 **1**（不是 VBA 习惯的 -1）。
    '     而 VBA 把这个原始值**不做归一化**地直接塞进 Boolean 变量，
    '     于是变量里躺着的就是 1。此后任何 `Not 变量` 都会退化成按位取反：
    '         Not 1 = -2  →  非 0 即真  →  `If Not hasChild Then` 永远成立！
    '     后果：刚刚从 ChildShapeRange 里解析出来的"组内那张图片"，
    '     立刻又被 sel.ShapeRange(1)（＝整个组合）覆盖掉，
    '     最终弹出「选中的对象不是图片（当前类型 6）」。
    '     用 Long 承接 + 显式 `<> 0` 判断可以彻底绕开这个陷阱。
    Dim hasChild As Long

    IR_ResolveTarget = False
    Set shp = Nothing
    errMsg = ""
    On Error Resume Next

    ' ---- 情形 ②：双击进入组合后，单击选中了组合里的具体子形状 ----
    hasChild = 0
    hasChild = sel.HasChildShapeRange      ' -1 / 1 = 有；0 = 无
    Err.Clear

    If hasChild <> 0 Then
        n = sel.ChildShapeRange.Count
        If n = 1 Then
            Set cand = sel.ChildShapeRange(1)
        ElseIf n > 1 Then
            errMsg = "一次只能替换一张图片。当前选中了 " & n & " 个对象。"
            Err.Clear
            Exit Function
        Else
            hasChild = 0
        End If
    End If

    ' ---- 情形 ①③：没进组合（或没选中子形状）→ 看整体选区 ----
    If hasChild = 0 Then
        n = sel.ShapeRange.Count
        If n <> 1 Then
            errMsg = "一次只能替换一张图片。当前选中了 " & n & " 个对象。"
            Err.Clear
            Exit Function
        End If
        Set cand = sel.ShapeRange(1)
    End If

    If cand Is Nothing Then
        errMsg = "没有解析到可替换的对象。"
        Err.Clear
        Exit Function
    End If

    ' 首次读取立刻缓存：COM 对象的属性不要在判断里反复读。
    cn = cand.Name
    ct = cand.Type
    IR_Trace "resolve:sel=" & cn & " type=" & ct & " child=" & hasChild
    Err.Clear

    ' ---- 直接就是图片 ----
    If ct = msoPicture Then
        Set shp = cand
        IR_ResolveTarget = True
        Exit Function
    End If

    ' ---- 是组合：往里找图片 ----
    If ct = msoGroup Then
        IR_CollectPictures cand, pics
        If pics.Count = 1 Then
            Set shp = pics(1)
            IR_ResolveTarget = True
            Exit Function
        ElseIf pics.Count = 0 Then
            errMsg = "所选组合里没有图片，无法替换。" & vbCrLf & _
                     "请改选一张图片，或内含图片的组合。"
            Exit Function
        Else
            errMsg = "所选组合里有 " & pics.Count & " 张图片，无法确定要替换哪一张。" & _
                     vbCrLf & vbCrLf & _
                     "请【双击进入该组合】，再单击要替换的那张图片，" & vbCrLf & _
                     "然后重新点「原位替换图片」。"
            Exit Function
        End If
    End If

    errMsg = "选中的对象不是图片（当前类型 " & ct & "）。" & vbCrLf & _
             "请选中一张图片，或内含图片的组合。"
    Err.Clear
End Function

' 组合内的原位替换。imgShp 必须是某个组合的成员。
' 成功返回 True；失败返回 False 并填充 errMsg。
Private Function IR_ReplaceInGroup(ByVal imgShp As Shape, _
                                   ByVal newImgPath As String, _
                                   ByRef errMsg As String, _
                                   ByRef outShp As Shape) As Boolean
    Dim chain    As New Collection
    Dim sld      As Slide
    Dim o        As Object
    Dim i        As Long
    Dim d        As Object
    Dim imgTmp   As String
    Dim tgt      As Shape
    Dim newShp   As Shape
    Dim names()  As String
    Dim topGrp   As Shape
    Dim animed   As Boolean
    Dim ungroupped As Boolean

    IR_ReplaceInGroup = False
    Set outShp = Nothing
    Set newShp = Nothing
    IR_Trace "grp:enter"
    On Error GoTo EH

    ' ---- 找所在幻灯片（沿 Parent 向上直到 Slide）----
    Set o = imgShp.Parent
    i = 0
    Do While Not (TypeOf o Is Slide) And i < 50
        Set o = o.Parent
        i = i + 1
    Loop
    If Not (TypeOf o Is Slide) Then
        errMsg = "无法定位图片所在的幻灯片。"
        Exit Function
    End If
    Set sld = o

    ' ---- 前置校验：尽量把失败挡在「解组之前」，避免组合被拆开 ----
    '
    ' 这里刻意只做**只读**检查（类型 / 文件存在性），不调用 IR_Capture。
    ' IR_Capture 会临时把裁剪清零再量完整矩形，在组合内部做这种写操作会
    ' 顺带扰动组合的包围盒。真正的属性快照放在解组之后做 —— 那时目标图片
    ' 已经是幻灯片上的普通图片，与既有逻辑完全一致。
    If imgShp.Type <> msoPicture Then
        errMsg = "所选对象不是图片（当前类型 " & imgShp.Type & "）。"
        Exit Function
    End If
    If Not IR_FileExists(newImgPath) Then
        errMsg = "新图片文件不存在：" & newImgPath
        Exit Function
    End If

    ' ---- 祖先链 ----
    If Not IR_BuildChain(imgShp, chain) Then
        errMsg = "无法定位图片所在的组合。"
        Exit Function
    End If
    IR_Trace "grp:chain=" & chain.Count

    ' ---- 唯一重命名 + 逐层快照 ----
    Set m_names = CreateObject("Scripting.Dictionary")
    m_renameIdx = 0
    IR_RenameTree chain(1)
    imgTmp = imgShp.Name
    Set m_chain = New Collection
    For i = 1 To chain.Count
        m_chain.Add IR_SnapLayer(chain(i))
    Next i

    ' ---- 动画先抓（只有最外层组合有意义）----
    animed = False
    On Error Resume Next
    If m_chain(1)("anim") Then
        chain(1).PickupAnimation
        animed = (Err.Number = 0)
        Err.Clear
    End If
    On Error GoTo EH

    ' ---- 变换归零（外 → 内），保证解组不烘焙 ----
    For i = 1 To chain.Count
        IR_ZeroTransform chain(i)
    Next i
    IR_Trace "grp:zeroed"

    ' ---- 解组（外 → 内）----
    For i = 1 To chain.Count
        chain(i).Ungroup
    Next i
    ungroupped = True
    IR_Trace "grp:ungrouped"

    ' ---- 目标图片现在落到幻灯片上 ----
    Set tgt = Nothing
    On Error Resume Next
    Set tgt = sld.Shapes.Item(imgTmp)
    Err.Clear
    On Error GoTo EH
    If tgt Is Nothing Then
        errMsg = "取消组合后未找到目标图片，操作已中止。"
        Exit Function
    End If

    ' ---- 复用普通替换逻辑 ----
    If Not ImageReplace_ReplaceShape(tgt, newImgPath, errMsg, newShp) Then
        Exit Function
    End If
    IR_Trace "grp:replaced"

    ' ---- 新图沿用原图的临时名，重组时按名字即可对上 ----
    If Not newShp Is Nothing Then
        On Error Resume Next
        newShp.Name = imgTmp
        Err.Clear
        On Error GoTo EH
    End If

    ' ---- 重组（内 → 外）+ 逐层复原属性 ----
    Set topGrp = Nothing
    For i = m_chain.Count To 1 Step -1
        Set d = m_chain(i)
        names = Split(CStr(d("members")), Chr$(1))
        Set topGrp = sld.Shapes.Range(names).Group()
        IR_Trace "grp:regroup-level" & i
        topGrp.Name = d("name")
        topGrp.Left = d("l")
        topGrp.Top = d("t")
        topGrp.Width = d("w")
        topGrp.Height = d("h")
        IR_ApplyTransform topGrp, d
        IR_ShiftZ topGrp, CLng(d("z"))
    Next i

    ' ---- 动画还原 ----
    If animed And Not topGrp Is Nothing Then
        On Error Resume Next
        topGrp.ApplyAnimation
        Err.Clear
        On Error GoTo EH
    End If

    ' ---- 恢复所有形状的原名（新图继承被替换图片的名字）----
    If Not topGrp Is Nothing Then IR_RestoreNames topGrp

    Set outShp = newShp
    IR_ReplaceInGroup = True
    IR_Trace "grp:done"
    Exit Function

EH:
    errMsg = "组合内替换失败：" & Err.Number & " " & Err.Description
    IR_ReplaceInGroup = False
    ' 已经解组的话，尽力把组合重新拼回去，避免把用户的结构拆散。
    If ungroupped Then IR_TryRebuild sld
End Function

' 尽力按快照把已经解开的组合重新拼回去（出错兜底用）。
' 这里刻意只做「重建结构」，不碰任何形状自身的属性 —— 属性早已在
' 解组前改过的话，回滚也救不回来，如实重建至少不让结构散落。
Private Sub IR_TryRebuild(ByVal sld As Slide)
    Dim i As Long
    Dim d As Object
    Dim names() As String
    On Error Resume Next
    If m_chain Is Nothing Then Exit Sub
    For i = m_chain.Count To 1 Step -1
        Set d = m_chain(i)
        names = Split(CStr(d("members")), Chr$(1))
        Dim g As Shape
        Set g = sld.Shapes.Range(names).Group()
        g.Name = d("name")
        IR_ApplyTransform g, d
        IR_ShiftZ g, CLng(d("z"))
    Next i
    Err.Clear
End Sub

' 把 shp 用 newImgPath 这张图原位替换。
' 成功返回 True；失败返回 False 并填充 errMsg。
' 失败语义：errMsg 非空；若失败发生在删除原图之前，原图不受影响。
Public Function ImageReplace_ReplaceShape(ByVal shp As Shape, _
                                          ByVal newImgPath As String, _
                                          ByRef errMsg As String, _
                                          Optional ByRef outShp As Shape) As Boolean
    Dim sld     As Slide
    Dim newShp  As Shape
    Dim srcName As String
    Dim pg      As Shape

    On Error GoTo EH
    errMsg = ""
    ImageReplace_ReplaceShape = False
    Set outShp = Nothing

    If shp Is Nothing Then
        errMsg = "未指定要替换的图片。"
        Exit Function
    End If
    If shp.Type <> msoPicture Then
        errMsg = "所选对象不是图片（当前类型 " & shp.Type & "）。"
        Exit Function
    End If
    If Len(newImgPath) = 0 Then
        errMsg = "未指定新图片路径。"
        Exit Function
    End If
    ' 文件存在性检查。
    ' 刻意**不用 VBA 内置的 Dir 函数**：Dir 属于宿主内置函数，在 .ppam 这种
    ' 隐藏工程里保存后会被 VBE 规范化成小写 dir，重新加载时有被当成未声明
    ' 标识符的风险，而报出来的只有一句无法定位的「隐藏的模块中的编译错误」。
    ' 改用 Scripting.FileSystemObject —— 与「选页另存」插件保持同一套做法。
    If Not IR_FileExists(newImgPath) Then
        errMsg = "新图片文件不存在：" & newImgPath
        Exit Function
    End If
    ' 图片在组合（编组）里 → 交给组合流程：
    '   快照组合 → 变换归零 → 解组 → 回到本函数做普通替换 → 重组 → 复原组合属性。
    ' 解组之后目标图片就是幻灯片上的普通图片，会直接命中下面的普通分支。
    '
    ' 注意判据是 ParentGroup 而不是 Parent —— 组内形状的 Parent 返回的是
    ' 幻灯片，用 Parent 判断不出来（详见 IR_ParentGroup 的说明）。
    Set pg = IR_ParentGroup(shp)
    If Not pg Is Nothing Then
        ImageReplace_ReplaceShape = IR_ReplaceInGroup(shp, newImgPath, errMsg, outShp)
        Exit Function
    End If

    ' 走到这里说明不在组合里，只可能是幻灯片（或母版 / 备注页等不支持的位置）
    If Not (TypeOf shp.Parent Is Slide) Then
        errMsg = "所选图片不在幻灯片上，无法原位替换。"
        Exit Function
    End If

    Set sld = shp.Parent
    srcName = shp.Name

    ' ① 快照
    If Not IR_Capture(shp) Then
        errMsg = "无法读取源图片属性（可能不是标准图片对象）。"
        Exit Function
    End If

    ' ② + ③ 迁移与还原
    If Not IR_Build(sld, shp, newImgPath, newShp, errMsg) Then
        Exit Function
    End If

    ' 兜底：确保名称回到原值（正式还原在 IR_Restore 里）
    On Error Resume Next
    If Len(srcName) > 0 And newShp.Name <> srcName Then newShp.Name = srcName
    On Error GoTo EH

    Set outShp = newShp
    ImageReplace_ReplaceShape = True
    Exit Function

EH:
    errMsg = "替换失败：" & Err.Number & " " & Err.Description
    ImageReplace_ReplaceShape = False
End Function

' 把一张图片的关键属性汇总成一行文本，便于自动化对比与排查。
' 参数声明为 Optional：这样不带参数调用（Application.Run 无参）也能成立，
' 于是它同时充当本模块的「健康检查」入口 —— 外部脚本可以借此确认
' m_ImageReplace 真的被加载进 PowerPoint 并且能编译、能执行，
' 而不是只躺在包里。传 Nothing 时原样返回 "(Nothing)"，无任何副作用。
Public Function ImageReplace_Describe(Optional ByVal shp As Shape) As String
    Dim r As String

    On Error Resume Next
    If shp Is Nothing Then
        ImageReplace_Describe = "(Nothing)"
        Exit Function
    End If

    r = "NAME=" & shp.Name
    r = r & "|L=" & IR_Fmt(shp.Left) & "|T=" & IR_Fmt(shp.Top)
    r = r & "|W=" & IR_Fmt(shp.Width) & "|H=" & IR_Fmt(shp.Height)
    r = r & "|ROT=" & IR_Fmt(shp.Rotation)
    r = r & "|FLIPH=" & shp.HorizontalFlip & "|FLIPV=" & shp.VerticalFlip
    r = r & "|Z=" & shp.ZOrderPosition
    r = r & "|CROP=" & IR_Fmt(shp.PictureFormat.CropLeft) & "," & _
                        IR_Fmt(shp.PictureFormat.CropTop) & "," & _
                        IR_Fmt(shp.PictureFormat.CropRight) & "," & _
                        IR_Fmt(shp.PictureFormat.CropBottom)
    r = r & "|BRIGHT=" & IR_Fmt(shp.PictureFormat.Brightness)
    r = r & "|CONTRAST=" & IR_Fmt(shp.PictureFormat.Contrast)
    r = r & "|COLORTYPE=" & shp.PictureFormat.ColorType
    r = r & "|FILLTRANS=" & IR_Fmt(shp.Fill.Transparency)
    r = r & "|ALT=" & shp.AlternativeText
    r = r & "|TITLE=" & shp.Title
    r = r & "|ANIM=" & shp.AnimationSettings.Animate
    r = r & "|LINK=" & shp.ActionSettings(ppMouseClick).Hyperlink.Address

    Err.Clear
    On Error GoTo 0
    ImageReplace_Describe = r
End Function

' ============================================================
'  工具
' ============================================================

' 数值格式化：整数值不显示多余小数位（288 → "288"，而非 "288."）
Private Function IR_Fmt(ByVal v As Double) As String
    Dim i As Long
    On Error Resume Next
    If Abs(v) < 0.0005 Then
        IR_Fmt = "0"
        Exit Function
    End If
    i = Int(Abs(v) + 0.5)
    If Abs(Abs(v) - i) < 0.0005 Then
        IR_Fmt = IIf(v < 0, "-", "") & CStr(i)
    Else
        IR_Fmt = Format$(v, "0.000")
    End If
End Function
