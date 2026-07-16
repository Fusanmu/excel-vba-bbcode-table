Option Explicit

'========================
' 可调参数（按需改）
'========================
Private Const TABLE_WIDTH As String = "70%"     ' 例如 "100%" / "70%"
Private Const TABLE_BG As String = "white"      ' table 默认背景（论坛支持的颜色名或 #RRGGBB）

Private Const USE_ROW_MODE_COLOR As Boolean = True      ' 是否给 tr 上“行主色”
Private Const TR_COLOR_USE_HEX As Boolean = True        ' tr 颜色用 #RRGGBB（否则尝试映射成 lightblue 等少量名字）

Private Const USE_CELL_BACKCOLOR As Boolean = True      ' 是否对“异色单元格”套 [backcolor]
Private Const BACKCOLOR_PAD_FULLWIDTH As Boolean = True ' backcolor 内侧是否加全角空格，增强“色块感”
Private Const ESCAPE_TEXT_BRACKETS As Boolean = True    ' 是否把文本里的 [ ] 转义，避免破坏 BBCode

'========================
' 入口宏
'========================
Public Sub CopySelectionToBBCodeTable()
    Dim rng As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "请先在表格里选中一个区域。", vbExclamation
        Exit Sub
    End If

    Set rng = Selection

    If rng.Cells.CountLarge > 200000 Then
        If MsgBox("选区很大，导出可能较慢。继续吗？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If

    Dim bb As String
    bb = BuildBBCodeTable(rng)

    CopyTextToClipboard bb
    MsgBox "已复制 BBCode 到剪贴板。去论坛直接粘贴即可。", vbInformation
End Sub

'========================
' 生成整张表
'========================
Private Function BuildBBCodeTable(ByVal rng As Range) As String
    Dim rCount As Long, cCount As Long
    rCount = rng.Rows.Count
    cCount = rng.Columns.Count

    Dim sb As String
    sb = sb & "[table=" & TABLE_WIDTH & "," & TABLE_BG & "]" & vbCrLf

    ' 用来跳过合并单元格中非左上角的格子
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim r As Long, c As Long
    For r = 1 To rCount
        Dim rowModeColor As Variant
        If USE_ROW_MODE_COLOR Then
            rowModeColor = GetRowModeFillColor(rng.Rows(r))
        Else
            rowModeColor = Empty
        End If

        Dim trTag As String
        trTag = BuildTrTag(rowModeColor)
        sb = sb & trTag & vbCrLf

        For c = 1 To cCount
            Dim cell As Range
            Set cell = rng.Cells(r, c)

            Dim td As String
            td = BuildTd(cell, rng, rowModeColor, seen)
            If Len(td) > 0 Then sb = sb & td & vbCrLf
        Next c

        sb = sb & "[/tr]" & vbCrLf
    Next r

    sb = sb & "[/table]"
    BuildBBCodeTable = sb
End Function

'========================
' tr 标签（行底色）
'========================
Private Function BuildTrTag(ByVal rowModeColor As Variant) As String
    If IsEmpty(rowModeColor) Then
        BuildTrTag = "[tr]"
        Exit Function
    End If

    If TR_COLOR_USE_HEX Then
        BuildTrTag = "[tr=" & ColorLongToHex(CLng(rowModeColor)) & "]"
    Else
        BuildTrTag = "[tr=" & ClosestNamedColor(CLng(rowModeColor)) & "]"
    End If
End Function

'========================
' td 构建（含合并单元格、对齐、字体、backcolor）
'========================
Private Function BuildTd(ByVal cell As Range, ByVal whole As Range, ByVal rowModeColor As Variant, ByVal seen As Object) As String
    ' 合并单元格处理：只输出 MergeArea 的左上角
    Dim rs As Long, cs As Long
    rs = 1
    cs = 1

    If cell.MergeCells Then
        Dim ma As Range
        Set ma = cell.MergeArea

        Dim key As String
        key = ma.Address(External:=True)

        If Not seen.Exists(key) Then
            seen.Add key, True
            If cell.Address = ma.Cells(1, 1).Address Then
                rs = ma.Rows.Count
                cs = ma.Columns.Count
            Else
                BuildTd = "" ' 非左上角直接跳过
                Exit Function
            End If
        Else
            BuildTd = ""
            Exit Function
        End If
    End If

    Dim widthPct As Long
    widthPct = GetColumnWidthPercent(cell, whole, cs)

    Dim content As String
    content = GetCellDisplayText(cell)

    If ESCAPE_TEXT_BRACKETS Then
        content = Replace(content, "[", "&#91;")
        content = Replace(content, "]", "&#93;")
    End If

    ' 换行 -> [br]
    content = NormalizeLineBreaksToBr(content)

    ' 空内容给个不可见占位，避免 backcolor 看不出来
    If Len(content) = 0 Then content = "　"

    ' 字体样式（按整个单元格，不做逐字富文本）
    content = ApplyFontStyle(cell, content)

    ' 用 backcolor 模拟“单元格底色差异”
    If USE_CELL_BACKCOLOR Then
        content = WrapWithBackcolor(cell, content, rowModeColor)
    End If

    ' 对齐
    content = ApplyAlignment(cell, content)

    Dim tdOpen As String
    tdOpen = "[td=" & rs & "," & cs
    If widthPct > 0 Then tdOpen = tdOpen & "," & CStr(widthPct) & "%"
    tdOpen = tdOpen & "]"

    BuildTd = tdOpen & content & "[/td]"
End Function

'========================
' 行主色：统计该行在选区内最常见的填充色（忽略无填充）
'========================
Private Function GetRowModeFillColor(ByVal rowRng As Range) As Variant
    Dim freq As Object
    Set freq = CreateObject("Scripting.Dictionary")

    Dim cell As Range
    For Each cell In rowRng.Cells
        If HasCellFill(cell) Then
            Dim k As String
            k = CStr(cell.Interior.Color)

            If freq.Exists(k) Then
                freq(k) = CLng(freq(k)) + 1
            Else
                freq.Add k, 1
            End If
        End If
    Next cell

    If freq.Count = 0 Then
        GetRowModeFillColor = Empty
        Exit Function
    End If

    Dim bestKey As String, bestN As Long
    bestN = -1

    Dim kk As Variant
    For Each kk In freq.Keys
        If CLng(freq(kk)) > bestN Then
            bestN = CLng(freq(kk))
            bestKey = CStr(kk)
        End If
    Next kk

    GetRowModeFillColor = CLng(bestKey)
End Function

Private Function HasCellFill(ByVal cell As Range) As Boolean
    On Error GoTo SafeOut
    HasCellFill = (cell.Interior.Pattern <> xlNone) And (cell.Interior.ColorIndex <> xlColorIndexNone)
    Exit Function

SafeOut:
    HasCellFill = False
End Function

'========================
' backcolor 包裹（仅当：单元格有填充色 且 <> 行主色）
'========================
Private Function WrapWithBackcolor(ByVal cell As Range, ByVal s As String, ByVal rowModeColor As Variant) As String
    If Not HasCellFill(cell) Then
        WrapWithBackcolor = s
        Exit Function
    End If

    If Not IsEmpty(rowModeColor) Then
        If cell.Interior.Color = CLng(rowModeColor) Then
            WrapWithBackcolor = s
            Exit Function
        End If
    End If

    Dim inner As String
    inner = s
    If Len(inner) = 0 Then inner = "　"

    If BACKCOLOR_PAD_FULLWIDTH Then
        WrapWithBackcolor = "[backcolor=" & ColorLongToHex(cell.Interior.Color) & "]　" & inner & "　[/backcolor]"
    Else
        WrapWithBackcolor = "[backcolor=" & ColorLongToHex(cell.Interior.Color) & "]" & inner & "[/backcolor]"
    End If
End Function

'========================
' 对齐
'========================
Private Function ApplyAlignment(ByVal cell As Range, ByVal s As String) As String
    On Error GoTo SafeOut

    Select Case cell.HorizontalAlignment
        Case xlHAlignCenter, xlCenter
            ApplyAlignment = "[align=center]" & s & "[/align]"
        Case xlHAlignRight
            ApplyAlignment = "[align=right]" & s & "[/align]"
        Case Else
            ApplyAlignment = s ' 默认左对齐不包
    End Select

    Exit Function

SafeOut:
    ApplyAlignment = s
End Function

'========================
' 字体样式（整格）
'========================
Private Function ApplyFontStyle(ByVal cell As Range, ByVal s As String) As String
    On Error GoTo SafeOut

    Dim out As String
    out = s

    ' 颜色
    If cell.Font.Color <> 0 Then
        out = "[color=" & ColorLongToHex(cell.Font.Color) & "]" & out & "[/color]"
    End If

    ' 字号：映射到 1-7（论坛常见）
    Dim szTag As String
    szTag = FontSizeToBBSize(cell.Font.Size)
    If Len(szTag) > 0 Then
        out = "[size=" & szTag & "]" & out & "[/size]"
    End If

    ' 粗斜下划线
    If cell.Font.Bold Then out = "[b]" & out & "[/b]"
    If cell.Font.Italic Then out = "[i]" & out & "[/i]"
    If cell.Font.Underline <> xlUnderlineStyleNone Then out = "[u]" & out & "[/u]"

    ApplyFontStyle = out
    Exit Function

SafeOut:
    ApplyFontStyle = s
End Function

Private Function FontSizeToBBSize(ByVal pt As Double) As String
    ' 简单映射：你可按论坛实际效果微调
    Select Case pt
        Case Is <= 8: FontSizeToBBSize = "1"
        Case 9 To 10: FontSizeToBBSize = "2"
        Case 11 To 12: FontSizeToBBSize = "3"
        Case 13 To 14: FontSizeToBBSize = "4"
        Case 15 To 18: FontSizeToBBSize = "5"
        Case 19 To 24: FontSizeToBBSize = "6"
        Case Else: FontSizeToBBSize = "7"
    End Select
End Function

'========================
' 单元格显示文本
'========================
Private Function GetCellDisplayText(ByVal cell As Range) As String
    On Error GoTo SafeOut
    GetCellDisplayText = cell.Text
    Exit Function

SafeOut:
    GetCellDisplayText = CStr(cell.Value2)
End Function

Private Function NormalizeLineBreaksToBr(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, vbLf, "[br]")
    NormalizeLineBreaksToBr = t
End Function

'========================
' 列宽百分比（简单平均；合并列取合并宽）
'========================
Private Function GetColumnWidthPercent(ByVal cell As Range, ByVal whole As Range, ByVal colspan As Long) As Long
    Dim totalCols As Long
    totalCols = whole.Columns.Count

    If totalCols <= 0 Then
        GetColumnWidthPercent = 0
        Exit Function
    End If

    ' 平均分配（更稳；不依赖 Excel 的实际列宽）
    Dim base As Double
    base = 100# / CDbl(totalCols)

    Dim w As Double
    w = base * CDbl(colspan)

    ' 取整到 1%
    GetColumnWidthPercent = CLng(Application.WorksheetFunction.Max(1, Application.WorksheetFunction.Round(w, 0)))
End Function

'========================
' 颜色：Excel Long -> #RRGGBB
'========================
Private Function ColorLongToHex(ByVal colorLong As Long) As String
    ' Excel 内部颜色是 BGR：&H00BBGGRR
    Dim r As Long, g As Long, b As Long
    r = (colorLong And &HFF&)
    g = (colorLong \ &H100&) And &HFF&
    b = (colorLong \ &H10000) And &HFF&

    ColorLongToHex = "#" & Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function

'========================
' 少量颜色名映射（仅当 TR_COLOR_USE_HEX=False）
'========================
Private Function ClosestNamedColor(ByVal colorLong As Long) As String
    ' 仅做很粗糙的近似：white / lightgrey / lightblue / lightgreen / lightyellow / lightpink
    Dim r As Long, g As Long, b As Long
    r = (colorLong And &HFF&)
    g = (colorLong \ &H100&) And &HFF&
    b = (colorLong \ &H10000) And &HFF&

    If r > 240 And g > 240 And b > 240 Then
        ClosestNamedColor = "white"
        Exit Function
    End If

    ' 灰
    If Abs(r - g) < 15 And Abs(g - b) < 15 Then
        ClosestNamedColor = "lightgrey"
        Exit Function
    End If

    ' 主色判断
    If b >= r And b >= g Then
        ClosestNamedColor = "lightblue"
    ElseIf g >= r And g >= b Then
        ClosestNamedColor = "lightgreen"
    ElseIf r >= g And r >= b Then
        ' 红偏粉/黄简单区分
        If g > 180 Then
            ClosestNamedColor = "lightyellow"
        Else
            ClosestNamedColor = "lightpink"
        End If
    Else
        ClosestNamedColor = "white"
    End If
End Function

'========================
' 复制到剪贴板（无需勾引用：晚绑定）
'========================
Private Sub CopyTextToClipboard(ByVal text As String)
    On Error GoTo Fallback

    Dim obj As Object
    Set obj = CreateObject("MSForms.DataObject")
    obj.SetText text
    obj.PutInClipboard
    Exit Sub

Fallback:
    ' 兜底：写到一个临时单元格让你手动复制
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(1)
    ws.Range("A1").Value = text
    MsgBox "自动复制失败，已把结果写入 " & ws.Name & "!A1，请手动复制。", vbExclamation
End Sub

