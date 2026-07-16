Option Explicit

'========================
' 可调参数（按需改）
'========================
Private Const TABLE_WIDTH As String = "70%"     ' 例如 "100%" / "70%"
Private Const TABLE_BG As String = "white"      ' table 默认背景（论坛支持的颜色名或 #RRGGBB）

Private Const USE_ROW_MODE_COLOR As Boolean = True      ' 是否给 tr 上“行主色”
Private Const TR_COLOR_USE_HEX As Boolean = True        ' tr 颜色用 #RRGGBB（否则尝试映射成 lightblue 等少量名字）

Private Const USE_NESTED_CELL_TABLE As Boolean = True   ' 是否用嵌套表格模拟单个单元格背景色
Private Const NESTED_TABLE_WIDTH As String = "100%"     ' 嵌套表格铺满外层单元格
Private Const USE_DISPLAY_FORMAT As Boolean = True      ' 是否读取条件格式等实际显示出来的填充色
Private Const ESCAPE_TEXT_BRACKETS As Boolean = False   ' False 时保留 [sframe] 等单元格内的可信 BBCode

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

    If rng.Areas.Count <> 1 Then
        MsgBox "请选择一个连续的矩形区域。", vbExclamation
        Exit Sub
    End If

    If ContainsPartialMergedArea(rng) Then
        MsgBox "选区只包含了部分合并单元格，请把相关合并单元格完整选中。", vbExclamation
        Exit Sub
    End If

    If rng.Cells.CountLarge > 200000 Then
        If MsgBox("选区很大，导出可能较慢。继续吗？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If

    Dim bb As String
    bb = RangeToBBCodeTable(rng)

    CopyTextToClipboard bb
    MsgBox "已复制 BBCode 到剪贴板。去论坛直接粘贴即可。", vbInformation
End Sub

'========================
' 生成整张表
'========================
Public Function RangeToBBCodeTable(ByVal rng As Range) As String
    Dim rCount As Long, cCount As Long
    rCount = rng.Rows.Count
    cCount = rng.Columns.Count

    ' Collection + Join 避免大选区反复拼接长字符串造成性能退化
    Dim parts As Collection
    Set parts = New Collection
    parts.Add "[table=" & TABLE_WIDTH & "," & TABLE_BG & "]"

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
        parts.Add trTag

        For c = 1 To cCount
            Dim cell As Range
            Set cell = rng.Cells(r, c)

            Dim td As String
            td = BuildTd(cell, rng, rowModeColor, seen)
            If Len(td) > 0 Then parts.Add td
        Next c

        parts.Add "[/tr]"
    Next r

    parts.Add "[/table]"

    Dim output() As String
    ReDim output(0 To parts.Count - 1)

    Dim i As Long
    For i = 1 To parts.Count
        output(i - 1) = CStr(parts(i))
    Next i

    RangeToBBCodeTable = Join(output, vbCrLf)
End Function

Private Function ContainsPartialMergedArea(ByVal rng As Range) As Boolean
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim cell As Range
    For Each cell In rng.Cells
        If cell.MergeCells Then
            Dim ma As Range
            Set ma = cell.MergeArea

            Dim key As String
            key = ma.Address(External:=True)

            If Not seen.Exists(key) Then
                seen.Add key, True

                Dim overlap As Range
                Set overlap = Application.Intersect(rng, ma)
                If overlap Is Nothing Then
                    ContainsPartialMergedArea = True
                    Exit Function
                End If

                If overlap.Cells.CountLarge <> ma.Cells.CountLarge Then
                    ContainsPartialMergedArea = True
                    Exit Function
                End If
            End If
        End If
    Next cell
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
' td 构建（含合并单元格、对齐、字体、嵌套背景表格）
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

    Dim preserveBBCode As Boolean
    preserveBBCode = (Not ESCAPE_TEXT_BRACKETS) And ContainsPairedBBCode(content)

    If ESCAPE_TEXT_BRACKETS Then
        content = Replace(content, "[", "&#91;")
        content = Replace(content, "]", "&#93;")
    End If

    ' Keylol 不识别 [br]，保留为论坛能够渲染的真实换行
    content = NormalizeLineBreaks(content)

    ' 空内容给个不可见占位，确保背景色表格仍有可见高度
    If Len(content) = 0 Then content = "　"

    ' 已有论坛标签原样保留，避免 [sframe] 被 size/b 等样式标签包裹
    If Not preserveBBCode Then content = ApplyFontStyle(cell, content)

    ' 对齐放在嵌套表格里面，避免居中时把整张嵌套表格缩成内容宽度
    content = ApplyAlignment(cell, content)

    ' Keylol 的 td 不支持单独背景色，用铺满单元格的嵌套表格模拟
    Dim overrideBg As String
    overrideBg = GetCellBackgroundOverride(cell, rowModeColor)
    content = WrapWithNestedTable(content, overrideBg)

    Dim tdOpen As String
    ' Discuz/Keylol 的参数顺序是 colspan,rowspan,width
    tdOpen = "[td=" & cs & "," & rs
    If widthPct > 0 Then tdOpen = tdOpen & "," & CStr(widthPct) & "%"
    tdOpen = tdOpen & "]"

    BuildTd = tdOpen & content & "[/td]"
End Function

Private Function ContainsPairedBBCode(ByVal s As String) As Boolean
    Dim openPos As Long, closePos As Long
    openPos = InStr(1, s, "[", vbBinaryCompare)
    closePos = InStr(1, s, "[/", vbBinaryCompare)

    ContainsPairedBBCode = (openPos > 0) _
                           And (InStr(openPos + 1, s, "]", vbBinaryCompare) > 0) _
                           And (closePos > openPos)
End Function

'========================
' 行主色：仅当某种填充色占该行严格多数时采用
' 无填充也参与统计，避免一个孤立色块被误判成整行底色
'========================
Private Function GetRowModeFillColor(ByVal rowRng As Range) As Variant
    Dim freq As Object
    Set freq = CreateObject("Scripting.Dictionary")

    Dim cell As Range
    Dim totalCells As Long
    totalCells = 0

    For Each cell In rowRng.Cells
        totalCells = totalCells + 1

        Dim fillColor As Variant
        fillColor = GetCellFillColor(cell)

        If Not IsEmpty(fillColor) Then
            Dim k As String
            k = CStr(CLng(fillColor))

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

    If bestN * 2 <= totalCells Then
        GetRowModeFillColor = Empty
    Else
        GetRowModeFillColor = CLng(bestKey)
    End If
End Function

Private Function GetCellFillColor(ByVal cell As Range) As Variant
    If USE_DISPLAY_FORMAT Then
        On Error GoTo UseInterior

        If cell.DisplayFormat.Interior.Pattern <> xlNone _
           And cell.DisplayFormat.Interior.ColorIndex <> xlColorIndexNone Then
            GetCellFillColor = CLng(cell.DisplayFormat.Interior.Color)
        Else
            GetCellFillColor = Empty
        End If
        Exit Function
    End If

UseInterior:
    On Error GoTo NoFill

    If cell.Interior.Pattern <> xlNone _
       And cell.Interior.ColorIndex <> xlColorIndexNone Then
        GetCellFillColor = CLng(cell.Interior.Color)
    Else
        GetCellFillColor = Empty
    End If
    Exit Function

NoFill:
    GetCellFillColor = Empty
End Function

'========================
' 计算单元格相对行背景需要覆盖的颜色
'========================
Private Function GetCellBackgroundOverride(ByVal cell As Range, ByVal rowModeColor As Variant) As String
    Dim fillColor As Variant
    fillColor = GetCellFillColor(cell)

    If IsEmpty(fillColor) Then
        ' tr 有底色时，无填充单元格需要恢复表格默认背景
        If Not IsEmpty(rowModeColor) Then GetCellBackgroundOverride = TABLE_BG
        Exit Function
    End If

    If Not IsEmpty(rowModeColor) Then
        If CLng(fillColor) = CLng(rowModeColor) Then Exit Function
    End If

    GetCellBackgroundOverride = ColorLongToHex(CLng(fillColor))
End Function

Private Function WrapWithNestedTable(ByVal s As String, ByVal bgColor As String) As String
    If Not USE_NESTED_CELL_TABLE Or Len(bgColor) = 0 Then
        WrapWithNestedTable = s
        Exit Function
    End If

    WrapWithNestedTable = "[table=" & NESTED_TABLE_WIDTH & "," & bgColor & "]" _
                          & "[tr][td]" & s & "[/td][/tr][/table]"
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
        Case Is <= 10: FontSizeToBBSize = "2"
        Case Is <= 12: FontSizeToBBSize = "3"
        Case Is <= 14: FontSizeToBBSize = "4"
        Case Is <= 18: FontSizeToBBSize = "5"
        Case Is <= 24: FontSizeToBBSize = "6"
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

Private Function NormalizeLineBreaks(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, vbLf, vbCrLf)
    NormalizeLineBreaks = t
End Function

'========================
' 列宽百分比（按 Excel 实际列宽；合并列取各列宽度之和）
'========================
Private Function GetColumnWidthPercent(ByVal cell As Range, ByVal whole As Range, ByVal colspan As Long) As Long
    Dim totalCols As Long
    totalCols = whole.Columns.Count

    If totalCols <= 0 Then
        GetColumnWidthPercent = 0
        Exit Function
    End If

    On Error GoTo EqualWidth

    Dim totalWidth As Double, targetWidth As Double
    Dim i As Long, firstIndex As Long, lastIndex As Long

    For i = 1 To totalCols
        totalWidth = totalWidth + CDbl(whole.Columns(i).ColumnWidth)
    Next i

    firstIndex = cell.Column - whole.Column + 1
    lastIndex = firstIndex + colspan - 1
    If firstIndex < 1 Then firstIndex = 1
    If lastIndex > totalCols Then lastIndex = totalCols

    For i = firstIndex To lastIndex
        targetWidth = targetWidth + CDbl(whole.Columns(i).ColumnWidth)
    Next i

    If totalWidth <= 0 Or targetWidth <= 0 Then GoTo EqualWidth

    Dim w As Double
    w = 100# * targetWidth / totalWidth

    GetColumnWidthPercent = CLng(Int(w + 0.5))
    If GetColumnWidthPercent < 1 Then GetColumnWidthPercent = 1
    If GetColumnWidthPercent > 100 Then GetColumnWidthPercent = 100
    Exit Function

EqualWidth:
    GetColumnWidthPercent = CLng(Int((100# * CDbl(colspan) / CDbl(totalCols)) + 0.5))
    If GetColumnWidthPercent < 1 Then GetColumnWidthPercent = 1
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
    ' 兜底：新建工作簿，避免覆盖当前工作簿中的 A1
    Dim ws As Worksheet
    Set ws = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    ws.Range("A1").Value = text
    MsgBox "自动复制失败，已新建工作簿并把结果写入 A1，请手动复制后关闭该工作簿。", vbExclamation
End Sub
