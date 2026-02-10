Attribute VB_Name = "MathdokuCandidates"
Option Explicit

' No "hidden" error handlers: unexpected errors show details.
Private Sub ShowError(ByVal procName As String)
    MsgBox "Error in " & procName & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Source: " & Err.Source, _
           vbCritical, "Mathdoku"
End Sub

Private Sub ApplyValueStyle(ByVal shp As Shape, ByVal fontName As String, ByVal fontSize As Single, ByVal fontColorRGB As Long, ByVal boldValue As MsoTriState)
    With shp.TextFrame.TextRange
        .ParagraphFormat.Alignment = ppAlignCenter
        .Font.Name = fontName
        .Font.Size = fontSize
        .Font.Color.RGB = fontColorRGB
        .Font.Bold = boldValue
    End With
End Sub

Private Sub ApplyCandidatesStyle(ByVal shp As Shape, ByVal fontName As String, ByVal fontSize As Single, ByVal fontColorRGB As Long)
    With shp.TextFrame.TextRange
        .ParagraphFormat.Alignment = ppAlignLeft
        .Font.Name = fontName
        .Font.Size = fontSize
        .Font.Color.RGB = fontColorRGB
        .Font.Bold = msoFalse
    End With
End Sub

' Import into PowerPoint VBA:
'   Alt+F11 -> File -> Import File... -> select this .bas
' Then add these two macros to the Quick Access Toolbar:
'   - EnterFinalValue
'   - EditCellCandidates
'
' What it does:
'   - Lets you solve directly on the slide using hotkeys:
'       1) Select a cell (click it or Tab to it)
'       2) EnterFinalValue -> prompt -> sets value, clears candidates, removes that value
'          from candidates in the same row and column (one undo step)
'       3) EditCellCandidates -> prompt prefilled with current candidates -> normalizes
'          (dedupe + sort) and formats like the Mathdoku app (one undo step)
'
' Requirements (works with the PPTX created by your generator):
'   - Each cell has shapes named:
'       VALUE_A1         (full-cell textbox)
'       CANDIDATES_A1    (candidates textbox for the same cell)
'       CELL_A1          (click/tab hitbox for the cell)
'   - Required hidden metadata shape named `MATHDOKU_META` with a line `size: N`
'
' Note: PowerPoint can’t auto-format on every keystroke without an add-in.
' This is designed as: type freely -> press the hotkey.

Private Function GetGridSizeFromSlide(ByVal sld As Slide) As Long
    Dim meta As Shape
    Set meta = GetShapeByName(sld, "MATHDOKU_META")
    If meta Is Nothing Then
        MsgBox "Missing required shape 'MATHDOKU_META'. Re-generate the slide/template.", vbExclamation, "Mathdoku"
        GetGridSizeFromSlide = 0
        Exit Function
    End If

    If meta.HasTextFrame = msoFalse Or meta.TextFrame.HasText = msoFalse Then
        MsgBox "Shape 'MATHDOKU_META' has no text. Re-generate the slide/template.", vbExclamation, "Mathdoku"
        GetGridSizeFromSlide = 0
        Exit Function
    End If

    Dim t As String, lines() As String, i As Long, ln As String
    t = meta.TextFrame.TextRange.Text
    ' PowerPoint often uses Vertical Tab (Chr(11)) as a line separator.
    t = Replace(t, Chr$(11), vbLf)
    t = Replace(t, vbLf, vbCrLf)
    lines = Split(t, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If LCase$(Left$(ln, 5)) = "size:" Then
            ln = Trim$(Mid$(ln, 6))
            If IsNumeric(ln) Then
                GetGridSizeFromSlide = CLng(ln)
                Exit Function
            End If
        End If
    Next i

    MsgBox "Shape 'MATHDOKU_META' is missing a valid 'size: N' line.", vbExclamation, "Mathdoku"
    GetGridSizeFromSlide = 0
End Function

Private Function DigitsOnly(ByVal text As String) As String
    Dim i As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then
            out = out & ch
        End If
    Next i
    DigitsOnly = out
End Function

Private Function FormatLikeApp(ByVal digits As String, ByVal gridSize As Long) As String
    ' For size N, app-like candidates are split across 2 lines:
    '   line1: ceil(N/2) digits
    '   line2: remaining digits
    ' Digits are separated by a single space for consistent placement.

    Dim perLine As Long
    If gridSize <= 0 Then gridSize = 9
    perLine = (gridSize + 1) \ 2 ' ceil(N/2)
    If perLine <= 0 Then perLine = 5

    Dim d As String
    d = digits
    If Len(d) <= perLine Then
        FormatLikeApp = JoinDigitsWithSpaces(d)
        Exit Function
    End If

    Dim line1 As String, line2 As String
    line1 = Left$(d, perLine)
    line2 = Mid$(d, perLine + 1)

    FormatLikeApp = JoinDigitsWithSpaces(line1) & vbCrLf & JoinDigitsWithSpaces(line2)
End Function

Private Function JoinDigitsWithSpaces(ByVal digits As String) As String
    Dim i As Long, out As String
    out = ""
    For i = 1 To Len(digits)
        out = out & Mid$(digits, i, 1)
        If i < Len(digits) Then out = out & " "
    Next i
    JoinDigitsWithSpaces = out
End Function

Private Function ActiveSlide() As Slide
    On Error GoTo ErrHandler

    If ActiveWindow Is Nothing Then
        Set ActiveSlide = Nothing
        Exit Function
    End If

    Set ActiveSlide = ActiveWindow.View.Slide
    Exit Function

ErrHandler:
    ShowError "ActiveSlide"
    Set ActiveSlide = Nothing
End Function

Private Function ParseA1FromName(ByVal nm As String, ByRef r As Long, ByRef c As Long) As Boolean
    ' Extract trailing cell ref like "_A1" from a shape name.
    Dim i As Long, token As String, colCh As String, rowPart As String

    i = InStrRev(nm, "_")
    If i = 0 Then
        ParseA1FromName = False
        Exit Function
    End If

    token = Mid$(nm, i + 1)
    If Len(token) < 2 Then
        ParseA1FromName = False
        Exit Function
    End If

    colCh = UCase$(Left$(token, 1))
    If colCh < "A" Or colCh > "Z" Then
        ParseA1FromName = False
        Exit Function
    End If

    rowPart = Mid$(token, 2)
    If Not IsNumeric(rowPart) Then
        ParseA1FromName = False
        Exit Function
    End If

    c = Asc(colCh) - Asc("A") + 1
    r = CLng(rowPart)
    If r <= 0 Or c <= 0 Then
        ParseA1FromName = False
        Exit Function
    End If

    ParseA1FromName = True
End Function

Private Function CellRefA1(ByVal r As Long, ByVal c As Long) As String
    CellRefA1 = Chr$(Asc("A") + c - 1) & CStr(r)
End Function

Private Function GetSelectedCellRC(ByRef r As Long, ByRef c As Long) As Boolean
    On Error GoTo Fail

    If ActiveWindow Is Nothing Then GoTo Fail
    If ActiveWindow.Selection Is Nothing Then GoTo Fail

    Dim shp As Shape
    Set shp = Nothing

    If ActiveWindow.Selection.Type = ppSelectionShapes Then
        If ActiveWindow.Selection.ShapeRange.Count <> 1 Then GoTo Fail
        Set shp = ActiveWindow.Selection.ShapeRange(1)
    ElseIf ActiveWindow.Selection.Type = ppSelectionText Then
        ' When editing text, selection is text; get its parent shape.
        Set shp = ActiveWindow.Selection.TextRange.Parent
    Else
        GoTo Fail
    End If

    If shp Is Nothing Then GoTo Fail

    ' Direct parse from shape name (CELL_/VALUE_/CANDIDATES_)
    GetSelectedCellRC = ParseA1FromName(shp.Name, r, c)
    Exit Function

Fail:
    GetSelectedCellRC = False
End Function

Private Function GetShapeByName(ByVal sld As Slide, ByVal nm As String) As Shape
    ' IMPORTANT: do NOT iterate sld.Shapes. Some locked shapes can throw Err 440
    ' during enumeration. Direct lookup by name is safer.
    On Error GoTo ErrHandler
    Set GetShapeByName = sld.Shapes(nm)
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "Mathdoku.GetShapeByName", "Cannot access shape '" & nm & "': " & Err.Description
End Function

Private Function NormalizeCandidatesDigits(ByVal raw As String, ByVal gridSize As Long) As String
    Dim d As String
    d = DigitsOnly(raw)

    Dim present(1 To 9) As Boolean
    Dim i As Long, v As Long, out As String
    out = ""

    For i = 1 To Len(d)
        v = CLng(Mid$(d, i, 1))
        If v >= 1 And v <= gridSize And v <= 9 Then
            present(v) = True
        End If
    Next i

    For v = 1 To 9
        If v > gridSize Then Exit For
        If present(v) Then out = out & CStr(v)
    Next v

    NormalizeCandidatesDigits = out
End Function

Private Sub RemoveCandidateDigit(ByVal sld As Slide, ByVal r As Long, ByVal c As Long, ByVal digit As Long, ByVal gridSize As Long)
    Dim nm As String
    nm = "CANDIDATES_" & CellRefA1(r, c)

    Dim shp As Shape
    Set shp = GetShapeByName(sld, nm)
    If shp.HasTextFrame = msoFalse Then Exit Sub

    Dim candFontName As String, candFontSize As Single, candFontColor As Long
    candFontName = shp.TextFrame.TextRange.Font.Name
    candFontSize = shp.TextFrame.TextRange.Font.Size
    candFontColor = shp.TextFrame.TextRange.Font.Color.RGB

    Dim currentDigits As String
    currentDigits = NormalizeCandidatesDigits(shp.TextFrame.TextRange.Text, gridSize)
    If Len(currentDigits) = 0 Then Exit Sub

    Dim i As Long, out As String, ch As String
    out = ""
    For i = 1 To Len(currentDigits)
        ch = Mid$(currentDigits, i, 1)
        If CLng(ch) <> digit Then out = out & ch
    Next i

    If Len(out) = 0 Then
        shp.TextFrame.TextRange.Text = " "
    Else
        shp.TextFrame.TextRange.Text = FormatLikeApp(out, gridSize)
    End If
    ApplyCandidatesStyle shp, candFontName, candFontSize, candFontColor
End Sub

Public Sub EnterFinalValue()
    On Error GoTo ErrHandler

    Dim r As Long, c As Long
    If Not GetSelectedCellRC(r, c) Then
        MsgBox "Select a cell first (click inside a cell, or Tab between cells).", vbInformation, "Mathdoku"
        Exit Sub
    End If

    Dim sld As Slide
    Set sld = ActiveSlide()
    If sld Is Nothing Then Exit Sub

    Dim sz As Long
    sz = GetGridSizeFromSlide(sld)
    If sz <= 0 Then Exit Sub

    Dim valueShp As Shape, candShp As Shape
    Dim cellRef As String
    cellRef = CellRefA1(r, c)
    Set valueShp = GetShapeByName(sld, "VALUE_" & cellRef)
    Set candShp = GetShapeByName(sld, "CANDIDATES_" & cellRef)

    Dim valueFontName As String, valueFontSize As Single, valueFontColor As Long, valueBold As MsoTriState
    valueFontName = valueShp.TextFrame.TextRange.Font.Name
    valueFontSize = valueShp.TextFrame.TextRange.Font.Size
    valueFontColor = valueShp.TextFrame.TextRange.Font.Color.RGB
    valueBold = valueShp.TextFrame.TextRange.Font.Bold

    Dim candFontName As String, candFontSize As Single, candFontColor As Long
    candFontName = candShp.TextFrame.TextRange.Font.Name
    candFontSize = candShp.TextFrame.TextRange.Font.Size
    candFontColor = candShp.TextFrame.TextRange.Font.Color.RGB

    Dim currentValue As String
    currentValue = ""
    If valueShp.HasTextFrame Then currentValue = Trim$(valueShp.TextFrame.TextRange.Text)

    Dim inputVal As String
    inputVal = InputBox("Final value for cell " & cellRef & " (1-" & sz & "):", "Mathdoku", currentValue)
    If Len(inputVal) = 0 Then Exit Sub ' cancel/empty -> no-op

    Dim d As String
    d = DigitsOnly(inputVal)
    If Len(d) = 0 Then Exit Sub

    Dim v As Long
    v = CLng(Left$(d, 1))
    If v < 1 Or v > sz Or v > 9 Then
        MsgBox "Value must be 1-" & sz & ".", vbExclamation, "Mathdoku"
        Exit Sub
    End If

    Application.StartNewUndoEntry

    ' Set final value
    valueShp.TextFrame.TextRange.Text = CStr(v)
    ApplyValueStyle valueShp, valueFontName, valueFontSize, valueFontColor, valueBold

    ' Clear candidates in this cell
    If Not candShp Is Nothing Then
        candShp.TextFrame.TextRange.Text = " "
        ApplyCandidatesStyle candShp, candFontName, candFontSize, candFontColor
    End If

    ' Remove this candidate from same row and column
    Dim i As Long
    For i = 1 To sz
        If i <> c Then RemoveCandidateDigit sld, r, i, v, sz
        If i <> r Then RemoveCandidateDigit sld, i, c, v, sz
    Next i

    Exit Sub

ErrHandler:
    ShowError "EnterFinalValue"
End Sub

Public Sub EditCellCandidates()
    On Error GoTo ErrHandler

    Dim r As Long, c As Long
    If Not GetSelectedCellRC(r, c) Then
        MsgBox "Select a cell first (click inside a cell, or Tab between cells).", vbInformation, "Mathdoku"
        Exit Sub
    End If

    Dim sld As Slide
    Set sld = ActiveSlide()
    If sld Is Nothing Then Exit Sub

    Dim sz As Long
    sz = GetGridSizeFromSlide(sld)
    If sz <= 0 Then Exit Sub

    Dim candShp As Shape
    Dim cellRef As String
    cellRef = CellRefA1(r, c)
    Set candShp = GetShapeByName(sld, "CANDIDATES_" & cellRef)

    Dim candFontName As String, candFontSize As Single, candFontColor As Long
    candFontName = candShp.TextFrame.TextRange.Font.Name
    candFontSize = candShp.TextFrame.TextRange.Font.Size
    candFontColor = candShp.TextFrame.TextRange.Font.Color.RGB

    Dim currentDigits As String
    currentDigits = NormalizeCandidatesDigits(candShp.TextFrame.TextRange.Text, sz)

    Dim inputVal As String
    inputVal = InputBox("Candidates for cell " & cellRef & " (any order; duplicates ok):", "Mathdoku", currentDigits)
    If Len(inputVal) = 0 Then Exit Sub ' cancel/empty -> no-op

    Dim normalized As String
    normalized = NormalizeCandidatesDigits(inputVal, sz)

    Application.StartNewUndoEntry

    If Len(normalized) = 0 Then
        candShp.TextFrame.TextRange.Text = " "
    Else
        candShp.TextFrame.TextRange.Text = FormatLikeApp(normalized, sz)
    End If
    ApplyCandidatesStyle candShp, candFontName, candFontSize, candFontColor

    Exit Sub

ErrHandler:
    ShowError "EditCellCandidates"
End Sub

