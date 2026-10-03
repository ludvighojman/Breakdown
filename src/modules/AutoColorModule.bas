Option Explicit

Private Const NAME_PREFIX As String = "AutoColor_"

' Input cell format choices (see InputFillName and InputBorderName)
Public Const INPUT_FILL_COUNT As Long = 10      ' Fills 1 to 10; 0 is no fill
Public Const INPUT_BORDER_COUNT As Long = 5     ' Borders 1 to 5; 0 is no border
Public Const NO_FILL As Long = -1
Public Const DEFAULT_INPUT_FILL As Long = 1     ' Light yellow
Public Const DEFAULT_INPUT_BORDER As Long = 1   ' Thin light grey

' The input cell preview in Settings is drawn in screen pixels, the way Excel
' draws a cell at 100% zoom: PREVIEW_SCALE pixels per point (a Retina screen
' on Mac), and lines LINE_PIXELS wide
#If Mac Then
Public Const PREVIEW_SCALE As Single = 2
Private Const LINE_PIXELS As Long = 2
#Else
Public Const PREVIEW_SCALE As Single = 4 / 3
Private Const LINE_PIXELS As Long = 1
#End If

' Function to get color from saved settings or return default
Private Function GetSavedColor(colorName As String, defaultColor As Long) As Long
    On Error Resume Next
    Dim colorValue As String
    colorValue = ThisWorkbook.Names(NAME_PREFIX & colorName).RefersTo
    If Err.Number = 0 And colorValue <> "" Then
        GetSavedColor = CLng(Mid(colorValue, 2)) ' Remove the = sign
    Else
        GetSavedColor = defaultColor
    End If
    On Error GoTo 0
End Function

Public Sub AutoColorCells()
    If TypeName(Selection) <> "Range" Then Exit Sub

    ' Saved colors, or the financial-modelling convention by default
    Dim colorInput As Long: colorInput = GetSavedColor("Input", 16711680)                       ' Blue #0000FF: hard-coded inputs
    Dim colorFormula As Long: colorFormula = GetSavedColor("Formula", 0)                        ' Black: formulas on the same sheet
    Dim colorWorksheetLink As Long: colorWorksheetLink = GetSavedColor("WorksheetLink", 32768)  ' Green #008000: formulas using other sheets
    Dim colorExternal As Long: colorExternal = GetSavedColor("ExternalLink", 255)              ' Red #FF0000: other workbooks, external data, errors

    ' Only cells with content (constants or formulas), excluding blanks
    Dim rng As Range
    Dim usedCells As Range
    Dim formulaCells As Range
    Set rng = Selection
    On Error Resume Next
    Set usedCells = rng.SpecialCells(xlCellTypeConstants)
    Set formulaCells = rng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If usedCells Is Nothing Then
        Set usedCells = formulaCells
    ElseIf Not formulaCells Is Nothing Then
        Set usedCells = Union(usedCells, formulaCells)
    End If
    If usedCells Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    Dim cell As Range
    Dim formulaText As String
    Dim inputCells As Range
    For Each cell In usedCells
        If IsErrorValue(cell) Then
            ColorNonInput cell, colorExternal, colorInput
        ElseIf cell.HasFormula Then
            ' Text inside "quotes" is not part of any reference
            formulaText = WithoutStringLiterals(cell.Formula)
            If IsWorkbookLink(formulaText) Or IsExternalData(formulaText) Then
                ColorNonInput cell, colorExternal, colorInput
            ElseIf InStr(1, formulaText, "!") > 0 Then
                ColorNonInput cell, colorWorksheetLink, colorInput
            ElseIf IsInput(cell, formulaText) Then
                AddCell inputCells, cell
            Else
                ColorNonInput cell, colorFormula, colorInput
            End If
        ElseIf cell.Hyperlinks.Count > 0 Then
            ' Hyperlinks keep their own formatting
        ElseIf IsInput(cell, "") Then
            AddCell inputCells, cell
        End If
    Next cell

    ' Inputs last, so taking the format off a former input next to one can't
    ' remove the border they share
    If Not inputCells Is Nothing Then FormatInputs inputCells, colorInput

    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------------------
' Input cell format: one of the fills and one of the borders below, chosen in
' Settings > Auto-Color. Choice 0 keeps the cell's own fill or borders. Both
' are saved by number, so add new choices at the end of each list.
' ---------------------------------------------------------------------------

Public Function InputFillName(ByVal fillIndex As Long) As String
    InputFillName = Array("No fill (keep the cell's own)", "Light yellow", "Pale yellow", "Cream", _
                          "Peach (Excel's Input style)", "Light orange", "Light blue", "Blue-grey", _
                          "Light turquoise", "Light green", "Light grey")(fillIndex)
End Function

' The fill's RGB() colour, or NO_FILL
Public Function InputFillColor(ByVal fillIndex As Long) As Long
    InputFillColor = Array(NO_FILL, RGB(255, 255, 153), RGB(255, 242, 204), RGB(255, 255, 204), _
                           RGB(255, 204, 153), RGB(252, 228, 214), RGB(221, 235, 247), RGB(231, 240, 243), _
                           RGB(204, 255, 255), RGB(226, 239, 218), RGB(242, 242, 242))(fillIndex)
End Function

Public Function InputBorderName(ByVal borderIndex As Long) As String
    InputBorderName = Array("No border (keep the cell's own)", "Thin light grey", "Thin grey", "Thin black", _
                            "Dotted grey", "Dashed grey")(borderIndex)
End Function

' A border's line style (xlNone for choice 0), weight and colour
Public Sub GetInputBorder(ByVal borderIndex As Long, edgeStyle As Long, edgeWeight As Long, edgeColor As Long)
    edgeStyle = xlContinuous
    edgeWeight = xlThin
    Select Case borderIndex
        Case 1: edgeColor = RGB(217, 217, 217)
        Case 2: edgeColor = RGB(128, 128, 128)
        Case 3: edgeColor = RGB(0, 0, 0)
        Case 4: edgeColor = RGB(128, 128, 128): edgeWeight = xlHairline   ' Excel draws a hairline dotted
        Case 5: edgeColor = RGB(128, 128, 128): edgeStyle = xlDash
        Case Else: edgeStyle = xlNone
    End Select
End Sub

Public Function SavedInputFill() As Long
    SavedInputFill = SavedChoice("InputFill", DEFAULT_INPUT_FILL, INPUT_FILL_COUNT)
End Function

Public Function SavedInputBorder() As Long
    SavedInputBorder = SavedChoice("InputBorder", DEFAULT_INPUT_BORDER, INPUT_BORDER_COUNT)
End Function

Public Sub SaveInputFormat(ByVal fillIndex As Long, ByVal borderIndex As Long)
    SaveSetting "Breakdown", "AutoColor", "InputFill", CStr(fillIndex)
    SaveSetting "Breakdown", "AutoColor", "InputBorder", CStr(borderIndex)
End Sub

' A saved choice from 0 to lastChoice, or defaultChoice
Private Function SavedChoice(settingName As String, ByVal defaultChoice As Long, ByVal lastChoice As Long) As Long
    Dim saved As String
    saved = GetSetting("Breakdown", "AutoColor", settingName, "")
    SavedChoice = defaultChoice
    If saved Like "#" Or saved Like "##" Then
        If CLng(saved) <= lastChoice Then SavedChoice = CLng(saved)
    End If
End Function

' A picture of a cell with this fill and border, pixelsWide x pixelsHigh, for
' the preview in Settings
Public Function InputCellPicture(ByVal pixelsWide As Long, ByVal pixelsHigh As Long, _
                                 ByVal fillIndex As Long, ByVal borderIndex As Long) As Object
    Dim bytes() As Byte
    Dim fillColor As Long
    Dim edgeStyle As Long
    Dim edgeWeight As Long
    Dim edgeColor As Long
    Dim pixelColor As Long
    Dim x As Long
    Dim y As Long

    fillColor = InputFillColor(fillIndex)
    If fillColor = NO_FILL Then fillColor = RGB(255, 255, 255)
    GetInputBorder borderIndex, edgeStyle, edgeWeight, edgeColor
    bytes = TraceUtils.NewBitmap(pixelsWide, pixelsHigh)
    For y = 0 To pixelsHigh - 1
        For x = 0 To pixelsWide - 1
            If edgeStyle = xlNone Then
                pixelColor = fillColor
            ElseIf x < LINE_PIXELS Or y < LINE_PIXELS Or x >= pixelsWide - LINE_PIXELS Or y >= pixelsHigh - LINE_PIXELS Then
                pixelColor = LinePixel(edgeStyle, edgeWeight, edgeColor, x, y)
            Else
                pixelColor = fillColor
            End If
            TraceUtils.SetPixel bytes, x, y, pixelColor
        Next x
    Next y
    Set InputCellPicture = TraceUtils.BitmapPicture(bytes, "breakdown-cell-v3-" & fillIndex & "-" & borderIndex & "-" & _
                                                    pixelsWide & "x" & pixelsHigh & ".bmp")
End Function

' One pixel (x, y) of a cell border. Excel draws dotted and dashed lines as
' diagonal patterns of screen pixels (measured in Excel for Mac at 100% zoom):
' a dotted hairline as a checkerboard, a dash as 3 pixels on and 1 off. The
' gaps show the white sheet, not the cell's fill.
Private Function LinePixel(ByVal edgeStyle As Long, ByVal edgeWeight As Long, ByVal edgeColor As Long, _
                           ByVal x As Long, ByVal y As Long) As Long
    Dim covered As Boolean
    If edgeStyle = xlDash Then
        covered = ((x + y) Mod 4 <> 1)
    ElseIf edgeWeight = xlHairline Then
        covered = ((x + y) Mod 2 = 0)
    Else
        covered = True
    End If
    If covered Then LinePixel = edgeColor Else LinePixel = RGB(255, 255, 255)
End Function

Private Sub AddCell(target As Range, cell As Range)
    If target Is Nothing Then Set target = cell Else Set target = Union(target, cell)
End Sub

' The input colour, and the chosen fill and border around every cell
Private Sub FormatInputs(inputCells As Range, fontColor As Long)
    Dim fillColor As Long
    Dim edgeStyle As Long
    Dim edgeWeight As Long
    Dim edgeColor As Long
    Dim area As Range
    fillColor = InputFillColor(SavedInputFill())
    GetInputBorder SavedInputBorder(), edgeStyle, edgeWeight, edgeColor
    For Each area In inputCells.Areas
        area.Font.Color = fontColor
        If fillColor <> NO_FILL Then
            area.Interior.Pattern = xlSolid
            area.Interior.Color = fillColor
        End If
        If edgeStyle <> xlNone Then
            With area.Borders
                .LineStyle = edgeStyle
                .Weight = edgeWeight
                .Color = edgeColor
            End With
        End If
    Next area
End Sub

' A cell that is not an input: its colour. A cell still in the input colour
' was an input before, and loses its input format too.
Private Sub ColorNonInput(cell As Range, fontColor As Long, inputColor As Long)
    If cell.Font.Color = inputColor And fontColor <> inputColor Then ClearInputFormat cell, inputColor
    cell.Font.Color = fontColor
End Sub

' Takes the input format off a former input: a fill from the list, and every
' border in one of the listed styles, except one shared with a neighbour
' still in the input colour. Other fills and borders stay.
Private Sub ClearInputFormat(cell As Range, inputColor As Long)
    Dim fillIndex As Long
    Dim edge As Variant
    If cell.Interior.Pattern = xlSolid Then
        For fillIndex = 1 To INPUT_FILL_COUNT
            If cell.Interior.Color = InputFillColor(fillIndex) Then
                cell.Interior.Pattern = xlNone
                Exit For
            End If
        Next fillIndex
    End If
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        If IsInputBorder(cell.Borders(edge)) Then
            If Not NeighbourInColor(cell, CLng(edge), inputColor) Then cell.Borders(edge).LineStyle = xlNone
        End If
    Next edge
End Sub

' True if a cell border has one of the listed styles
Private Function IsInputBorder(cellBorder As Border) As Boolean
    Dim borderIndex As Long
    Dim edgeStyle As Long
    Dim edgeWeight As Long
    Dim edgeColor As Long
    If cellBorder.LineStyle = xlNone Then Exit Function
    For borderIndex = 1 To INPUT_BORDER_COUNT
        GetInputBorder borderIndex, edgeStyle, edgeWeight, edgeColor
        If cellBorder.LineStyle = edgeStyle And cellBorder.Weight = edgeWeight And cellBorder.Color = edgeColor Then
            IsInputBorder = True
            Exit Function
        End If
    Next borderIndex
End Function

' True if the cell across the given edge has text in the colour fontColor
Private Function NeighbourInColor(cell As Range, edge As Long, fontColor As Long) As Boolean
    Dim rowStep As Long
    Dim columnStep As Long
    Dim neighbour As Range
    Select Case edge
        Case xlEdgeLeft: columnStep = -1
        Case xlEdgeTop: rowStep = -1
        Case xlEdgeBottom: rowStep = 1
        Case xlEdgeRight: columnStep = 1
    End Select

    On Error Resume Next    ' No neighbour past the edge of the sheet
    Set neighbour = cell.Offset(rowStep, columnStep)
    On Error GoTo 0
    If neighbour Is Nothing Then Exit Function
    If neighbour.Font.Color = fontColor Then NeighbourInColor = True
End Function

' Any error value except #N/A, which models often produce on purpose with NA()
Private Function IsErrorValue(cell As Range) As Boolean
    If Not IsError(cell.Value) Then Exit Function
    IsErrorValue = (cell.Value <> CVErr(xlErrNA))
End Function

' A reference to another workbook: "[Book.xlsx]Sheet1!A1",
' "'C:\Path\[My Book.xlsx]P&L'!A1" or "[Book.xlsx]!Name". A table reference
' such as "Table1[Sales]" is not: no sheet name and "!" follow its "]".
Private Function IsWorkbookLink(formulaText As String) As Boolean
    Dim closePos As Long
    Dim k As Long
    Dim ch As String
    Dim quoted As Boolean

    closePos = InStr(1, formulaText, "]")
    Do While closePos > 0
        quoted = InQuotedName(formulaText, closePos)
        k = closePos + 1
        Do While k <= Len(formulaText)
            ch = Mid$(formulaText, k, 1)
            If quoted Then
                ' Any character up to the closing quote ('' is an escaped quote)
                If ch = "'" Then
                    If Mid$(formulaText, k + 1, 1) = "!" Then
                        IsWorkbookLink = True
                        Exit Function
                    ElseIf Mid$(formulaText, k + 1, 1) = "'" Then
                        k = k + 1
                    Else
                        Exit Do
                    End If
                End If
            Else
                If ch = "!" Then
                    IsWorkbookLink = True
                    Exit Function
                End If
                If Not IsSheetNameChar(ch) Then Exit Do
            End If
            k = k + 1
        Loop
        closePos = InStr(closePos + 1, formulaText, "]")
    Loop
End Function

' A character of an unquoted sheet name
Private Function IsSheetNameChar(ch As String) As Boolean
    IsSheetNameChar = (ch Like "[A-Za-z0-9_.]") Or AscW(ch) < 0 Or AscW(ch) > 127
End Function

' True if position pos is inside a 'quoted sheet name'
Private Function InQuotedName(formulaText As String, ByVal pos As Long) As Boolean
    Dim i As Long
    For i = 1 To pos - 1
        If Mid$(formulaText, i, 1) = "'" Then InQuotedName = Not InQuotedName
    Next i
End Function

' Functions that pull data from outside the workbook
Private Function IsExternalData(formulaText As String) As Boolean
    IsExternalData = CallsFunction(formulaText, "WEBSERVICE") Or CallsFunction(formulaText, "SQL.REQUEST")
End Function

' True if formulaText calls the worksheet function functionName
Private Function CallsFunction(formulaText As String, functionName As String) As Boolean
    Dim upperText As String
    Dim pos As Long
    upperText = UCase$(formulaText)
    pos = InStr(1, upperText, functionName & "(")
    Do While pos > 0
        ' Not the end of a longer name (a "_xlfn." prefix is fine)
        If pos = 1 Then
            CallsFunction = True
        Else
            CallsFunction = Not (Mid$(upperText, pos - 1, 1) Like "[A-Z0-9_]")
        End If
        If CallsFunction Then Exit Function
        pos = InStr(pos + 1, upperText, functionName & "(")
    Loop
End Function

' formulaText with the contents of every "string literal" removed
Private Function WithoutStringLiterals(formulaText As String) As String
    Dim result As String
    Dim inString As Boolean
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(formulaText)
        ch = Mid$(formulaText, i, 1)
        If ch = """" Then
            ' Inside a string, "" is an escaped quote and keeps the string open
            inString = Not inString
            result = result & ch
        ElseIf Not inString Then
            result = result & ch
        End If
    Next i
    WithoutStringLiterals = result
End Function

' A hard-coded input: a number typed in, or a formula (formulaText, with its
' strings removed) that uses no cells, such as =1000*1.05. Text and dates are
' not inputs.
Private Function IsInput(cell As Range, formulaText As String) As Boolean
    If IsEmpty(cell.Value) Then Exit Function
    If VarType(cell.Value) = vbString Then Exit Function
    If IsDate(cell.Value) Then Exit Function

    If cell.HasFormula Then
        IsInput = Not ContainsCellReference(formulaText) Or IsOnlyNumbersAndOperators(formulaText)
    Else
        IsInput = True
    End If
End Function

Private Function IsOnlyNumbersAndOperators(ByVal formula As String) As Boolean
    ' Remove the equals sign if present
    If Left(formula, 1) = "=" Then formula = Mid(formula, 2)

    ' Only numbers, decimals, whitespace, and basic operators allowed
    Dim allowed As String
    allowed = "-+*/0123456789.,() " & vbTab & vbLf & vbCr & vbFormFeed & vbVerticalTab

    Dim i As Long
    For i = 1 To Len(formula)
        If InStr(1, allowed, Mid$(formula, i, 1), vbBinaryCompare) = 0 Then Exit Function
    Next i

    IsOnlyNumbersAndOperators = True
End Function

' ---------------------------------------------------------------------------
' Pattern helpers. These replace VBScript.RegExp, which does not exist on
' Excel for Mac. Each one reproduces the exact match semantics of the regex
' named in its comment, so results are identical on Windows and Mac.
' ---------------------------------------------------------------------------

Private Function IsAsciiLetter(ch As String) As Boolean
    IsAsciiLetter = (ch Like "[A-Za-z]")
End Function

Private Function IsAsciiDigit(ch As String) As Boolean
    IsAsciiDigit = (ch Like "[0-9]")
End Function

' Test for [$]?[A-Za-z]+[$]?[0-9]+|R[0-9]*C[0-9]*
Private Function ContainsCellReference(text As String) As Boolean
    Dim i As Long
    Dim j As Long

    For i = 1 To Len(text)
        ' A1 style: a letter, an optional $, then a digit
        If IsAsciiLetter(Mid$(text, i, 1)) Then
            j = i + 1
            If Mid$(text, j, 1) = "$" Then j = j + 1
            If IsAsciiDigit(Mid$(text, j, 1)) Then
                ContainsCellReference = True
                Exit Function
            End If
        End If

        ' R1C1 style: R, any digits, then C
        If Mid$(text, i, 1) = "R" Then
            j = i + 1
            Do While IsAsciiDigit(Mid$(text, j, 1))
                j = j + 1
            Loop
            If Mid$(text, j, 1) = "C" Then
                ContainsCellReference = True
                Exit Function
            End If
        End If
    Next i
End Function
