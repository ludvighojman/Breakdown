Option Explicit

Private Const NAME_PREFIX As String = "AutoColor_"

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
    For Each cell In usedCells
        If IsErrorValue(cell) Then
            cell.Font.Color = colorExternal
        ElseIf cell.HasFormula Then
            ' Text inside "quotes" is not part of any reference
            formulaText = WithoutStringLiterals(cell.Formula)
            If IsWorkbookLink(formulaText) Or IsExternalData(formulaText) Then
                cell.Font.Color = colorExternal
            ElseIf InStr(1, formulaText, "!") > 0 Then
                cell.Font.Color = colorWorksheetLink
            ElseIf IsInput(cell, formulaText) Then
                cell.Font.Color = colorInput
            Else
                cell.Font.Color = colorFormula
            End If
        ElseIf cell.Hyperlinks.Count > 0 Then
            ' Hyperlinks keep their own formatting
        ElseIf IsInput(cell, "") Then
            cell.Font.Color = colorInput
        End If
    Next cell

    Application.ScreenUpdating = True
End Sub

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
