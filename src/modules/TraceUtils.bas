Option Explicit

' Border colour of a rounded box that has none (see CornerPicture)
Public Const NO_BORDER As Long = -1

' Corner pictures made this session, keyed by size, colours and corner
Private CornerPictures As Collection

' Notes after the address in a trace row (see GetReferencesRecursive):
' "(from range X)" for a cell listed under a range it belongs to, "(ref N)"
' for the Nth reference in the parent's formula, and "(again)" for a cell
' whose own precedents are listed under an earlier row
Private Const ROW_FROM_RANGE As String = " (from range "
Private Const ROW_REF As String = " (ref "
Private Const ROW_AGAIN As String = " (again)"

' history: the trace windows to return to with Shift+Enter; omitted from
' the ribbon, which starts a fresh one.
Public Sub ShowTracePrecedents(Optional history As Collection)
    Dim root As Range
    If Not GetTraceRoot(root) Then Exit Sub

    ' Building the trace selects cells and draws arrows for every node; don't
    ' repaint the sheet for each step. Excel turns this back on when the macro ends.
    Application.ScreenUpdating = False

    Dim frm As New frmPrecedents
    frm.SetRoot root
    FillTraceList frm.lstPrecedents, root, True
    frm.ConvertToTreeView
    frm.SetHistory history
    Application.ScreenUpdating = True
    frm.ShowAtDefaultPosition
End Sub

' history: the trace windows to return to with Shift+Enter; omitted from
' the ribbon, which starts a fresh one.
Public Sub ShowTraceDependents(Optional history As Collection)
    Dim root As Range
    If Not GetTraceRoot(root) Then Exit Sub

    Application.ScreenUpdating = False

    Dim frm As New frmDependents
    frm.SetRoot root
    FillTraceList frm.lstDependents, root, False
    frm.ConvertToTreeView
    frm.SetHistory history
    Application.ScreenUpdating = True
    frm.ShowAtDefaultPosition
End Sub

' The cell to trace: the first cell of the selection
Private Function GetTraceRoot(root As Range) As Boolean
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell or range."
        Exit Function
    End If
    Set root = Selection.Cells(1, 1)
    GetTraceRoot = True
End Function

' Fill the trace window's designer list box with the traced cell and then one
' row per traced cell: the indented "  L1: [Book]Sheet!$A$1" line, its value
' and its formula. The window turns these rows into its tree.
Private Sub FillTraceList(lst As MSForms.ListBox, root As Range, blnPrecedents As Boolean)
    lst.Clear
    lst.ColumnCount = 3
    AddTraceRow lst, root.Worksheet.Name & "!" & root.address, root

    Dim processedCells As New Collection
    Dim maxDepth As Integer
    maxDepth = Val(GetSetting("Breakdown", "FormulaTracing", "MaxDepth", "10"))

    Dim traceLines() As String
    Dim colonPos As Long
    Dim traced As Range
    Dim i As Long
    traceLines = Split(GetReferencesRecursive(root, blnPrecedents, 1, maxDepth, processedCells), vbCr)
    For i = 0 To UBound(traceLines)
        colonPos = InStr(traceLines(i), ": ")
        If colonPos > 0 Then
            Set traced = Nothing
            On Error Resume Next
            Set traced = Range(TraceRowAddress(Mid$(traceLines(i), colonPos + 2)))
            On Error GoTo 0
            AddTraceRow lst, traceLines(i), traced
        End If
    Next i

    If lst.ListCount > 0 Then lst.ListIndex = 0
End Sub

Private Sub AddTraceRow(lst As MSForms.ListBox, rowText As String, cell As Range)
    lst.AddItem
    lst.List(lst.ListCount - 1, 0) = rowText
    If cell Is Nothing Then
        lst.List(lst.ListCount - 1, 1) = "#N/A"
        lst.List(lst.ListCount - 1, 2) = "#N/A"
    Else
        lst.List(lst.ListCount - 1, 1) = GetCellValueAsString(cell)
        ' A range such as F3:F9 has no single formula (.Formula is an array)
        If cell.Cells.Count = 1 Then
            lst.List(lst.ListCount - 1, 2) = cell.formula
        Else
            lst.List(lst.ListCount - 1, 2) = ""
        End If
    End If
End Sub

Public Function GetCellValueAsString(cell As Range) As String
    On Error Resume Next
    If IsError(cell.value) Then
        GetCellValueAsString = "#ERROR"
    ElseIf IsEmpty(cell.value) Then
        GetCellValueAsString = ""
    Else
        ' Show the value as the cell displays it; fall back to the raw value
        ' when the column is too narrow and Excel shows ####
        Dim shownText As String
        shownText = Trim(cell.text)
        If Len(Replace(shownText, "#", "")) = 0 Then
            GetCellValueAsString = CStr(cell.value)
        Else
            GetCellValueAsString = shownText
        End If
    End If
    On Error GoTo 0
End Function


Private Function GetReferencesRecursive(aCell As Range, blnPrecedents As Boolean, currentLevel As Integer, maxLevel As Integer, processedCells As Collection) As String
    If currentLevel > maxLevel Then Exit Function

    ' Each cell is traced once; this also stops circular references
    Dim cellKey As String
    cellKey = aCell.address(External:=True)
    If AlreadyTraced(processedCells, cellKey) Then Exit Function
    processedCells.Add cellKey, cellKey

    Dim originalSelection As Range
    Set originalSelection = Selection

    ' Turn the direct references into ranges, without duplicates
    Dim precedentsList As New Collection
    Dim precedentAddresses() As String
    Dim precedentCell As Range
    Dim isDuplicate As Boolean
    Dim i As Long
    Dim j As Long
    precedentAddresses = Split(GetDirectReferences(aCell, blnPrecedents), vbCr)
    For i = 0 To UBound(precedentAddresses)
        If Len(Trim(precedentAddresses(i))) > 0 Then
            Set precedentCell = Nothing
            On Error Resume Next
            Set precedentCell = Range(Trim(precedentAddresses(i)))
            On Error GoTo 0

            If Not precedentCell Is Nothing Then
                isDuplicate = False
                For j = 1 To precedentsList.Count
                    If precedentsList(j).address(External:=True) = precedentCell.address(External:=True) Then
                        isDuplicate = True
                        Exit For
                    End If
                Next j
                If Not isDuplicate Then precedentsList.Add precedentCell
            End If
        End If
    Next i

    originalSelection.Select

    ' The rows to list: for a formula's precedents, one per reference in the
    ' order they are written, so a cell used twice is listed twice; otherwise
    ' one per dependent. Each is Array(index into precedentsList, number of
    ' the reference in the formula or 0).
    Dim refRows As Collection
    If blnPrecedents And aCell.HasFormula Then
        Set refRows = ReferenceRows(precedentsList, aCell)
    Else
        Set refRows = New Collection
        For i = 1 To precedentsList.Count
            refRows.Add Array(i, 0)
        Next i
    End If

    ' List each row and recurse into its cell straight away, so its own
    ' precedents are indented under it. A cell listed before, in this formula
    ' or elsewhere in the tree, is marked "(again)" and not traced again: its
    ' precedents are under its first row. (Dependents listed before are left
    ' out instead.)
    Dim allResults As String
    Dim rowStart As String
    Dim note As String
    Dim refRow As Variant
    Dim listed() As Boolean
    Dim precedent As Range
    Dim cell As Range
    ReDim listed(0 To precedentsList.Count)
    rowStart = String(currentLevel * 2, " ") & "L" & currentLevel & ": "
    For Each refRow In refRows
        Set precedent = precedentsList(refRow(0))
        If refRow(1) > 0 Then note = ROW_REF & refRow(1) & ")" Else note = ""
        If listed(refRow(0)) Then
            allResults = allResults & rowStart & precedent.address(External:=True) & note & ROW_AGAIN & vbCr
        ElseIf precedent.Cells.Count > 1 Then
            ' A range: list it, then each of its cells one level deeper
            allResults = allResults & rowStart & precedent.address(External:=True) & note & vbCr
            For Each cell In precedent.Cells
                If Not AlreadyTraced(processedCells, cell.address(External:=True)) Then
                    allResults = allResults & String((currentLevel + 1) * 2, " ") & "L" & (currentLevel + 1) & ": " & cell.address(External:=True) & ROW_FROM_RANGE & precedent.address(External:=True) & ")" & vbCr
                    allResults = allResults & GetReferencesRecursive(cell, blnPrecedents, currentLevel + 2, maxLevel, processedCells)
                End If
            Next cell
        ElseIf Not AlreadyTraced(processedCells, precedent.address(External:=True)) Then
            allResults = allResults & rowStart & precedent.address(External:=True) & note & vbCr
            allResults = allResults & GetReferencesRecursive(precedent, blnPrecedents, currentLevel + 1, maxLevel, processedCells)
        ElseIf blnPrecedents Then
            allResults = allResults & rowStart & precedent.address(External:=True) & note & ROW_AGAIN & vbCr
        End If
        listed(refRow(0)) = True
    Next refRow

    GetReferencesRecursive = allResults
End Function

' True if the cell has already been traced, i.e. it is shown elsewhere in the tree
Private Function AlreadyTraced(processedCells As Collection, cellKey As String) As Boolean
    On Error Resume Next
    Dim probe As String
    probe = processedCells(cellKey)
    AlreadyTraced = (Err.Number = 0)
    On Error GoTo 0
End Function

' Helper function to safely get direct precedents without navigation issues
Private Function GetDirectReferences(aCell As Range, blnPrecedents As Boolean) As String
    
    ' Early exit for dependents if cell is unlikely to have dependents
    If Not blnPrecedents Then
        ' For dependents: only proceed if cell has a non-empty value
        If IsEmpty(aCell.Value) Then
            GetDirectReferences = ""
            Exit Function
        End If
    End If
    
    Dim originalSelection As Range
    Dim originalSheet As Worksheet
    Set originalSelection = Selection
    Set originalSheet = ActiveSheet
    
    ' Safely select the target cell (handle cross-sheet references)
    On Error Resume Next
    aCell.Parent.Activate  ' Activate the worksheet first
    Dim activateError As Long
    activateError = Err.Number
    
    If activateError <> 0 Then
        ' If we can't activate the sheet, return empty result
        On Error GoTo 0
        GetDirectReferences = ""
        Exit Function
    End If
    
    aCell.Select  ' Now select the cell
    Dim selectError As Long
    selectError = Err.Number
    
    If selectError <> 0 Then
        ' If we can't select the cell, return empty result
        On Error GoTo 0
        originalSheet.Activate
        originalSelection.Select
        GetDirectReferences = ""
        Exit Function
    End If
    On Error GoTo 0
    
    Dim i As Long
    Dim results As String
    Dim safetyLimit As Long
    safetyLimit = Val(GetSetting("Breakdown", "FormulaTracing", "SafetyLimit", "100"))
    
    aCell.Parent.ClearArrows
    
    ' Check if cell has references before calling Show methods to avoid beep
    If blnPrecedents Then
        ' For precedents: only call if cell has a formula
        If aCell.HasFormula Then
            aCell.ShowPrecedents
        End If
    Else
        ' For dependents: only proceed if cell has a value that could be referenced
        ' Similar to HasFormula check for precedents
        If Not IsEmpty(aCell.Value) Then
            ' Cell has a value, so it might have dependents - call ShowDependents to create arrows
            aCell.ShowDependents
        End If
    End If

    ' Collect direct precedents/dependents with better error handling
    i = 0
    Do
        i = i + 1
        
        On Error Resume Next
        Dim navResult As Range
        Set navResult = aCell.NavigateArrow(blnPrecedents, 1, i)
        Dim navError As Long
        navError = Err.Number
        
        ' Check if navigation was successful and we didn't return to original cell
        If navError = 0 And Not navResult Is Nothing Then
            
            If navResult.address(External:=True) <> aCell.address(External:=True) Then
                results = results & navResult.address(External:=True) & vbCr
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
        On Error GoTo 0
        
        ' Safety check to prevent infinite loops
        If i > safetyLimit Then
            Exit Do
        End If
        
    Loop

    ' Collect precedents from other sheets/workbooks
    i = 1
    Do
        i = i + 1
        
        On Error Resume Next
        Set navResult = aCell.NavigateArrow(blnPrecedents, i, 1)
        Dim crossNavError As Long
        crossNavError = Err.Number
        
        If crossNavError = 0 And Not navResult Is Nothing Then
            
            If navResult.address(External:=True) <> aCell.address(External:=True) Then
                results = results & navResult.address(External:=True) & vbCr
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
        On Error GoTo 0
        
        ' Safety check
        If i > safetyLimit Then
            Exit Do
        End If
        
    Loop

    aCell.Parent.ClearArrows
    
    ' Restore original worksheet and selection
    originalSheet.Activate
    originalSelection.Select
    
    GetDirectReferences = results
End Function

' One entry per reference in formulaCell's formula, in the order written:
' Array(index into precedents, number of the reference among the formula's
' references). Each reference is matched to the precedent it names exactly,
' or else to the first one it overlaps. Precedents no reference names (reached
' through a defined name, say) follow, numbered 0.
Private Function ReferenceRows(precedents As Collection, formulaCell As Range) As Collection
    Dim addresses() As String
    Dim used() As Boolean
    Dim span As Variant
    Dim target As Range
    Dim targetAddress As String
    Dim spanNumber As Long
    Dim match As Long
    Dim k As Long

    Set ReferenceRows = New Collection
    If precedents.Count = 0 Then Exit Function
    ReDim addresses(1 To precedents.Count)
    ReDim used(1 To precedents.Count)
    For k = 1 To precedents.Count
        addresses(k) = precedents(k).address(External:=True)
    Next k

    For Each span In ReferenceSpans(formulaCell.formula)
        spanNumber = spanNumber + 1
        match = 0
        Set target = SpanRange(CStr(span(2)), CStr(span(3)), formulaCell.Worksheet)
        If Not target Is Nothing Then
            targetAddress = target.address(External:=True)
            For k = 1 To precedents.Count
                If addresses(k) = targetAddress Then
                    match = k
                    Exit For
                End If
            Next k
            If match = 0 Then
                For k = 1 To precedents.Count
                    If RangesOverlap(target, precedents(k)) Then
                        match = k
                        Exit For
                    End If
                Next k
            End If
        End If
        If match > 0 Then
            ReferenceRows.Add Array(match, spanNumber)
            used(match) = True
        End If
    Next span

    For k = 1 To precedents.Count
        If Not used(k) Then ReferenceRows.Add Array(k, 0)
    Next k
End Function

' ---------------------------------------------------------------------------
' Reference scanner, shared by the trace builder (one row per reference) and
' the trace windows (the highlighted chip in the formula card)
' ---------------------------------------------------------------------------

' Find the A1 reference in formulaText that points at relationCell: the first
' one naming exactly that cell or range, or else the first range containing
' it. Unqualified references point at formulaSheet. Text inside string
' literals is skipped. Returns the 1-based start and the length of the match.
Public Function FindReferenceSpan(formulaText As String, formulaSheet As Worksheet, relationCell As Range, _
                                   spanStart As Long, spanLength As Long) As Boolean
    Dim span As Variant
    Dim target As Range
    Dim relationAddress As String

    relationAddress = relationCell.address(External:=True)
    For Each span In ReferenceSpans(formulaText)
        Set target = SpanRange(CStr(span(2)), CStr(span(3)), formulaSheet)
        If Not target Is Nothing Then
            If target.address(External:=True) = relationAddress Then
                spanStart = span(0)
                spanLength = span(1)
                FindReferenceSpan = True
                Exit Function
            End If
        End If
    Next span

    For Each span In ReferenceSpans(formulaText)
        Set target = SpanRange(CStr(span(2)), CStr(span(3)), formulaSheet)
        If Not target Is Nothing Then
            If RangesOverlap(target, relationCell) Then
                spanStart = span(0)
                spanLength = span(1)
                FindReferenceSpan = True
                Exit Function
            End If
        End If
    Next span
End Function

' Every A1 reference in formulaText, left to right, as Array(start, length,
' sheet part, cell part). Text inside string literals is skipped.
Public Function ReferenceSpans(formulaText As String) As Collection
    Dim spans As New Collection
    Dim i As Long
    Dim refEnd As Long
    Dim sheetPart As String
    Dim cellPart As String
    Dim previousChar As String

    i = 1
    Do While i <= Len(formulaText)
        If i > 1 Then previousChar = Mid$(formulaText, i - 1, 1) Else previousChar = ""
        If Mid$(formulaText, i, 1) = """" Then
            i = SkipStringLiteral(formulaText, i)
        ElseIf IsNameChar(previousChar) Then
            ' Inside a longer name such as a function name
            i = i + 1
        Else
            refEnd = ParseReference(formulaText, i, sheetPart, cellPart)
            If refEnd = 0 Then
                i = i + 1
            Else
                spans.Add Array(i, refEnd - i, sheetPart, cellPart)
                i = refEnd
            End If
        End If
    Loop
    Set ReferenceSpans = spans
End Function

' Parse "[Sheet!]A1[:B2]" (sheet optionally quoted, possibly with [Book]) at
' position i. Returns the position just after it, or 0 if there is no A1
' reference here.
Private Function ParseReference(f As String, ByVal i As Long, sheetPart As String, cellPart As String) As Long
    Dim j As Long
    Dim k As Long
    sheetPart = ""
    cellPart = ""
    j = i

    ' Optional sheet prefix: 'Quoted name'! or Unquoted!
    If Mid$(f, j, 1) = "'" Then
        k = j + 1
        Do While k <= Len(f)
            If Mid$(f, k, 1) <> "'" Then
                k = k + 1
            ElseIf Mid$(f, k + 1, 1) = "'" Then
                k = k + 2  ' Escaped quote
            Else
                Exit Do
            End If
        Loop
        If Mid$(f, k + 1, 1) <> "!" Then Exit Function
        sheetPart = Replace(Mid$(f, j + 1, k - j - 1), "''", "'")
        j = k + 2
    Else
        k = j
        Do While IsSheetChar(Mid$(f, k, 1))
            k = k + 1
        Loop
        If k > j And Mid$(f, k, 1) = "!" Then
            sheetPart = Mid$(f, j, k - j)
            j = k + 1
        End If
    End If

    ' A cell, optionally followed by :cell
    Dim cellLength As Long
    Dim endPos As Long
    cellLength = CellRefLength(f, j)
    If cellLength = 0 Then Exit Function
    endPos = j + cellLength
    If Mid$(f, endPos, 1) = ":" Then
        cellLength = CellRefLength(f, endPos + 1)
        If cellLength > 0 Then endPos = endPos + 1 + cellLength
    End If

    ' Not the start of a longer name or a function call such as LOG10(
    If IsNameChar(Mid$(f, endPos, 1)) Or Mid$(f, endPos, 1) = "(" Then Exit Function

    cellPart = Mid$(f, j, endPos - j)
    ParseReference = endPos
End Function

' Length of an A1 cell reference ($A$1, A1, XFD1048576) at position j, or 0
Private Function CellRefLength(f As String, ByVal j As Long) As Long
    Dim k As Long
    Dim letterCount As Long
    Dim digitCount As Long
    k = j
    If Mid$(f, k, 1) = "$" Then k = k + 1
    Do While Mid$(f, k, 1) Like "[A-Za-z]"
        letterCount = letterCount + 1
        k = k + 1
    Loop
    If letterCount = 0 Or letterCount > 3 Then Exit Function
    If Mid$(f, k, 1) = "$" Then k = k + 1
    Do While Mid$(f, k, 1) Like "[0-9]"
        digitCount = digitCount + 1
        k = k + 1
    Loop
    If digitCount = 0 Or digitCount > 7 Then Exit Function
    CellRefLength = k - j
End Function

' The range a reference points to, from a formula on formulaSheet: sheetPart
' is "", "Sheet" or "[Book]Sheet". Nothing if it can't be resolved (a closed
' workbook, say).
Private Function SpanRange(sheetPart As String, cellPart As String, formulaSheet As Worksheet) As Range
    On Error GoTo Unresolved
    Dim ws As Worksheet
    Dim closePos As Long
    If Len(sheetPart) = 0 Then
        Set ws = formulaSheet
    ElseIf Left$(sheetPart, 1) = "[" Then
        closePos = InStr(sheetPart, "]")
        Set ws = Workbooks(Mid$(sheetPart, 2, closePos - 2)).Worksheets(Mid$(sheetPart, closePos + 1))
    Else
        Set ws = formulaSheet.Parent.Worksheets(sheetPart)
    End If
    Set SpanRange = ws.Range(cellPart)
Unresolved:
End Function

' True if two ranges on the same sheet share a cell
Private Function RangesOverlap(a As Range, b As Range) As Boolean
    On Error GoTo NoOverlap
    If a.Worksheet.Name <> b.Worksheet.Name Then Exit Function
    If a.Worksheet.Parent.Name <> b.Worksheet.Parent.Name Then Exit Function
    RangesOverlap = Not (Application.Intersect(a, b) Is Nothing)
NoOverlap:
End Function

' ---------------------------------------------------------------------------
' Trace rows, read back by the trace windows
' ---------------------------------------------------------------------------

' The address in a trace row's text (after its "L1: " prefix), without notes
Public Function TraceRowAddress(ByVal rowText As String) As String
    Dim note As Variant
    Dim pos As Long
    For Each note In Array(ROW_FROM_RANGE, ROW_REF, ROW_AGAIN)
        pos = InStr(rowText, note)
        If pos > 0 Then rowText = Left$(rowText, pos - 1)
    Next note
    TraceRowAddress = Trim$(rowText)
End Function

' The number of the reference in its parent's formula that a trace row
' stands for, or 0 if it has none
Public Function TraceRowRef(rowText As String) As Long
    Dim pos As Long
    pos = InStr(rowText, ROW_REF)
    If pos > 0 Then TraceRowRef = Val(Mid$(rowText, pos + Len(ROW_REF)))
End Function

' True for a row whose cell's own precedents are listed under an earlier row
Public Function TraceRowIsRepeat(rowText As String) As Boolean
    TraceRowIsRepeat = (InStr(rowText, ROW_AGAIN) > 0)
End Function

' Position just after the string literal that starts at i ("" is an escaped quote)
Private Function SkipStringLiteral(f As String, ByVal i As Long) As Long
    Dim k As Long
    k = i + 1
    Do While k <= Len(f)
        If Mid$(f, k, 1) <> """" Then
            k = k + 1
        ElseIf Mid$(f, k + 1, 1) = """" Then
            k = k + 2
        Else
            Exit Do
        End If
    Loop
    SkipStringLiteral = k + 1
End Function

' Characters that continue a name, number or reference
Private Function IsNameChar(ch As String) As Boolean
    IsNameChar = (ch Like "[A-Za-z0-9_.$!]")
End Function

' Characters of an unquoted sheet name, including a [Book] prefix
Private Function IsSheetChar(ch As String) As Boolean
    IsSheetChar = (ch Like "[A-Za-z0-9_.]") Or ch = "[" Or ch = "]"
End Function

' ---------------------------------------------------------------------------
' Corner pictures for the trace windows' rounded boxes. MSForms labels are
' square, so each corner of a rounded box is a small picture: the exact
' anti-aliased pixels of a quarter circle, blended from the background colour
' into the box's colour (and a one-pixel border ring, if it has one).
' ---------------------------------------------------------------------------

' The picture of one corner of a rounded box, r pixels square; corner 0 is
' top-left, 1 top-right, 2 bottom-left, 3 bottom-right. Made the first time it
' is needed: written as a BMP file to the temp folder and loaded. Nothing if
' that fails, and the trace window then draws its corners another way.
Public Function CornerPicture(ByVal r As Long, ByVal bgColor As Long, ByVal borderColor As Long, _
                              ByVal fillColor As Long, ByVal corner As Long) As Object
    Dim key As String
    Dim path As String
    Dim bytes() As Byte
    Dim fileNumber As Integer
    Dim fileOpen As Boolean

    If CornerPictures Is Nothing Then Set CornerPictures = New Collection
    key = "v1-" & r & "-" & Hex$(bgColor) & "-" & Hex$(borderColor) & "-" & Hex$(fillColor) & "-" & corner
    On Error Resume Next
    Set CornerPicture = CornerPictures(key)
    On Error GoTo 0
    If Not CornerPicture Is Nothing Then Exit Function

    On Error GoTo Failed
    path = CornerFolder() & "breakdown-corner-" & key & ".bmp"
    bytes = CornerBitmap(r, bgColor, borderColor, fillColor, corner)
    fileNumber = FreeFile
    Open path For Binary Access Write As #fileNumber
    fileOpen = True
    Put #fileNumber, 1, bytes
    Close #fileNumber
    fileOpen = False
    Set CornerPicture = LoadPicture(path)
    CornerPictures.Add CornerPicture, key
    Exit Function

Failed:
    If fileOpen Then Close #fileNumber
    Set CornerPicture = Nothing
End Function

' Excel's temp folder (inside its sandbox on Mac), else the add-in's folder
Private Function CornerFolder() As String
#If Mac Then
    CornerFolder = Environ("TMPDIR")
#Else
    CornerFolder = Environ("TEMP")
#End If
    If Len(CornerFolder) = 0 Then CornerFolder = ThisWorkbook.Path
    If Right$(CornerFolder, 1) <> Application.PathSeparator Then CornerFolder = CornerFolder & Application.PathSeparator
End Function

' A 24-bit BMP of an r x r corner square. The corner's circle has radius r and
' is centred on the square's inner corner; the border ring, if any, is the
' pixel between radius r - 1 and r.
Private Function CornerBitmap(ByVal r As Long, ByVal bgColor As Long, ByVal borderColor As Long, _
                              ByVal fillColor As Long, ByVal corner As Long) As Byte()
    Dim bytes() As Byte
    Dim stride As Long
    Dim pos As Long
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim cx As Double
    Dim cy As Double
    Dim outer As Double
    Dim inner As Double

    ' Header: 14-byte file header and 40-byte info header; rows run bottom-up
    ' and are padded to a multiple of 4 bytes
    stride = ((3 * r + 3) \ 4) * 4
    ReDim bytes(0 To 54 + stride * r - 1)
    bytes(0) = 66                        ' "B"
    bytes(1) = 77                        ' "M"
    PutLong bytes, 2, 54 + stride * r    ' File size
    PutLong bytes, 10, 54                ' Offset of the pixels
    PutLong bytes, 14, 40                ' Info header size
    PutLong bytes, 18, r                 ' Width
    PutLong bytes, 22, r                 ' Height
    bytes(26) = 1                        ' Planes
    bytes(28) = 24                       ' Bits per pixel
    PutLong bytes, 34, stride * r        ' Pixel data size
    PutLong bytes, 38, 2835              ' 72 dpi
    PutLong bytes, 42, 2835

    If corner Mod 2 = 0 Then cx = r Else cx = 0
    If corner < 2 Then cy = r Else cy = 0
    For j = 0 To r - 1
        pos = 54 + (r - 1 - j) * stride
        For i = 0 To r - 1
            outer = DiscCoverage(i, j, cx, cy, r)
            If borderColor = NO_BORDER Then inner = outer Else inner = DiscCoverage(i, j, cx, cy, r - 1)
            ' Blue, green, red
            For c = 2 To 0 Step -1
                bytes(pos) = BlendChannel(c, bgColor, borderColor, fillColor, outer, inner)
                pos = pos + 1
            Next c
        Next i
    Next j
    CornerBitmap = bytes
End Function

' Share of pixel (i, j) inside the circle of radius r around (cx, cy), in
' pixels: exact for pixels wholly inside or outside, else from 8 x 8 samples
Private Function DiscCoverage(ByVal i As Long, ByVal j As Long, ByVal cx As Double, ByVal cy As Double, _
                              ByVal r As Double) As Double
    Dim d As Double
    Dim sx As Double
    Dim sy As Double
    Dim a As Long
    Dim b As Long
    Dim hits As Long

    If r <= 0 Then Exit Function
    d = Sqr((i + 0.5 - cx) ^ 2 + (j + 0.5 - cy) ^ 2)
    If d <= r - 0.71 Then
        DiscCoverage = 1
        Exit Function
    End If
    If d >= r + 0.71 Then Exit Function

    For a = 0 To 7
        For b = 0 To 7
            sx = i + (a + 0.5) / 8 - cx
            sy = j + (b + 0.5) / 8 - cy
            If sx * sx + sy * sy <= r * r Then hits = hits + 1
        Next b
    Next a
    DiscCoverage = hits / 64
End Function

' One colour channel (0 red, 1 green, 2 blue) of a corner pixel: background
' outside the circle, border in the ring, fill inside
Private Function BlendChannel(ByVal channel As Long, ByVal bgColor As Long, ByVal borderColor As Long, _
                              ByVal fillColor As Long, ByVal outer As Double, ByVal inner As Double) As Byte
    Dim v As Double
    v = ColorChannel(bgColor, channel) * (1 - outer) _
        + ColorChannel(borderColor, channel) * (outer - inner) _
        + ColorChannel(fillColor, channel) * inner
    v = Int(v + 0.5)
    If v < 0 Then v = 0
    If v > 255 Then v = 255
    BlendChannel = v
End Function

' One channel of an RGB() colour (&HBBGGRR); NO_BORDER counts as white, and is
' only ever weighted zero
Private Function ColorChannel(ByVal color As Long, ByVal channel As Long) As Long
    If color < 0 Then
        ColorChannel = 255
    ElseIf channel = 0 Then
        ColorChannel = color And &HFF
    ElseIf channel = 1 Then
        ColorChannel = (color \ &H100) And &HFF
    Else
        ColorChannel = (color \ &H10000) And &HFF
    End If
End Function

' Write a Long into four bytes, least significant first
Private Sub PutLong(bytes() As Byte, ByVal pos As Long, ByVal value As Long)
    bytes(pos) = value And &HFF
    bytes(pos + 1) = (value \ &H100) And &HFF
    bytes(pos + 2) = (value \ &H10000) And &HFF
    bytes(pos + 3) = (value \ &H1000000) And &HFF
End Sub
