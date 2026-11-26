Option Explicit

' Border colour of a rounded box that has none (see CornerPicture)
Public Const NO_BORDER As Long = -1

' Default of Settings > Formula Tracing > Second-level precedent tracing
' threshold (see SecondLevelThreshold)
Public Const DEFAULT_SECOND_LEVEL_THRESHOLD As Long = 50

' Corner pictures made this session, keyed by size, colours and corner
Private CornerPictures As Collection

' Notes after the address in a trace row (see ListReferences):
' "(from range X)" for a cell listed under a range it belongs to, "(ref N)"
' for the Nth reference in the parent's formula, and "(again)" for a cell
' whose own precedents are listed under an earlier row
Private Const ROW_FROM_RANGE As String = " (from range "
Private Const ROW_REF As String = " (ref "
Private Const ROW_AGAIN As String = " (again)"

' The budget of one trace, so that a cell with a huge dependency network
' can't run Excel out of memory or keep it busy indefinitely: at most
' MAX_TRACE_ROWS rows, and no cell traced after MAX_TRACE_SECONDS. Past the
' budget the traced cell's own references are still listed, untraced, so the
' first level stays whole (see ListReferences).
Private Const MAX_TRACE_ROWS As Long = 10000
Private Const MAX_TRACE_SECONDS As Single = 30
' A range of more cells than this shows no value in its row: reading it
' would copy the value of every cell
Private Const MAX_RANGE_VALUE_CELLS As Long = 10000
' Rows between progress updates in the status bar during a long trace
Private Const PROGRESS_ROWS As Long = 250

' The trace being built (see BuildTrace): its rows in order and the cell or
' range each stands for, the most arrows followed from one cell, and where
' the trace started
Private TraceRows() As String
Private TraceCells() As Range
Private TraceRowCount As Long
Private TraceSafetyLimit As Long
Private HomeSheet As Worksheet
Private HomeSelection As Range
Private HomeNames As Collection         ' Names defined in the workbook the trace started in
Private ClearedSheets As Collection     ' Sheets whose tracer arrows the trace has cleared
Private EventsBefore As Boolean         ' Application.EnableEvents when the trace started
Private TraceStart As Single            ' Timer when the trace started
Private TraceStopped As String          ' "rows" or "time" once the budget is spent, else ""
Private TraceNoun As String             ' "precedents" or "dependents", for the status bar
Private ShowedProgress As Boolean       ' The status bar shows the trace's progress
Private StatusBefore As Variant         ' ... and showed this before it

' Open a trace window on the selected cell. history: the trace windows to
' return to with Shift+Enter; omitted from the ribbon, which starts a fresh
' one. False if no window opened: nothing selected, or the user cancelled at
' the warning for a cell with many direct precedents.
Public Function ShowTracePrecedents(Optional history As Collection) As Boolean
    Dim root As Range
    Dim listRows As Variant
    Dim notice As Variant
    If Not GetTraceRoot(root) Then Exit Function
    If Not BuildTrace(root, True, listRows, notice) Then Exit Function

    ' Excel turns screen updating back on when the macro ends
    Application.ScreenUpdating = False
    Dim frm As New frmPrecedents
    frm.SetRoot root
    frm.SetNotice notice(0), notice(1), notice(2)
    frm.ConvertToTreeView listRows
    frm.SetHistory history
    Application.ScreenUpdating = True
    frm.ShowAtDefaultPosition
    ShowTracePrecedents = True
End Function

' As ShowTracePrecedents, for the cells that refer to the selected cell
Public Function ShowTraceDependents(Optional history As Collection) As Boolean
    Dim root As Range
    Dim listRows As Variant
    Dim notice As Variant
    If Not GetTraceRoot(root) Then Exit Function
    If Not BuildTrace(root, False, listRows, notice) Then Exit Function

    Application.ScreenUpdating = False
    Dim frm As New frmDependents
    frm.SetRoot root
    frm.SetNotice notice(0), notice(1), notice(2)
    frm.ConvertToTreeView listRows
    frm.SetHistory history
    Application.ScreenUpdating = True
    frm.ShowAtDefaultPosition
    ShowTraceDependents = True
End Function

' The most direct precedents a cell can have for its trace to go past the
' first level (Settings > Formula Tracing). The default when nothing valid is
' saved, so existing users get it without setting anything.
Public Function SecondLevelThreshold() As Long
    Dim saved As String
    saved = Trim$(GetSetting("Breakdown", "FormulaTracing", "SecondLevelThreshold", CStr(DEFAULT_SECOND_LEVEL_THRESHOLD)))
    If IsPositiveWholeNumber(saved) Then
        SecondLevelThreshold = CLng(saved)
    Else
        SecondLevelThreshold = DEFAULT_SECOND_LEVEL_THRESHOLD
    End If
End Function

' True for digits only, from 1 to 999,999,999 (so it fits a Long)
Public Function IsPositiveWholeNumber(ByVal numberText As String) As Boolean
    If Len(numberText) = 0 Or Len(numberText) > 9 Then Exit Function
    If numberText Like "*[!0-9]*" Then Exit Function
    IsPositiveWholeNumber = (CLng(numberText) >= 1)
End Function

' The cell to trace: the first cell of the selection
Private Function GetTraceRoot(root As Range) As Boolean
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell or range."
        Exit Function
    End If
    Set root = Selection.Cells(1, 1)
    GetTraceRoot = True
End Function

' Trace root's precedents (blnPrecedents) or dependents into the rows of a
' trace window: listRows(0 To n, 0 To 2) holds the traced cell and then one
' row per traced cell, each the indented "  L1: [Book]Sheet!$A$1" line, its
' value and its formula. The window turns these rows into its tree.
'
' Before anything is traced past the cell itself, its direct precedents are
' counted. More than the threshold in Settings asks the user first, and if
' they go ahead only the first level is traced: its cells' own precedents
' are never looked up. notice: Array(caption, explanation, incomplete) for
' the window (see TraceNotice). False if the user cancels: nothing more is
' traced and no window should open.
Private Function BuildTrace(root As Range, blnPrecedents As Boolean, listRows As Variant, notice As Variant) As Boolean
    Dim processedCells As New Collection
    Dim rootReferences As String
    Dim maxDepth As Integer
    Dim threshold As Long
    Dim firstLevelOnly As Boolean
    Dim failNumber As Long
    Dim failSource As String
    Dim failDescription As String

    maxDepth = Val(GetSetting("Breakdown", "FormulaTracing", "MaxDepth", "10"))
    TraceSafetyLimit = Val(GetSetting("Breakdown", "FormulaTracing", "SafetyLimit", "100"))
    threshold = SecondLevelThreshold()

    ' Building the trace selects cells and draws arrows; don't repaint the
    ' sheet for each step. Finding references moves around the workbook, so
    ' do it without running the workbook's own event macros, then go back to
    ' the selection the trace started from, also if it fails.
    Application.ScreenUpdating = False
    BeginTrace blnPrecedents
    On Error GoTo Failed

    ' The traced cell's own references: found once, for the check and the trace
    rootReferences = GetDirectReferences(root, blnPrecedents)
    If blnPrecedents Then
        If CellCountExceeds(rootReferences, threshold) Then
            If Not ConfirmFirstLevelTrace(threshold) Then
                EndTrace
                Exit Function
            End If
            firstLevelOnly = True
            maxDepth = 1
        End If
    End If

    ' The time budget starts now, not before the user answered
    TraceStart = Timer
    ListReferences root, blnPrecedents, 1, maxDepth, processedCells, rootReferences
    On Error GoTo 0
    EndTrace

    listRows = TraceListRows(root)
    notice = TraceNotice(firstLevelOnly, threshold)
    BuildTrace = True
    Exit Function

Failed:
    failNumber = Err.Number
    failSource = Err.Source
    failDescription = Err.Description
    EndTrace
    Erase TraceRows
    Erase TraceCells
    Err.Raise failNumber, failSource, failDescription
End Function

' Start a trace from the active sheet and selection, with event macros off
Private Sub BeginTrace(blnPrecedents As Boolean)
    Set HomeSheet = ActiveSheet
    Set HomeSelection = Selection
    Set HomeNames = DefinedNames(HomeSheet.Parent)
    Set ClearedSheets = New Collection
    TraceRowCount = 0
    ReDim TraceRows(1 To 64)
    ReDim TraceCells(1 To 64)
    TraceStopped = ""
    TraceStart = Timer
    If blnPrecedents Then TraceNoun = "precedents" Else TraceNoun = "dependents"
    ShowedProgress = False
    EventsBefore = Application.EnableEvents
    Application.EnableEvents = False
End Sub

' Go back to where the trace started, run event macros again if they ran
' before, and give the status bar back
Private Sub EndTrace()
    On Error Resume Next
    HomeSheet.Activate
    HomeSelection.Select
    Application.EnableEvents = EventsBefore
    If ShowedProgress Then Application.StatusBar = StatusBefore
    ShowedProgress = False
    Set HomeSheet = Nothing
    Set HomeSelection = Nothing
    Set HomeNames = Nothing
    Set ClearedSheets = Nothing
End Sub

' True if the cells in references (addresses joined with vbCr, as
' GetDirectReferences gives them) number more than limit: a range counts as
' all its cells, an address listed twice once, and overlapping ranges in
' full. Counting stops as soon as it is past the limit, so a reference to a
' whole column costs no more than one to a single cell.
Private Function CellCountExceeds(references As String, ByVal limit As Long) As Boolean
    Dim addresses() As String
    Dim counted As New Collection
    Dim target As Range
    Dim targetAddress As String
    Dim total As Double
    Dim i As Long

    addresses = Split(references, vbCr)
    For i = 0 To UBound(addresses)
        If Len(Trim$(addresses(i))) > 0 Then
            Set target = Nothing
            On Error Resume Next
            Set target = Range(Trim$(addresses(i)))
            On Error GoTo 0
            If Not target Is Nothing Then
                targetAddress = target.address(External:=True)
                If Not HasKey(counted, targetAddress) Then
                    counted.Add True, targetAddress
                    total = total + target.CountLarge
                    If total > limit Then
                        CellCountExceeds = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

' Warn before tracing a cell with more direct precedents than the threshold.
' True if the user goes ahead, with first-level tracing only.
Private Function ConfirmFirstLevelTrace(ByVal threshold As Long) As Boolean
    Dim message As String
    message = "This cell contains more than " & Format$(threshold, "#,##0") & " direct precedents. " & _
             "Analysing this cell may result in reduced performance or application instability. " & _
             "Second-level tracing will be automatically disabled to reduce computational load." & _
             vbNewLine & vbNewLine & "Do you want to continue anyway?"
    Application.ScreenUpdating = True
    ConfirmFirstLevelTrace = (MsgBox(message, vbOKCancel + vbExclamation + vbDefaultButton2, _
                                     "Large number of precedents detected") = vbOK)
    If ConfirmFirstLevelTrace Then Application.ScreenUpdating = False
End Function

' What the trace window says about a limited trace: Array(caption,
' explanation, incomplete), with an empty caption if the trace went as far
' as the settings allow. incomplete: the budget stopped the trace.
Private Function TraceNotice(ByVal firstLevelOnly As Boolean, ByVal threshold As Long) As Variant
    Dim captionText As String
    Dim detail As String
    Dim limitText As String

    If firstLevelOnly Then
        captionText = "Second-level trace disabled"
        detail = "Second-level precedent tracing has been disabled because the number of direct " & _
                 "precedents exceeds your configured threshold (" & Format$(threshold, "#,##0") & "). " & _
                 "This helps maintain application performance and stability. " & _
                 "You can adjust the threshold in Settings."
    End If
    If Len(TraceStopped) > 0 Then
        If Len(captionText) = 0 Then captionText = "Trace stopped early"
        If Len(detail) > 0 Then detail = detail & " "
        If TraceStopped = "rows" Then
            limitText = "after " & Format$(MAX_TRACE_ROWS, "#,##0") & " cells"
        Else
            limitText = "after " & MAX_TRACE_SECONDS & " seconds"
        End If
        detail = detail & "The trace stopped " & limitText & " to keep Excel responsive, so the list " & _
                 "is incomplete. Select a cell and press Enter to trace from there."
    End If
    TraceNotice = Array(captionText, detail, Len(TraceStopped) > 0)
End Function

' True once the trace has spent its budget (see MAX_TRACE_ROWS), recording
' which part ran out
Private Function OverBudget() As Boolean
    Dim elapsed As Single
    If Len(TraceStopped) = 0 Then
        If TraceRowCount >= MAX_TRACE_ROWS Then
            TraceStopped = "rows"
        Else
            elapsed = Timer - TraceStart
            If elapsed < 0 Then elapsed = elapsed + 86400   ' Past midnight
            If elapsed > MAX_TRACE_SECONDS Then TraceStopped = "time"
        End If
    End If
    OverBudget = (Len(TraceStopped) > 0)
End Function

' Add a row to the trace being built, for a cell or range
Private Sub AddTraceLine(rowText As String, cell As Range)
    TraceRowCount = TraceRowCount + 1
    If TraceRowCount > UBound(TraceRows) Then
        ReDim Preserve TraceRows(1 To 2 * UBound(TraceRows))
        ReDim Preserve TraceCells(1 To UBound(TraceRows))
    End If
    TraceRows(TraceRowCount) = rowText
    Set TraceCells(TraceRowCount) = cell
    If TraceRowCount Mod PROGRESS_ROWS = 0 Then ShowTraceProgress
End Sub

' A long trace shows how far it has got in the status bar
Private Sub ShowTraceProgress()
    On Error Resume Next
    If Not ShowedProgress Then StatusBefore = Application.StatusBar
    ShowedProgress = True
    Application.StatusBar = "Breakdown: tracing " & TraceNoun & "... " & Format$(TraceRowCount, "#,##0") & " cells"
End Sub

' The rows for the trace window: the traced cell, then the trace rows. The
' trace's own rows are let go.
Private Function TraceListRows(root As Range) As Variant
    Dim listRows() As Variant
    Dim listRow As Variant
    Dim i As Long
    Dim j As Long

    ReDim listRows(0 To TraceRowCount, 0 To 2)
    listRow = TraceRow(root.Worksheet.Name & "!" & root.address, root)
    For j = 0 To 2
        listRows(0, j) = listRow(j)
    Next j
    For i = 1 To TraceRowCount
        listRow = TraceRow(TraceRows(i), TraceCells(i))
        For j = 0 To 2
            listRows(i, j) = listRow(j)
        Next j
    Next i
    Erase TraceRows
    Erase TraceCells
    TraceListRows = listRows
End Function

' One row of the trace window: Array(row text, value, formula)
Private Function TraceRow(rowText As String, cell As Range) As Variant
    If cell Is Nothing Then
        TraceRow = Array(rowText, "#N/A", "#N/A")
    ElseIf cell.CountLarge = 1 Then
        TraceRow = Array(rowText, GetCellValueAsString(cell), CStr(cell.formula))
    ElseIf cell.CountLarge <= MAX_RANGE_VALUE_CELLS Then
        ' A range such as F3:F9 has no single formula (.Formula is an array)
        TraceRow = Array(rowText, GetCellValueAsString(cell), "")
    Else
        TraceRow = Array(rowText, "", "")
    End If
End Function

Public Function GetCellValueAsString(cell As Range) As String
    Dim cellValue As Variant
    On Error Resume Next
    cellValue = cell.value
    If Err.Number <> 0 Then
        GetCellValueAsString = "#ERROR"     ' The value can't be read
    ElseIf IsError(cellValue) Then
        GetCellValueAsString = "#ERROR"
    ElseIf IsEmpty(cellValue) Then
        GetCellValueAsString = ""
    Else
        ' Show the value as the cell displays it; fall back to the raw value
        ' when the column is too narrow and Excel shows ####
        Dim shownText As String
        shownText = Trim(cell.text)
        If Len(Replace(shownText, "#", "")) = 0 Then
            GetCellValueAsString = CStr(cellValue)
        Else
            GetCellValueAsString = shownText
        End If
    End If
    On Error GoTo 0
End Function


' Add the rows for aCell's references (or the cells referring to it) to the
' trace, each followed by the rows for its own, depth first. knownReferences:
' aCell's references if already found, as GetDirectReferences gives them.
' Once the budget is spent (see OverBudget) nothing more is traced; only the
' traced cell's own references (level 1) are still listed, so the first
' level stays whole, and a range among them stops listing its cells.
Private Sub ListReferences(aCell As Range, blnPrecedents As Boolean, currentLevel As Integer, maxLevel As Integer, processedCells As Collection, _
                           Optional knownReferences As Variant)
    If currentLevel > maxLevel Then Exit Sub
    If OverBudget() Then Exit Sub

    ' Each cell is traced once; this also stops circular references
    Dim cellKey As String
    cellKey = aCell.address(External:=True)
    If HasKey(processedCells, cellKey) Then Exit Sub
    processedCells.Add cellKey, cellKey

    ' Turn the direct references into ranges, without duplicates
    Dim precedentsList As New Collection
    Dim precedentAddresses() As String
    Dim listedAddresses() As String
    Dim precedentCell As Range
    Dim precedentAddress As String
    Dim isDuplicate As Boolean
    Dim i As Long
    Dim j As Long
    If IsMissing(knownReferences) Then
        precedentAddresses = Split(GetDirectReferences(aCell, blnPrecedents), vbCr)
    Else
        precedentAddresses = Split(CStr(knownReferences), vbCr)
    End If
    ReDim listedAddresses(0 To UBound(precedentAddresses) + 1)
    For i = 0 To UBound(precedentAddresses)
        If Len(Trim(precedentAddresses(i))) > 0 Then
            Set precedentCell = Nothing
            On Error Resume Next
            Set precedentCell = Range(Trim(precedentAddresses(i)))
            On Error GoTo 0

            If Not precedentCell Is Nothing Then
                precedentAddress = precedentCell.address(External:=True)
                isDuplicate = False
                For j = 1 To precedentsList.Count
                    If listedAddresses(j) = precedentAddress Then
                        isDuplicate = True
                        Exit For
                    End If
                Next j
                If Not isDuplicate Then
                    precedentsList.Add precedentCell
                    listedAddresses(precedentsList.Count) = precedentAddress
                End If
            End If
        End If
    Next i

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

    ' List each row and trace its cell straight away, so its own references
    ' are indented under it. A cell listed before, in this formula or
    ' elsewhere in the tree, is marked "(again)" and not traced again: its
    ' references are under its first row. (Dependents listed before are left
    ' out instead.)
    Dim rowStart As String
    Dim note As String
    Dim refRow As Variant
    Dim listed() As Boolean
    Dim precedent As Range
    Dim cell As Range
    Dim cellAddress As String
    ReDim listed(0 To precedentsList.Count)
    rowStart = String(currentLevel * 2, " ") & "L" & currentLevel & ": "
    For Each refRow In refRows
        If currentLevel > 1 Then
            If OverBudget() Then Exit Sub
        End If
        Set precedent = precedentsList(refRow(0))
        precedentAddress = listedAddresses(refRow(0))
        If refRow(1) > 0 Then note = ROW_REF & refRow(1) & ")" Else note = ""
        If listed(refRow(0)) Then
            AddTraceLine rowStart & precedentAddress & note & ROW_AGAIN, precedent
        ElseIf precedent.CountLarge > 1 Then
            ' A range: list it, then each of its cells one level deeper.
            ' (CountLarge: a whole sheet has more cells than a Long holds.)
            AddTraceLine rowStart & precedentAddress & note, precedent
            For Each cell In precedent.Cells
                If OverBudget() Then Exit For
                cellAddress = cell.address(External:=True)
                If Not HasKey(processedCells, cellAddress) Then
                    AddTraceLine String((currentLevel + 1) * 2, " ") & "L" & (currentLevel + 1) & ": " & cellAddress & ROW_FROM_RANGE & precedentAddress & ")", cell
                    ListReferences cell, blnPrecedents, currentLevel + 2, maxLevel, processedCells
                End If
            Next cell
        ElseIf Not HasKey(processedCells, precedentAddress) Then
            AddTraceLine rowStart & precedentAddress & note, precedent
            ListReferences precedent, blnPrecedents, currentLevel + 1, maxLevel, processedCells
        ElseIf blnPrecedents Then
            AddTraceLine rowStart & precedentAddress & note & ROW_AGAIN, precedent
        End If
        listed(refRow(0)) = True
    Next refRow
End Sub

' True if the collection has an item with this key (for the traced cells:
' the cell is shown elsewhere in the tree)
Private Function HasKey(items As Collection, key As String) As Boolean
    On Error Resume Next
    Dim probe As Variant
    probe = items(key)
    HasKey = (Err.Number = 0)
    On Error GoTo 0
End Function

' The names defined in a workbook, keyed by name without any sheet prefix
Private Function DefinedNames(book As Workbook) As Collection
    Dim definedName As Variant
    Set DefinedNames = New Collection
    On Error Resume Next
    For Each definedName In book.Names
        DefinedNames.Add True, Mid$(definedName.Name, InStrRev(definedName.Name, "!") + 1)
    Next definedName
End Function

' The cells aCell refers to (blnPrecedents) or that refer to it, one address
' per line, found by following Excel's tracer arrows from the cell
Private Function GetDirectReferences(aCell As Range, blnPrecedents As Boolean) As String
    ' Only a formula has precedents (the arrows give a merged cell its own
    ' merged area), and only a cell with a value dependents. Most formulas'
    ' precedents can be read from the formula itself.
    Dim parsed As String
    If blnPrecedents Then
        If Not aCell.HasFormula And Not aCell.MergeCells Then Exit Function
        If FormulaPrecedents(aCell, parsed) Then
            GetDirectReferences = parsed
            Exit Function
        End If
    ElseIf IsEmpty(aCell.Value) Then
        Exit Function
    End If

    ' Select the cell, going there from the workbook the trace started in.
    ' The trace goes back to its starting selection when it ends.
    On Error Resume Next
    If ActiveWorkbook.Name <> HomeSheet.Parent.Name Then HomeSheet.Activate
    If ActiveSheet.Name <> aCell.Parent.Name Or ActiveWorkbook.Name <> aCell.Parent.Parent.Name Then aCell.Parent.Activate
    If Err.Number <> 0 Then Exit Function
    aCell.Select
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    Dim i As Long
    Dim results As String
    Dim navResult As Range
    Dim navError As Long
    Dim sheetKey As String

    ' Follow only this cell's arrows: clear any the sheet had when the trace
    ' first came to it, and this cell's once followed
    sheetKey = aCell.Parent.Parent.Name & "|" & aCell.Parent.Name
    If Not HasKey(ClearedSheets, sheetKey) Then
        aCell.Parent.ClearArrows
        ClearedSheets.Add True, sheetKey
    End If
    ' (Showing precedents of a cell without a formula would beep)
    If Not blnPrecedents Then
        aCell.ShowDependents
    ElseIf aCell.HasFormula Then
        aCell.ShowPrecedents
    End If

    ' Direct precedents/dependents on the same sheet
    i = 0
    Do
        i = i + 1

        On Error Resume Next
        Set navResult = aCell.NavigateArrow(blnPrecedents, 1, i)
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
        If i > TraceSafetyLimit Then Exit Do
    Loop

    ' Those on other sheets and in other workbooks
    i = 1
    Do
        i = i + 1

        On Error Resume Next
        Set navResult = aCell.NavigateArrow(blnPrecedents, i, 1)
        navError = Err.Number

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

        If i > TraceSafetyLimit Then Exit Do
    Loop

    aCell.Parent.ClearArrows
    GetDirectReferences = results
End Function

' A formula cell's precedents read from the formula itself, joined with vbCr
' like GetDirectReferences, when that is certain to give the cells Excel's
' tracer arrows give (following arrows takes milliseconds a cell). It is when
' the cell is not merged, and the formula only refers to cells with plain A1
' references to visible, unprotected sheets of the workbook the trace
' started in, none of them the cell itself or merged cells, and only calls
' built-in functions: no names, tables, spills, whole rows or columns, other
' workbooks, 3D references, intersections or LET variables. For such a
' formula the arrows give each cell or range it refers to, once. False
' otherwise.
Private Function FormulaPrecedents(aCell As Range, results As String) As Boolean
    Dim formulaText As String
    Dim spans As Collection
    Dim span As Variant
    Dim target As Range
    Dim targetAddress As String
    Dim cellAddress As String
    Dim found As New Collection
    Dim sameSheetCount As Long
    Dim otherSheetCount As Long

    If aCell.Worksheet.Parent.Name <> HomeSheet.Parent.Name Then Exit Function
    If Not ArrowsWork(aCell.Worksheet) Then Exit Function
    If aCell.HasArray Or aCell.MergeCells Then Exit Function
    formulaText = aCell.formula
    If Left$(formulaText, 1) <> "=" Then Exit Function
    Set spans = ReferenceSpans(formulaText)
    If Not OnlyBuiltInsAndOperators(formulaText, spans) Then Exit Function

    cellAddress = aCell.address(External:=True)
    For Each span In spans
        If Left$(span(2), 1) = "[" Then Exit Function
        Set target = SpanRange(CStr(span(2)), CStr(span(3)), aCell.Worksheet)
        If target Is Nothing Then Exit Function
        ' Off the cell's sheet, every reference is a link of one arrow
        If target.Worksheet.Name <> aCell.Worksheet.Name Then
            If Not ArrowsWork(target.Worksheet) Then Exit Function
            otherSheetCount = otherSheetCount + 1
        End If
        ' The arrow to a merged cell gives its whole merged area
        If IsNull(target.MergeCells) Then Exit Function
        If target.MergeCells Then Exit Function
        targetAddress = target.address(External:=True)
        If targetAddress = cellAddress Then Exit Function
        If Not HasKey(found, targetAddress) Then
            found.Add targetAddress, targetAddress
            If target.Worksheet.Name = aCell.Worksheet.Name Then sameSheetCount = sameSheetCount + 1
            results = results & targetAddress & vbCr
        End If
    Next span

    ' No more than GetDirectReferences follows
    If sameSheetCount >= TraceSafetyLimit Or otherSheetCount >= TraceSafetyLimit Then Exit Function
    FormulaPrecedents = True
End Function

' True for a sheet where Excel can show and follow tracer arrows
Private Function ArrowsWork(ws As Worksheet) As Boolean
    ArrowsWork = (ws.Visible = xlSheetVisible And Not ws.ProtectContents)
End Function

' True if formulaText, apart from string literals and the references in
' spans, holds only calls of built-in functions, TRUE and FALSE, numbers and
' operators. A name could stand for cells and a space could be an
' intersection, so either gives False.
Private Function OnlyBuiltInsAndOperators(formulaText As String, spans As Collection) As Boolean
    Dim i As Long
    Dim j As Long
    Dim spanIndex As Long
    Dim nextSpan As Long
    Dim ch As String
    Dim word As String

    spanIndex = 1
    i = 1
    Do While i <= Len(formulaText)
        If spanIndex <= spans.Count Then nextSpan = spans(spanIndex)(0) Else nextSpan = 0
        ch = Mid$(formulaText, i, 1)
        If i = nextSpan Then
            i = i + spans(spanIndex)(1)
            spanIndex = spanIndex + 1
        ElseIf ch = """" Then
            i = SkipStringLiteral(formulaText, i)
        ElseIf ch Like "[A-Za-z_]" Then
            ' A function name, or TRUE or FALSE
            j = i
            Do While Mid$(formulaText, j, 1) Like "[A-Za-z0-9_.]"
                j = j + 1
            Loop
            word = Mid$(formulaText, i, j - i)
            If Mid$(formulaText, j, 1) = "(" Then
                If HasKey(HomeNames, word) Then Exit Function   ' A LAMBDA name, say
            ElseIf UCase$(word) <> "TRUE" And UCase$(word) <> "FALSE" Then
                Exit Function
            End If
            i = j
        ElseIf InStr(1, "0123456789.+-*/^&=<>(),;%{}", ch) > 0 Then
            i = i + 1
        Else
            Exit Function
        End If
    Loop
    OnlyBuiltInsAndOperators = True
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
    Dim bytes() As Byte

    If CornerPictures Is Nothing Then Set CornerPictures = New Collection
    key = "v1-" & r & "-" & Hex$(bgColor) & "-" & Hex$(borderColor) & "-" & Hex$(fillColor) & "-" & corner
    On Error Resume Next
    Set CornerPicture = CornerPictures(key)
    On Error GoTo 0
    If Not CornerPicture Is Nothing Then Exit Function

    bytes = CornerBitmap(r, bgColor, borderColor, fillColor, corner)
    Set CornerPicture = BitmapPicture(bytes, "breakdown-corner-" & key & ".bmp")
    If Not CornerPicture Is Nothing Then CornerPictures.Add CornerPicture, key
End Function

' A bitmap loaded as a picture, through a file of that name in Excel's temp
' folder (inside its sandbox on Mac); Nothing if that fails
Public Function BitmapPicture(bytes() As Byte, ByVal fileName As String) As Object
    Dim folder As String
    Dim fileNumber As Integer
    Dim fileOpen As Boolean

    On Error GoTo Failed
#If Mac Then
    folder = Environ("TMPDIR")
#Else
    folder = Environ("TEMP")
#End If
    If Len(folder) = 0 Then folder = ThisWorkbook.Path
    If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator

    fileNumber = FreeFile
    Open folder & fileName For Binary Access Write As #fileNumber
    fileOpen = True
    Put #fileNumber, 1, bytes
    Close #fileNumber
    fileOpen = False
    Set BitmapPicture = LoadPicture(folder & fileName)
    Exit Function

Failed:
    If fileOpen Then Close #fileNumber
    Set BitmapPicture = Nothing
End Function

' A black 24-bit BMP, pixelsWide x pixelsHigh: a 14-byte file header and a
' 40-byte info header, then the rows bottom-up, each padded to a multiple of
' 4 bytes
Public Function NewBitmap(ByVal pixelsWide As Long, ByVal pixelsHigh As Long) As Byte()
    Dim bytes() As Byte
    Dim stride As Long

    stride = ((3 * pixelsWide + 3) \ 4) * 4
    ReDim bytes(0 To 54 + stride * pixelsHigh - 1)
    bytes(0) = 66                                ' "B"
    bytes(1) = 77                                ' "M"
    PutLong bytes, 2, 54 + stride * pixelsHigh   ' File size
    PutLong bytes, 10, 54                        ' Offset of the pixels
    PutLong bytes, 14, 40                        ' Info header size
    PutLong bytes, 18, pixelsWide
    PutLong bytes, 22, pixelsHigh
    bytes(26) = 1                                ' Planes
    bytes(28) = 24                               ' Bits per pixel
    PutLong bytes, 34, stride * pixelsHigh       ' Pixel data size
    PutLong bytes, 38, 2835                      ' 72 dpi
    PutLong bytes, 42, 2835
    NewBitmap = bytes
End Function

' Set pixel (x, y) of a NewBitmap, counted from the top-left, to an RGB() colour
Public Sub SetPixel(bytes() As Byte, ByVal x As Long, ByVal y As Long, ByVal rgbColor As Long)
    Dim pixelsWide As Long
    Dim pixelsHigh As Long
    Dim pos As Long

    pixelsWide = bytes(18) + bytes(19) * &H100&
    pixelsHigh = bytes(22) + bytes(23) * &H100&
    pos = 54 + (pixelsHigh - 1 - y) * (((3 * pixelsWide + 3) \ 4) * 4) + 3 * x
    bytes(pos) = (rgbColor \ &H10000) And &HFF   ' Blue
    bytes(pos + 1) = (rgbColor \ &H100) And &HFF ' Green
    bytes(pos + 2) = rgbColor And &HFF           ' Red
End Sub

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

    bytes = NewBitmap(r, r)
    stride = ((3 * r + 3) \ 4) * 4
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
