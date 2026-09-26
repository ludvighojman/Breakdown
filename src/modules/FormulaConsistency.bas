Attribute VB_Name = "FormulaConsistency"
Option Explicit

Private Const PATTERN_HORIZONTAL As Long = xlHorizontal     ' Horizontal lines
Private Const COLOR_CONSISTENT As Long = 14348800           ' Green
Private Const COLOR_INCONSISTENT As Long = 255              ' Red

' Very hidden sheet in the add-in that keeps the original fill of every
' formula cell on each marked sheet, one row per cell: key ("Book|Sheet"),
' address, pattern, color, pattern color. Each marked sheet also has a row
' with its key and no address, so a sheet counts as marked while its key is
' in column A.
Private Const STASH_SHEET As String = "OriginalFormat"

' Marks the active sheet's formulas as consistent or inconsistent with their
' right-hand neighbour; on a sheet that is already marked, removes the marks
Public Sub CheckHorizontalConsistency()
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If HasMarks(ws) Then
        RemoveMarks ws
        MsgBox "Formula consistency formatting has been removed.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    StoreOriginalFormatting ws

    Dim usedRng As Range
    Set usedRng = ws.UsedRange

    ' First pass - check and store horizontal consistency information
    Dim horizontalConsistentFormulasR1C1 As New Collection
    Dim R As Long, c As Long

    ' Find horizontally consistent formula patterns
    For R = usedRng.Row To usedRng.Row + usedRng.Rows.Count - 1
        For c = usedRng.Column To usedRng.Column + usedRng.Columns.Count - 2
            If ws.Cells(R, c).HasFormula And ws.Cells(R, c + 1).HasFormula Then
                If ws.Cells(R, c).FormulaR1C1 = ws.Cells(R, c + 1).FormulaR1C1 Then
                    AddToCollection horizontalConsistentFormulasR1C1, ws.Cells(R, c).FormulaR1C1
                End If
            End If
        Next c
    Next R

    ' Check each cell for consistency with neighbors
    Dim cell As Range
    Dim isHorizConsistent As Boolean
    Dim checkHoriz As Boolean
    For Each cell In usedRng
        If cell.HasFormula Then
            checkHoriz = False

            ' Check horizontal consistency
            If cell.Column < usedRng.Columns.Count + usedRng.Column - 1 Then
                If cell.Offset(0, 1).HasFormula Then
                    checkHoriz = True
                    isHorizConsistent = (cell.FormulaR1C1 = cell.Offset(0, 1).FormulaR1C1)
                Else
                    ' Check if this is the last cell in a consistent sequence
                    isHorizConsistent = IsInCollection(horizontalConsistentFormulasR1C1, cell.FormulaR1C1)
                    checkHoriz = isHorizConsistent
                End If
            Else
                ' Last column - check if part of consistent sequence
                isHorizConsistent = IsInCollection(horizontalConsistentFormulasR1C1, cell.FormulaR1C1)
                checkHoriz = isHorizConsistent
            End If

            ' Apply formatting based on consistency
            With cell.Interior
                .Pattern = xlNone
                If checkHoriz Then
                    .Pattern = PATTERN_HORIZONTAL
                    .PatternColor = IIf(isHorizConsistent, COLOR_CONSISTENT, COLOR_INCONSISTENT)
                End If
            End With
        End If
    Next cell

    Application.ScreenUpdating = True

    Dim msg As String
    msg = "Horizontal Formula Consistency Check Complete:" & vbNewLine & vbNewLine
    msg = msg & "Cells marked with:" & vbNewLine
    msg = msg & "- Green horizontal lines: Consistent with right neighbor" & vbNewLine
    msg = msg & "- Red horizontal lines: Inconsistent with right neighbor"
    MsgBox msg, vbInformation
End Sub

Private Sub AddToCollection(col As Collection, item As String)
    On Error Resume Next
    col.Add item, item  ' Using item as key prevents duplicates
    On Error GoTo 0
End Sub

Private Function IsInCollection(col As Collection, item As String) As Boolean
    Dim v As Variant
    For Each v In col
        If v = item Then
            IsInCollection = True
            Exit Function
        End If
    Next v
End Function

' ---------------------------------------------------------------------------
' Stash of original fills
' ---------------------------------------------------------------------------

Private Function StashKey(ws As Worksheet) As String
    StashKey = ws.Parent.Name & "|" & ws.Name
End Function

' The stash sheet, or Nothing if no sheet is marked
Private Function StashSheet() As Worksheet
    On Error Resume Next
    Set StashSheet = ThisWorkbook.Worksheets(STASH_SHEET)
    On Error GoTo 0
End Function

' All stash rows as a 2-D array (row, 1 To 5), or Empty if there are none
Private Function StashRows(stash As Worksheet) As Variant
    Dim lastRow As Long
    lastRow = stash.Cells(stash.Rows.Count, 1).End(xlUp).Row
    If lastRow = 1 And IsEmpty(stash.Cells(1, 1).Value) Then Exit Function
    StashRows = stash.Range(stash.Cells(1, 1), stash.Cells(lastRow, 5)).Value
End Function

Private Function HasMarks(ws As Worksheet) As Boolean
    Dim stash As Worksheet
    Dim stashData As Variant
    Dim key As String
    Dim i As Long

    Set stash = StashSheet()
    If stash Is Nothing Then Exit Function
    stashData = StashRows(stash)
    If IsEmpty(stashData) Then Exit Function

    key = StashKey(ws)
    For i = 1 To UBound(stashData, 1)
        If stashData(i, 1) = key Then
            HasMarks = True
            Exit Function
        End If
    Next i
End Function

' Append the sheet's marker row and the fill of each of its formula cells
Private Sub StoreOriginalFormatting(ws As Worksheet)
    Dim stash As Worksheet
    Set stash = StashSheet()
    If stash Is Nothing Then
        Set stash = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        stash.Name = STASH_SHEET
        stash.Visible = xlSheetVeryHidden
        stash.Columns("A:B").NumberFormat = "@"  ' Keys and addresses stay text
    End If

    Dim formulaCells As Range
    On Error Resume Next
    Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    Dim rowCount As Long
    rowCount = 1
    If Not formulaCells Is Nothing Then rowCount = rowCount + formulaCells.Cells.Count

    Dim data() As Variant
    Dim key As String
    Dim i As Long
    Dim cell As Range
    ReDim data(1 To rowCount, 1 To 5)
    key = StashKey(ws)
    data(1, 1) = key
    i = 1
    If Not formulaCells Is Nothing Then
        For Each cell In formulaCells.Cells
            i = i + 1
            data(i, 1) = key
            data(i, 2) = cell.Address
            data(i, 3) = cell.Interior.Pattern
            data(i, 4) = cell.Interior.Color
            data(i, 5) = cell.Interior.PatternColor
        Next cell
    End If

    Dim nextRow As Long
    nextRow = stash.Cells(stash.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow = 2 And IsEmpty(stash.Cells(1, 1).Value) Then nextRow = 1
    stash.Cells(nextRow, 1).Resize(rowCount, 5).Value = data
End Sub

' Restore the sheet's original fills and drop its rows from the stash;
' delete the stash once no sheet is marked
Private Sub RemoveMarks(ws As Worksheet)
    Dim stash As Worksheet
    Dim stashData As Variant
    Set stash = StashSheet()
    If stash Is Nothing Then Exit Sub
    stashData = StashRows(stash)
    If IsEmpty(stashData) Then Exit Sub

    Application.ScreenUpdating = False

    Dim key As String
    Dim kept() As Variant
    Dim keptCount As Long
    Dim i As Long
    Dim j As Long
    key = StashKey(ws)
    ReDim kept(1 To UBound(stashData, 1), 1 To 5)
    For i = 1 To UBound(stashData, 1)
        If stashData(i, 1) = key Then
            If Len(stashData(i, 2)) > 0 Then
                With ws.Range(stashData(i, 2)).Interior
                    .Pattern = stashData(i, 3)
                    If .Pattern <> xlNone Then
                        .Color = stashData(i, 4)
                        .PatternColor = stashData(i, 5)
                    End If
                End With
            End If
        Else
            keptCount = keptCount + 1
            For j = 1 To 5
                kept(keptCount, j) = stashData(i, j)
            Next j
        End If
    Next i

    If keptCount = 0 Then
        Application.DisplayAlerts = False
        stash.Visible = xlSheetVisible
        stash.Delete
        Application.DisplayAlerts = True
    Else
        stash.Cells.ClearContents
        stash.Cells(1, 1).Resize(keptCount, 5).Value = kept
    End If

    Application.ScreenUpdating = True
End Sub
