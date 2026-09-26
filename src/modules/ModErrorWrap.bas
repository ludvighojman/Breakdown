Attribute VB_Name = "ModErrorWrap"
' ModErrorWrap
Option Explicit

Sub WrapWithError(control As IRibbonControl)
    Dim selectedCell As Range
    Dim cellCount As Long
    Dim errorCount As Long
    Dim errorValue As String
    
    ' Get the saved error value or use default
    On Error Resume Next
    errorValue = ThisWorkbook.CustomDocumentProperties("ErrorValue")
    If Err.Number <> 0 Or errorValue = "" Then
        errorValue = "NA()"
    End If
    On Error GoTo 0
    
    cellCount = 0
    errorCount = 0

    ' Iterate through each cell in the selection
    For Each selectedCell In Selection
        ' Check if the cell contains a formula
        If selectedCell.HasFormula Then
            ' Wrap the formula with IFERROR using the saved error value
            selectedCell.Formula = "=IFERROR(" & Mid(selectedCell.Formula, 2) & ", " & errorValue & ")"
            cellCount = cellCount + 1
        Else
            errorCount = errorCount + 1
        End If
    Next selectedCell
    
    ' Provide feedback based on the results
    If cellCount = 0 Then
        MsgBox "None of the selected cells contained a formula that could be wrapped.", vbExclamation
    ElseIf errorCount > 0 Then
        MsgBox "Some of the selected cells did not contain formulas and were skipped.", vbExclamation
    End If
End Sub
