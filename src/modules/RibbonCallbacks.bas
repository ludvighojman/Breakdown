Attribute VB_Name = "RibbonCallbacks"
Option Explicit

' Ribbon button callbacks (see src/ribbon/customUI14.xml). Each runs its
' feature and reports an error instead of failing silently.

Public Sub FindAndDisplayPrecedents(control As IRibbonControl)
    On Error GoTo Failed
    TraceUtils.ShowTracePrecedents
    Exit Sub
Failed:
    ReportError "Trace Precedents", Err.Number, Err.Description
End Sub

Public Sub FindAndDisplayDependents(control As IRibbonControl)
    On Error GoTo Failed
    TraceUtils.ShowTraceDependents
    Exit Sub
Failed:
    ReportError "Trace Dependents", Err.Number, Err.Description
End Sub

Public Sub OnCheckHorizontalConsistency(control As IRibbonControl)
    On Error GoTo Failed
    FormulaConsistency.CheckHorizontalConsistency
    Exit Sub
Failed:
    ReportError "Horizontal Consistency", Err.Number, Err.Description
End Sub

Public Sub DoAutoColorCells(control As IRibbonControl)
    On Error GoTo Failed
    AutoColorModule.AutoColorCells
    Exit Sub
Failed:
    ReportError "Auto-color Numbers", Err.Number, Err.Description
End Sub

Public Sub ShowSettingsForm(control As IRibbonControl)
    On Error GoTo Failed
    Dim frm As New frmSettingsManager
    frm.Show vbModal
    Exit Sub
Failed:
    ReportError "Settings", Err.Number, Err.Description
End Sub

' Excel for Mac leaves the description of some errors empty; show the number then
Private Sub ReportError(feature As String, errorNumber As Long, description As String)
    If Len(description) = 0 Then description = "Error " & errorNumber
    MsgBox feature & " stopped with an error:" & vbNewLine & vbNewLine & description, vbExclamation, "Breakdown"
End Sub
