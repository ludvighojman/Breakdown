' clsDynamicButtonHandler
Option Explicit

Public WithEvents DynamicButton As MSForms.CommandButton
Private ButtonRole As String
Private ParentForm As Object

Public Sub Initialize(ByRef Button As MSForms.CommandButton, ByVal Role As String, ByRef Parent As Object)
    Set DynamicButton = Button
    ButtonRole = Role
    Set ParentForm = Parent
End Sub

Private Sub DynamicButton_Click()
    On Error GoTo ErrorHandler
    
    Select Case ButtonRole
        Case "Save"
            ParentForm.btnSave_Click
        Case "Reset"
            ParentForm.btnReset_Click
    End Select
    Exit Sub
    
ErrorHandler:
    Resume Next
End Sub
