' clsDynamicButtonHandler
Option Explicit

Public WithEvents DynamicButton As MSForms.CommandButton
Public WithEvents DynamicLabel As MSForms.Label
Private ButtonRole As String
Private ParentForm As Object

Public Sub Initialize(ByRef Button As MSForms.CommandButton, ByVal Role As String, ByRef Parent As Object)
    Set DynamicButton = Button
    ButtonRole = Role
    Set ParentForm = Parent
End Sub

' A label that works as a button; the parent form gets the label's Tag
Public Sub InitializeLabel(ByRef Lbl As MSForms.Label, ByVal Role As String, ByRef Parent As Object)
    Set DynamicLabel = Lbl
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

Private Sub DynamicLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo ErrorHandler

    Select Case ButtonRole
        Case "PickFill"
            ParentForm.PickInputFill CLng(DynamicLabel.Tag)
    End Select
    Exit Sub

ErrorHandler:
    Resume Next
End Sub
