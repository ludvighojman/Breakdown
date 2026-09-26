VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmErrorSettings
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmErrorSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmErrorSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private txtErrorValue As MSForms.TextBox
Private lblErrorValue As MSForms.Label
Private btnSave As MSForms.CommandButton
Private DynamicButtonHandlers As Collection

Public Sub InitializeInPanel(parentFrame As MSForms.Frame)
    
    ' Initialize the collection
    If DynamicButtonHandlers Is Nothing Then Set DynamicButtonHandlers = New Collection
    
    ' Create and position controls
    Set txtErrorValue = parentFrame.Controls.Add("Forms.TextBox.1", "txtErrorValue")
    With txtErrorValue
        .Left = 10
        .Top = 30
        .Width = 290
        .Height = 20
        .Text = GetSavedErrorValue()
    End With
    
    Set lblErrorValue = parentFrame.Controls.Add("Forms.Label.1", "lblErrorValue")
    With lblErrorValue
        .Left = 10
        .Top = 10
        .Caption = "Error Value (e.g., NA(), 0, """")"
    End With
    
    Set btnSave = parentFrame.Controls.Add("Forms.CommandButton.1", "btnSave")
    With btnSave
        .Left = 310
        .Top = 30
        .Width = 80
        .Height = 25
        .Caption = "Save"
    End With
    
    ' Attach button handler
    AttachButtonHandler btnSave, "Save"
    
End Sub

Private Sub AttachButtonHandler(ByRef Button As MSForms.CommandButton, ByVal Role As String)
    
    Dim ButtonHandler As clsDynamicButtonHandler
    Set ButtonHandler = New clsDynamicButtonHandler
    ButtonHandler.Initialize Button, Role, Me
    
    DynamicButtonHandlers.Add ButtonHandler
End Sub

Public Sub btnSave_Click()
    SaveErrorValue txtErrorValue.Text
    MsgBox "Error value saved successfully!", vbInformation
End Sub

Private Function GetSavedErrorValue() As String
    On Error Resume Next
    GetSavedErrorValue = ThisWorkbook.CustomDocumentProperties("ErrorValue")
    If Err.Number <> 0 Or GetSavedErrorValue = "" Then
        GetSavedErrorValue = "NA()"
    End If
    On Error GoTo 0
End Function

Private Sub SaveErrorValue(value As String)
    
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties("ErrorValue").Delete
    On Error GoTo 0
    
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties.Add _
        Name:="ErrorValue", _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        value:=value
    On Error GoTo 0
    
    ThisWorkbook.Save
End Sub

Private Sub UserForm_Terminate()
    Set DynamicButtonHandlers = Nothing
End Sub
