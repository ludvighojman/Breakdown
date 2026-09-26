VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAutoColor
   Caption         =   "Auto-Color Settings"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmAutoColor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAutoColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NAME_PREFIX As String = "AutoColor_"

' UI Controls
Private WithEvents btnInputColor As MSForms.CommandButton
Private WithEvents btnFormulaColor As MSForms.CommandButton
Private WithEvents btnWorksheetLinkColor As MSForms.CommandButton
Private WithEvents btnExternalColor As MSForms.CommandButton
Private WithEvents btnResetDefaults As MSForms.CommandButton
Private WithEvents btnSave As MSForms.CommandButton

Public Sub InitializeInPanel(parentFrame As MSForms.Frame)
    ' Create labels and color buttons
    CreateControls parentFrame
    
    ' Load saved colors
    LoadSavedColors
End Sub

Private Sub CreateControls(parentFrame As MSForms.Frame)
    Dim top As Long: top = 10
    Dim labelWidth As Long: labelWidth = 150
    Dim buttonWidth As Long: buttonWidth = 60
    Dim height As Long: height = 20
    Dim spacing As Long: spacing = 25
    
    ' Hard-coded inputs
    CreateLabelAndButton parentFrame, "Hard-coded inputs:", "btnInputColor", top, labelWidth, buttonWidth, height
    Set btnInputColor = parentFrame.Controls("btnInputColor")
    
    ' Formulas on the same sheet
    top = top + spacing
    CreateLabelAndButton parentFrame, "Formulas (same sheet):", "btnFormulaColor", top, labelWidth, buttonWidth, height
    Set btnFormulaColor = parentFrame.Controls("btnFormulaColor")
    
    ' Formulas that use other sheets
    top = top + spacing
    CreateLabelAndButton parentFrame, "Formulas (other sheets):", "btnWorksheetLinkColor", top, labelWidth, buttonWidth, height
    Set btnWorksheetLinkColor = parentFrame.Controls("btnWorksheetLinkColor")
    
    ' Other workbooks, external data and errors
    top = top + spacing
    CreateLabelAndButton parentFrame, "External links / errors:", "btnExternalColor", top, labelWidth, buttonWidth, height
    Set btnExternalColor = parentFrame.Controls("btnExternalColor")
    
    ' Add Reset Defaults button at the bottom
    Set btnResetDefaults = parentFrame.Controls.Add("Forms.CommandButton.1", "btnResetDefaults")
    With btnResetDefaults
        .Left = 10
        .Top = top + spacing * 8  ' Position below all other controls
        .Width = 120
        .Height = height
        .Caption = "Reset to Defaults"
    End With
    
    ' Add Save button next to Reset Defaults
    Set btnSave = parentFrame.Controls.Add("Forms.CommandButton.1", "btnSave")
    With btnSave
        .Left = 140  ' Position it next to Reset Defaults
        .Top = btnResetDefaults.Top
        .Width = 120
        .Height = btnResetDefaults.Height
        .Caption = "Save Changes"
    End With
End Sub

Private Sub CreateLabelAndButton(parentFrame As MSForms.Frame, labelText As String, buttonName As String, _
                               top As Long, labelWidth As Long, buttonWidth As Long, height As Long)
    ' Create label
    Dim lbl As MSForms.Label
    Set lbl = parentFrame.Controls.Add("Forms.Label.1")
    With lbl
        .Left = 10
        .Top = top
        .Width = labelWidth
        .Height = height
        .Caption = labelText
    End With
    
    ' Create color button
    Dim btn As MSForms.CommandButton
    Set btn = parentFrame.Controls.Add("Forms.CommandButton.1", buttonName)
    With btn
        .Left = labelWidth + 20
        .Top = top
        .Width = buttonWidth
        .Height = height
        .Caption = "Color"
    End With
End Sub

Private Sub LoadSavedColors()
    ' Load colors from Names or set defaults if not found
    btnInputColor.BackColor = GetSavedColor("Input", 16711680)         ' Blue
    btnFormulaColor.BackColor = GetSavedColor("Formula", 0)            ' Black
    btnWorksheetLinkColor.BackColor = GetSavedColor("WorksheetLink", 32768)     ' Green
    btnExternalColor.BackColor = GetSavedColor("ExternalLink", 255)    ' Red
End Sub

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

Private Sub SaveColor(colorName As String, colorValue As Long)
    On Error Resume Next
    ThisWorkbook.Names.Add NAME_PREFIX & colorName, "=" & colorValue
    On Error GoTo 0
End Sub

Private Function ShowColorDialog() As Long
    On Error Resume Next
    ShowColorDialog = -1
    
    ' Store current workbook colors(1) to restore later
    Dim originalColor As Long
    originalColor = ActiveWorkbook.Colors(1)
    
    ' Show color picker
    If Application.Dialogs(xlDialogEditColor).Show(1) Then
        ShowColorDialog = ActiveWorkbook.Colors(1)
    End If
    
    ' Restore original color
    ActiveWorkbook.Colors(1) = originalColor
    On Error GoTo 0
End Function

' Color button click events
Private Sub btnInputColor_Click()
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnInputColor.BackColor = newColor
    End If
End Sub

Private Sub btnFormulaColor_Click()
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnFormulaColor.BackColor = newColor
    End If
End Sub

Private Sub btnWorksheetLinkColor_Click()
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnWorksheetLinkColor.BackColor = newColor
    End If
End Sub

Private Sub btnExternalColor_Click()
    Dim newColor As Long
    newColor = ShowColorDialog
    If newColor <> -1 Then
        btnExternalColor.BackColor = newColor
    End If
End Sub

' Add the reset button click handler
Private Sub btnResetDefaults_Click()
    If MsgBox("Are you sure you want to reset all colors to their defaults?", _
              vbQuestion + vbYesNo, "Reset Colors") = vbYes Then
              
        ' Reset all colors to defaults
        btnInputColor.BackColor = 16711680        ' Blue
        btnFormulaColor.BackColor = 0             ' Black
        btnWorksheetLinkColor.BackColor = 32768   ' Green
        btnExternalColor.BackColor = 255          ' Red
        
        MsgBox "Colors have been reset to defaults. Click 'Save Changes' to persist these changes.", vbInformation
    End If
End Sub

' Add save button click handler
Private Sub btnSave_Click()
    ' Save all current colors
    SaveColor "Input", btnInputColor.BackColor
    SaveColor "Formula", btnFormulaColor.BackColor
    SaveColor "WorksheetLink", btnWorksheetLinkColor.BackColor
    SaveColor "ExternalLink", btnExternalColor.BackColor
    
    ' Save the workbook to persist the changes
    ThisWorkbook.Save
    
    MsgBox "Colors saved successfully!", vbInformation
End Sub 