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
Private WithEvents cboInputBorder As MSForms.ComboBox

' Input cell format: a row of fill swatches, a border list, and a preview
Private Const SWATCH_SIZE As Single = 18
Private Const SWATCH_GAP As Single = 5
Private Const PREVIEW_WIDTH As Single = 110
Private Const PREVIEW_HEIGHT As Single = 22

Private FillSwatches As Collection
Private SwatchHandlers As Collection
Private lblFillRing As MSForms.Label
Private lblFillName As MSForms.Label
Private imgPreview As MSForms.Image
Private lblPreviewText As MSForms.Label
Private SelectedFill As Long

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

    ' Input cell format: a fill, a border, and a preview in the input colour
    top = top + spacing + 15
    AddCaption parentFrame, "Input cell format:", 10, top, labelWidth
    top = top + spacing
    AddCaption parentFrame, "Fill", 10, top + 2, 45
    AddFillSwatches parentFrame, 60, top
    top = top + spacing + 3
    AddCaption parentFrame, "Border", 10, top + 2, 45
    Set cboInputBorder = parentFrame.Controls.Add("Forms.ComboBox.1", "cboInputBorder")
    With cboInputBorder
        .Left = 60
        .Top = top
        .Width = 200
        .Height = height
        .Style = fmStyleDropDownList
    End With
    Dim borderIndex As Long
    For borderIndex = 0 To INPUT_BORDER_COUNT
        cboInputBorder.AddItem InputBorderName(borderIndex)
    Next borderIndex
    top = top + spacing + 5
    AddCaption parentFrame, "Preview", 10, top + 4, 45
    AddPreview parentFrame, 60, top

    ' Add Reset Defaults button at the bottom
    Set btnResetDefaults = parentFrame.Controls.Add("Forms.CommandButton.1", "btnResetDefaults")
    With btnResetDefaults
        .Left = 10
        .Top = top + spacing * 3  ' Below all other controls
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

Private Function AddCaption(parentFrame As MSForms.Frame, captionText As String, ByVal leftPos As Single, _
                            ByVal topPos As Single, ByVal captionWidth As Single) As MSForms.Label
    Set AddCaption = parentFrame.Controls.Add("Forms.Label.1")
    With AddCaption
        .Move leftPos, topPos, captionWidth, 16
        .Caption = captionText
    End With
End Function

' The fills as a row of swatches. Clicking one picks it, and a ring behind it
' marks it.
Private Sub AddFillSwatches(parentFrame As MSForms.Frame, ByVal leftPos As Single, ByVal topPos As Single)
    Dim fillIndex As Long
    Dim swatch As MSForms.Label
    Dim handler As clsDynamicButtonHandler

    ' Added first, so it stays behind the swatches
    Set lblFillRing = parentFrame.Controls.Add("Forms.Label.1", "lblFillRing")
    With lblFillRing
        .Move leftPos - 3, topPos - 3, SWATCH_SIZE + 6, SWATCH_SIZE + 6
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(31, 31, 31)
    End With

    Set FillSwatches = New Collection
    Set SwatchHandlers = New Collection
    For fillIndex = 0 To INPUT_FILL_COUNT
        Set swatch = parentFrame.Controls.Add("Forms.Label.1", "lblFill" & fillIndex)
        With swatch
            .Move leftPos + fillIndex * (SWATCH_SIZE + SWATCH_GAP), topPos, SWATCH_SIZE, SWATCH_SIZE
            .Tag = CStr(fillIndex)
            .ControlTipText = InputFillName(fillIndex)
            .BackStyle = fmBackStyleOpaque
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = RGB(200, 200, 200)
            If InputFillColor(fillIndex) = NO_FILL Then
                .BackColor = RGB(255, 255, 255)
                .Caption = ChrW(&HD7)            ' A cross: no fill
                .Font.Size = 12
                .TextAlign = fmTextAlignCenter
                .ForeColor = RGB(128, 128, 128)
            Else
                .BackColor = InputFillColor(fillIndex)
            End If
        End With
        FillSwatches.Add swatch
        Set handler = New clsDynamicButtonHandler
        handler.InitializeLabel swatch, "PickFill", Me
        SwatchHandlers.Add handler
    Next fillIndex
End Sub

' A picture of the cell with a sample number on it, and the fill's name
Private Sub AddPreview(parentFrame As MSForms.Frame, ByVal leftPos As Single, ByVal topPos As Single)
    Set imgPreview = parentFrame.Controls.Add("Forms.Image.1", "imgPreview")
    With imgPreview
        ' A whole number of screen pixels, so the picture maps one to one
        .Move leftPos, topPos, Int(PREVIEW_WIDTH * PREVIEW_SCALE) / PREVIEW_SCALE, Int(PREVIEW_HEIGHT * PREVIEW_SCALE) / PREVIEW_SCALE
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .PictureSizeMode = fmPictureSizeModeStretch
        .PictureAlignment = fmPictureAlignmentTopLeft
    End With

    Set lblPreviewText = parentFrame.Controls.Add("Forms.Label.1", "lblPreviewText")
    With lblPreviewText
        .Move leftPos + 2, topPos + 4, imgPreview.Width - 6, imgPreview.Height - 6
        .Caption = "1,250.0"
        .Font.Size = 11
        .TextAlign = fmTextAlignRight
        .WordWrap = False
        .BackStyle = fmBackStyleTransparent
    End With

    Set lblFillName = AddCaption(parentFrame, "", leftPos + imgPreview.Width + 10, topPos + 5, 220)
    lblFillName.Font.Size = 8
    lblFillName.ForeColor = RGB(128, 128, 128)
End Sub

' Show the chosen fill and border: the ring on the swatch, and the preview in
' the input colour
Private Sub ShowPreview()
    Dim borderIndex As Long
    Dim fillColor As Long
    Dim edgeStyle As Long
    Dim edgeWeight As Long
    Dim edgeColor As Long
    Dim pic As Object

    If imgPreview Is Nothing Then Exit Sub
    borderIndex = cboInputBorder.ListIndex
    If borderIndex < 0 Then borderIndex = 0

    lblFillRing.Move FillSwatches(SelectedFill + 1).Left - 3, FillSwatches(SelectedFill + 1).Top - 3
    lblFillName.Caption = InputFillName(SelectedFill)
    lblPreviewText.ForeColor = btnInputColor.BackColor

    Set pic = InputCellPicture(Int(imgPreview.Width * PREVIEW_SCALE + 0.5), Int(imgPreview.Height * PREVIEW_SCALE + 0.5), _
                               SelectedFill, borderIndex)
    If Not pic Is Nothing Then
        imgPreview.BorderStyle = fmBorderStyleNone
        Set imgPreview.Picture = pic
        Exit Sub
    End If

    ' No picture: the fill, and a solid line for any border
    fillColor = InputFillColor(SelectedFill)
    If fillColor = NO_FILL Then fillColor = RGB(255, 255, 255)
    GetInputBorder borderIndex, edgeStyle, edgeWeight, edgeColor
    imgPreview.BackStyle = fmBackStyleOpaque
    imgPreview.BackColor = fillColor
    If edgeStyle = xlNone Then
        imgPreview.BorderStyle = fmBorderStyleNone
    Else
        imgPreview.BorderStyle = fmBorderStyleSingle
        imgPreview.BorderColor = edgeColor
    End If
End Sub

' A fill swatch was clicked
Public Sub PickInputFill(ByVal fillIndex As Long)
    SelectedFill = fillIndex
    ShowPreview
End Sub

Private Sub cboInputBorder_Change()
    ShowPreview
End Sub

Private Sub LoadSavedColors()
    ' Load colors from Names or set defaults if not found
    btnInputColor.BackColor = GetSavedColor("Input", 16711680)         ' Blue
    btnFormulaColor.BackColor = GetSavedColor("Formula", 0)            ' Black
    btnWorksheetLinkColor.BackColor = GetSavedColor("WorksheetLink", 32768)     ' Green
    btnExternalColor.BackColor = GetSavedColor("ExternalLink", 255)    ' Red
    SelectedFill = SavedInputFill()
    cboInputBorder.ListIndex = SavedInputBorder()
    ShowPreview
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
        ShowPreview
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
    If MsgBox("Are you sure you want to reset all colors and the input cell format to their defaults?", _
              vbQuestion + vbYesNo, "Reset Colors") = vbYes Then
              
        ' Reset all colors to defaults
        btnInputColor.BackColor = 16711680        ' Blue
        btnFormulaColor.BackColor = 0             ' Black
        btnWorksheetLinkColor.BackColor = 32768   ' Green
        btnExternalColor.BackColor = 255          ' Red
        SelectedFill = DEFAULT_INPUT_FILL
        cboInputBorder.ListIndex = DEFAULT_INPUT_BORDER
        ShowPreview
        
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
    SaveInputFormat SelectedFill, cboInputBorder.ListIndex
    
    ' Save the workbook to persist the changes
    ThisWorkbook.Save
    
    MsgBox "Colors saved successfully!", vbInformation
End Sub 