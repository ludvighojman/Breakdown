Option Explicit

Private txtMaxDepth As MSForms.TextBox
Private lblMaxDepth As MSForms.Label
Private txtSafetyLimit As MSForms.TextBox
Private lblSafetyLimit As MSForms.Label
Private optThemeDark As MSForms.OptionButton
Private optThemeLight As MSForms.OptionButton
Private btnSave As MSForms.CommandButton
Private btnReset As MSForms.CommandButton
Private DynamicButtonHandlers As Collection

Public Sub InitializeInPanel(parentFrame As MSForms.Frame)
    
    ' Initialize the collection
    If DynamicButtonHandlers Is Nothing Then Set DynamicButtonHandlers = New Collection
    
    ' Create title label
    Dim lblTitle As MSForms.Label
    Set lblTitle = parentFrame.Controls.Add("Forms.Label.1", "lblTitle")
    With lblTitle
        .Left = 10
        .Top = 10
        .Width = 400
        .Height = 20
        .Caption = "Formula Tracing Settings"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    ' Create Max Depth controls
    Set lblMaxDepth = parentFrame.Controls.Add("Forms.Label.1", "lblMaxDepth")
    With lblMaxDepth
        .Left = 10
        .Top = 40
        .Width = 300
        .Height = 15
        .Caption = "Maximum Trace Depth (1-50):"
    End With
    
    Set txtMaxDepth = parentFrame.Controls.Add("Forms.TextBox.1", "txtMaxDepth")
    With txtMaxDepth
        .Left = 10
        .Top = 60
        .Width = 100
        .Height = 20
        .text = GetSavedMaxDepth()
    End With
    
    ' Create description label for Max Depth
    Dim lblMaxDepthDesc As MSForms.Label
    Set lblMaxDepthDesc = parentFrame.Controls.Add("Forms.Label.1", "lblMaxDepthDesc")
    With lblMaxDepthDesc
        .Left = 120
        .Top = 60
        .Width = 250
        .Height = 20
        .Caption = "Controls recursion depth for precedent/dependent analysis"
        .Font.Size = 8
        .ForeColor = &H808080
    End With
    
    ' Create Safety Limit controls
    Set lblSafetyLimit = parentFrame.Controls.Add("Forms.Label.1", "lblSafetyLimit")
    With lblSafetyLimit
        .Left = 10
        .Top = 90
        .Width = 300
        .Height = 15
        .Caption = "Safety Navigation Limit (10-1000):"
    End With
    
    Set txtSafetyLimit = parentFrame.Controls.Add("Forms.TextBox.1", "txtSafetyLimit")
    With txtSafetyLimit
        .Left = 10
        .Top = 110
        .Width = 100
        .Height = 20
        .text = GetSavedSafetyLimit()
    End With
    
    ' Create description label for Safety Limit
    Dim lblSafetyLimitDesc As MSForms.Label
    Set lblSafetyLimitDesc = parentFrame.Controls.Add("Forms.Label.1", "lblSafetyLimitDesc")
    With lblSafetyLimitDesc
        .Left = 120
        .Top = 110
        .Width = 250
        .Height = 20
        .Caption = "Prevents infinite loops and Excel crashes with complex references"
        .Font.Size = 8
        .ForeColor = &H808080
    End With
    
    ' Theme of the trace windows
    Dim lblTheme As MSForms.Label
    Set lblTheme = parentFrame.Controls.Add("Forms.Label.1", "lblTheme")
    With lblTheme
        .Left = 10
        .Top = 140
        .Width = 300
        .Height = 15
        .Caption = "Trace window theme:"
    End With

    Set optThemeDark = AddThemeOption(parentFrame, "optThemeDark", "Dark", 10)
    Set optThemeLight = AddThemeOption(parentFrame, "optThemeLight", "Light", 80)
    ShowTheme GetSetting("Breakdown", "FormulaTracing", "Theme", "Dark")

    ' Create Save button
    Set btnSave = parentFrame.Controls.Add("Forms.CommandButton.1", "btnSave")
    With btnSave
        .Left = 10
        .Top = 200
        .Width = 100
        .Height = 25
        .Caption = "Save Settings"
    End With
    
    ' Create Reset button
    Set btnReset = parentFrame.Controls.Add("Forms.CommandButton.1", "btnReset")
    With btnReset
        .Left = 120
        .Top = 200
        .Width = 100
        .Height = 25
        .Caption = "Reset to Default"
    End With
    
    ' Attach button handlers
    AttachButtonHandler btnSave, "Save"
    AttachButtonHandler btnReset, "Reset"
    
End Sub

' One of the Dark / Light choices; the two share a group, so picking one
' clears the other
Private Function AddThemeOption(parentFrame As MSForms.Frame, ctlName As String, captionText As String, ByVal leftPos As Single) As MSForms.OptionButton
    Dim opt As MSForms.OptionButton
    Set opt = parentFrame.Controls.Add("Forms.OptionButton.1", ctlName)
    With opt
        .Left = leftPos
        .Top = 158
        .Width = 65
        .Height = 18
        .Caption = captionText
        .GroupName = "TraceTheme"
    End With
    Set AddThemeOption = opt
End Function

Private Sub ShowTheme(themeName As String)
    optThemeLight.Value = (themeName = "Light")
    optThemeDark.Value = Not optThemeLight.Value
End Sub

Private Function SelectedTheme() As String
    If optThemeLight.Value Then SelectedTheme = "Light" Else SelectedTheme = "Dark"
End Function

Private Sub AttachButtonHandler(ByRef Button As MSForms.CommandButton, ByVal Role As String)
    
    Dim ButtonHandler As clsDynamicButtonHandler
    Set ButtonHandler = New clsDynamicButtonHandler
    ButtonHandler.Initialize Button, Role, Me
    
    If DynamicButtonHandlers Is Nothing Then Set DynamicButtonHandlers = New Collection
    DynamicButtonHandlers.Add ButtonHandler
End Sub

Public Sub btnSave_Click()
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    Dim maxDepth As Long, safetyLimit As Long
    
    If Not IsNumeric(txtMaxDepth.text) Then
        MsgBox "Maximum Trace Depth must be a number between 1 and 50.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(txtSafetyLimit.text) Then
        MsgBox "Safety Navigation Limit must be a number between 10 and 1000.", vbExclamation
        Exit Sub
    End If
    
    maxDepth = CLng(txtMaxDepth.text)
    safetyLimit = CLng(txtSafetyLimit.text)
    
    If maxDepth < 1 Or maxDepth > 50 Then
        MsgBox "Maximum Trace Depth must be between 1 and 50.", vbExclamation
        Exit Sub
    End If
    
    If safetyLimit < 10 Or safetyLimit > 1000 Then
        MsgBox "Safety Navigation Limit must be between 10 and 1000.", vbExclamation
        Exit Sub
    End If
    
    ' Save to registry
    SaveFormulaTracingSettings maxDepth, safetyLimit, SelectedTheme()
    
    MsgBox "Formula Tracing settings saved. A new theme applies to the next trace window you open.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error saving settings: " & Err.Description, vbCritical
End Sub

Public Sub btnReset_Click()
    On Error GoTo ErrorHandler
    
    ' Confirm reset
    Dim result As VbMsgBoxResult
    result = MsgBox("Reset Formula Tracing settings to default values?" & vbNewLine & _
                   "Max Depth: 10" & vbNewLine & _
                   "Safety Limit: 100" & vbNewLine & _
                   "Theme: Dark", vbYesNo + vbQuestion, "Reset Settings")
    
    If result = vbYes Then
        ' Reset to defaults
        txtMaxDepth.text = "10"
        txtSafetyLimit.text = "100"
        ShowTheme "Dark"
        
        ' Save defaults to registry
        SaveFormulaTracingSettings 10, 100, "Dark"
        
        MsgBox "Settings reset to default values!", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error resetting settings: " & Err.Description, vbCritical
End Sub

Private Function GetSavedMaxDepth() As String
    On Error Resume Next
    GetSavedMaxDepth = GetSetting("Breakdown", "FormulaTracing", "MaxDepth", "10")
    If Err.Number <> 0 Then GetSavedMaxDepth = "10"
    On Error GoTo 0
End Function

Private Function GetSavedSafetyLimit() As String
    On Error Resume Next
    GetSavedSafetyLimit = GetSetting("Breakdown", "FormulaTracing", "SafetyLimit", "100")
    If Err.Number <> 0 Then GetSavedSafetyLimit = "100"
    On Error GoTo 0
End Function

Private Sub SaveFormulaTracingSettings(maxDepth As Long, safetyLimit As Long, themeName As String)
    On Error GoTo ErrorHandler
    
    ' Save to the user's preferences (the registry on Windows, a plist on Mac)
    SaveSetting "Breakdown", "FormulaTracing", "MaxDepth", CStr(maxDepth)
    SaveSetting "Breakdown", "FormulaTracing", "SafetyLimit", CStr(safetyLimit)
    SaveSetting "Breakdown", "FormulaTracing", "Theme", themeName
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "SaveFormulaTracingSettings", Err.Description
End Sub
