' frmSettingsManager
Option Explicit
' Control declarations
Private WithEvents lstCategories As MSForms.ListBox
Private AutoColorPanel As MSForms.Frame
Private autoColorSettings As frmAutoColor
Private ErrorPanel As MSForms.Frame
Private formulaTracingSettings As frmFormulaTracingSettings
Private FormulaTracingPanel As MSForms.Frame

' Layout, matching the trace windows: accent band, header with title and
' breadcrumb, category sidebar, settings page, status bar
Private Const ACCENT_HEIGHT As Long = 3
Private Const HEADER_HEIGHT As Long = 44       ' Below the accent band
Private Const SIDEBAR_WIDTH As Long = 160
Private Const CONTENT_PADDING As Long = 12
Private Const PANEL_WIDTH As Long = 410        ' Size the settings pages are laid out for
Private Const PANEL_HEIGHT As Long = 450
Private Const STATUS_HEIGHT As Long = 22
Private Const CAPTION_HEIGHT As Long = 14
Private Const FORM_WIDTH As Long = 600
Private Const FORM_HEIGHT As Long = 570

' Font sizes in points
Private Const TITLE_SIZE As Single = 13
Private Const TEXT_SIZE As Single = 11
Private Const CAPTION_SIZE As Single = 10

' Colours (&HBBGGRR), the same Excel-style palette as the trace windows
Private Const COLOR_ACCENT As Long = &H467321      ' RGB(33, 115, 70), Excel green
Private Const COLOR_WHITE As Long = &HFFFFFF
Private Const COLOR_HEADER As Long = &HF7F7F7      ' RGB(247, 247, 247)
Private Const COLOR_SIDEBAR As Long = &HF3F3F3     ' RGB(243, 243, 243)
Private Const COLOR_RULE As Long = &HD4D4D4        ' RGB(212, 212, 212)
Private Const COLOR_FIELD_BORDER As Long = &HC8C8C8 ' RGB(200, 200, 200)
Private Const COLOR_TEXT As Long = &H1F1F1F        ' RGB(31, 31, 31)
Private Const COLOR_SECONDARY As Long = &H616161   ' RGB(97, 97, 97)
Private Const SYSTEM_TEXT_COLOR As Long = &H80000012  ' Default text colour of new controls


Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    ' Initialize form layout
    InitializeFormLayout
    
    ' Backgrounds and rules first, so everything added after them is drawn on top
    AddPanel "lblAccent", COLOR_ACCENT
    AddPanel "lblHeaderArea", COLOR_HEADER
    AddPanel "lblHeaderRule", COLOR_RULE
    AddPanel "lblSidebar", COLOR_SIDEBAR
    AddPanel "lblSidebarRule", COLOR_RULE
    AddPanel "lblStatusBar", COLOR_SIDEBAR
    AddPanel "lblStatusRule", COLOR_RULE
    
    ' Create navigation listbox with event handling
    Set lstCategories = Me.Controls.Add("Forms.ListBox.1", "lstCategories")
    With lstCategories
        .BackColor = COLOR_SIDEBAR
        .ForeColor = COLOR_TEXT
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .Font.Name = UiFontName()
        .Font.Size = TEXT_SIZE
    End With
    
    ' Header and status bar text
    AddTextLabel "lblTitle", "Settings", TITLE_SIZE, True, COLOR_TEXT
    AddTextLabel "lblBreadcrumb", "", CAPTION_SIZE, False, COLOR_SECONDARY
    AddTextLabel "lblStatus", "Settings apply to every workbook. Each page has its own Save button.", _
                 CAPTION_SIZE, False, COLOR_SECONDARY
    
    ' Create panels
    InitializePanels
    
    ' Give every settings page the same look
    ApplyTheme AutoColorPanel
    ApplyTheme ErrorPanel
    ApplyTheme FormulaTracingPanel
    
    InitializeHierarchyList
    
    LayoutChrome
    
    Exit Sub
    
ErrorHandler:
    Resume Next
End Sub

Private Sub InitializePanels()
    Set AutoColorPanel = AddSettingsPage("AutoColorPanel")
    Set ErrorPanel = AddSettingsPage("ErrorPanel")
    Set FormulaTracingPanel = AddSettingsPage("FormulaTracingPanel")

    ' Build each settings page inside its frame
    Set autoColorSettings = New frmAutoColor
    autoColorSettings.InitializeInPanel AutoColorPanel
    Dim errorSettings As New frmErrorSettings
    errorSettings.InitializeInPanel ErrorPanel
    Set formulaTracingSettings = New frmFormulaTracingSettings
    formulaTracingSettings.InitializeInPanel FormulaTracingPanel
End Sub

' An empty, hidden frame for one settings page; LayoutChrome places it
Private Function AddSettingsPage(ctlName As String) As MSForms.Frame
    Dim pageFrame As MSForms.Frame
    Set pageFrame = Me.Controls.Add("Forms.Frame.1", ctlName)
    With pageFrame
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    Set AddSettingsPage = pageFrame
End Function

' Show the requested panel and hide others
Private Sub ShowPanel(panelName As String)
    AutoColorPanel.Visible = (panelName = "Auto-Color")
    ErrorPanel.Visible = (panelName = "Error")
    FormulaTracingPanel.Visible = (panelName = "Formula Tracing")
    
    ' Breadcrumb in the header: group > page
    Dim groupName As String
    If panelName = "Formula Tracing" Then groupName = "Formulas" Else groupName = "Formatting"
    Me.Controls("lblBreadcrumb").Caption = groupName & "  " & ChrW(&H203A) & "  " & panelName
    
End Sub

Private Sub UserForm_Terminate()
    Set AutoColorPanel = Nothing
    Set autoColorSettings = Nothing
    Set ErrorPanel = Nothing
    Set formulaTracingSettings = Nothing
    Set FormulaTracingPanel = Nothing
    Set lstCategories = Nothing
End Sub

Private Sub lstCategories_Click()
    On Error GoTo ErrorHandler
    
    Dim selectedCategory As String
    selectedCategory = Trim(lstCategories.text)
    Dim i As Integer
    
    If lstCategories.List(lstCategories.ListIndex, 1) = "HEADER" Then
        ' Handle different headers - prevent selection by redirecting to first item under header
        Select Case selectedCategory
            Case "FORMATTING"
                lstCategories.ListIndex = 1  ' Select Auto-Color (first item under Formatting)
                ShowPanel "Auto-Color"
            Case "FORMULAS"
                ' Find Formula Tracing item index
                For i = 0 To lstCategories.ListCount - 1
                    If Trim(lstCategories.List(i, 0)) = "Formula Tracing" Then
                        lstCategories.ListIndex = i
                        ShowPanel "Formula Tracing"
                        Exit For
                    End If
                Next i
            Case Else
                lstCategories.ListIndex = 1  ' Default to Auto-Color
                ShowPanel "Auto-Color"
        End Select
        Exit Sub
    End If
    
    ' Remove any potential hidden characters
    selectedCategory = Replace(selectedCategory, vbTab, "")
    selectedCategory = Replace(selectedCategory, vbCr, "")
    selectedCategory = Replace(selectedCategory, vbLf, "")
    selectedCategory = Trim(selectedCategory)
    
    Select Case selectedCategory
        Case "Auto-Color", "Error", "Formula Tracing"
            ShowPanel selectedCategory
    End Select
    Exit Sub

ErrorHandler:
    Resume Next
End Sub

Private Sub InitializeFormLayout()
    Me.BackColor = COLOR_WHITE
    Me.Caption = "Settings"
    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT
End Sub

Private Sub InitializeHierarchyList()
    lstCategories.Clear
    
    With lstCategories
        .AddItem "FORMATTING"
        .List(.ListCount - 1, 1) = "HEADER"
        .AddItem "   Auto-Color"
        .AddItem "   Error"
        .AddItem "FORMULAS"
        .List(.ListCount - 1, 1) = "HEADER"
        .AddItem "   Formula Tracing"
        .ListIndex = 1
    End With
    
    ShowPanel "Auto-Color"
End Sub

' Lay out again once visible; Excel for Mac may not apply size changes to
' controls made before the form is shown
Private Sub UserForm_Activate()
    LayoutChrome
End Sub

' Place the accent band, header, sidebar, settings pages and status bar
Private Sub LayoutChrome()
    On Error Resume Next

    ' Client area size; fall back to estimates if the host doesn't report it
    Dim clientWidth As Single
    Dim clientHeight As Single
    clientWidth = Me.InsideWidth
    clientHeight = Me.InsideHeight
    If clientWidth <= 0 Then clientWidth = Me.Width - 4
    If clientHeight <= 0 Then clientHeight = Me.Height - 22  ' approx. title bar

    Dim bodyTop As Single
    Dim statusTop As Single
    bodyTop = ACCENT_HEIGHT + HEADER_HEIGHT + 1
    statusTop = clientHeight - STATUS_HEIGHT

    ' Accent band and header
    Me.Controls("lblAccent").Move 0, 0, clientWidth, ACCENT_HEIGHT
    Me.Controls("lblHeaderArea").Move 0, ACCENT_HEIGHT, clientWidth, HEADER_HEIGHT
    Me.Controls("lblHeaderRule").Move 0, ACCENT_HEIGHT + HEADER_HEIGHT, clientWidth, 1
    Me.Controls("lblTitle").Move CONTENT_PADDING, ACCENT_HEIGHT + 8, clientWidth - 2 * CONTENT_PADDING, 18
    Me.Controls("lblBreadcrumb").Move CONTENT_PADDING, ACCENT_HEIGHT + 26, clientWidth - 2 * CONTENT_PADDING, CAPTION_HEIGHT

    ' Category sidebar
    Me.Controls("lblSidebar").Move 0, bodyTop, SIDEBAR_WIDTH, statusTop - bodyTop
    Me.Controls("lblSidebarRule").Move SIDEBAR_WIDTH, bodyTop, 1, statusTop - bodyTop
    lstCategories.IntegralHeight = False
    lstCategories.Move 4, bodyTop + 6, SIDEBAR_WIDTH - 8, statusTop - bodyTop - 12

    ' Settings pages, all in the same place
    Dim panelLeft As Single
    Dim panelTop As Single
    Dim panelHeight As Single
    panelLeft = SIDEBAR_WIDTH + 1 + CONTENT_PADDING
    panelTop = bodyTop + CONTENT_PADDING
    panelHeight = statusTop - CONTENT_PADDING - panelTop
    If panelHeight > PANEL_HEIGHT Then panelHeight = PANEL_HEIGHT
    Dim panel As Variant
    For Each panel In Array(AutoColorPanel, ErrorPanel, FormulaTracingPanel)
        panel.Move panelLeft, panelTop, PANEL_WIDTH, panelHeight
    Next panel

    ' Status bar
    Me.Controls("lblStatusBar").Move 0, statusTop, clientWidth, STATUS_HEIGHT
    Me.Controls("lblStatusRule").Move 0, statusTop, clientWidth, 1
    Me.Controls("lblStatus").Move CONTENT_PADDING, statusTop + Int((STATUS_HEIGHT - CAPTION_HEIGHT) / 2), _
                                  clientWidth - 2 * CONTENT_PADDING, CAPTION_HEIGHT

    Me.Repaint
    On Error GoTo 0
End Sub

' Give the controls on a settings page the fonts and flat, light look of the
' trace windows. Font sizes are kept so nothing gets clipped. Colour swatch
' buttons keep their colour, and captions that already have their own colour
' keep it.
Private Sub ApplyTheme(panel As MSForms.Frame)
    On Error Resume Next
    panel.BackColor = COLOR_WHITE
    panel.BorderStyle = fmBorderStyleNone
    panel.SpecialEffect = fmSpecialEffectFlat

    Dim ctl As Object
    For Each ctl In panel.Controls
        ctl.Font.Name = UiFontName()
        Select Case TypeName(ctl)
            Case "Label", "CheckBox", "OptionButton"
                If ctl.BorderStyle = fmBorderStyleNone Then ctl.BackStyle = fmBackStyleTransparent
                If ctl.ForeColor = SYSTEM_TEXT_COLOR Then ctl.ForeColor = COLOR_TEXT
            Case "TextBox", "ComboBox", "ListBox"
                ctl.BackColor = COLOR_WHITE
                ctl.ForeColor = COLOR_TEXT
                ctl.SpecialEffect = fmSpecialEffectFlat
                ctl.BorderStyle = fmBorderStyleSingle
                ctl.BorderColor = COLOR_FIELD_BORDER
        End Select
    Next ctl
    On Error GoTo 0
End Sub

' Add a caption label with a transparent background
Private Function AddTextLabel(ctlName As String, captionText As String, fontSize As Single, isBold As Boolean, textColor As Long) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", ctlName)
    With lbl
        .Caption = captionText
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
        .ForeColor = textColor
        .Font.Name = UiFontName()
        .Font.Size = fontSize
        .Font.Bold = isBold
    End With
    Set AddTextLabel = lbl
End Function

' Add a flat coloured rectangle used as a background or rule
Private Function AddPanel(ctlName As String, fillColor As Long) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", ctlName)
    With lbl
        .Caption = ""
        .BackStyle = fmBackStyleOpaque
        .BackColor = fillColor
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleNone
    End With
    Set AddPanel = lbl
End Function

' Fonts that ship with the OS on each platform
Private Function UiFontName() As String
#If Mac Then
    UiFontName = "Helvetica Neue"
#Else
    UiFontName = "Segoe UI"
#End If
End Function
