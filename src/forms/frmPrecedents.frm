Option Explicit

' TreeNode structure for hierarchical display
Private Type TreeNode
    text As String
    level As Integer
    IsExpanded As Boolean
    hasChildren As Boolean
    IsVisible As Boolean
    originalData As String  ' Store original listbox data (address, value, formula)
    refIndex As Long        ' Which reference in the parent's formula this row is, or 0
    isRepeat As Boolean     ' The cell is listed again: not counted as a new precedent
    copyPending As Boolean  ' A repeat row whose rows are copied in when first opened
End Type

Private TreeData() As TreeNode
Private TreeCount As Integer
Private VisibleNodes() As Integer       ' Tree node index shown on each list row
Private VisibleCount As Integer
Private VisibleRails() As String        ' Per list row: "1"/"0" per level, whether that level's tree line continues below

' Workbook and sheet of the traced cell; addresses on that sheet are shown short
Private RootBook As String
Private RootSheet As String
Private TracedCell As Range
Private UniqueCount As Long             ' Cells in the tree, not counting repeat rows

' Trace-into history: Enter hides this window and adds it here, Shift+Enter
' shows the last one again. Shared by every window in one drill-down chain;
' closing the visible window closes the hidden ones too.
Private TraceHistory As Collection
Private SwitchingWindows As Boolean     ' Unloading for Shift+Enter, not closing

' The list is drawn from labels: a fixed number of row slots, filled from
' VisibleNodes starting at TopRow as the list scrolls
Private SelectedRow As Long             ' 0-based list row, -1 for none
Private TopRow As Long                  ' List row shown in the first slot
Private SlotCount As Long
Private MaxRailDepth As Long            ' Deepest tree-line label created so far
Private ListTop As Single
Private ListWidth As Single
Private IsRendering As Boolean          ' Ignore scrollbar events we cause ourselves

' Formula card: the formula shown and the reference drawn as a green chip
Private CardFormula As String
Private CardHighlightStart As Long
Private CardLines As Long
Private TokenCount As Long              ' Formula card labels created so far

Private UiFont As String                ' Chosen once per window by UiFontName

' What each rounded box and formula token currently shows, keyed by name, so
' redraws leave unchanged ones alone (see PlaceRoundBox and PlaceToken)
Private Shown As Collection
Private CircleSizes As Collection       ' Measured label size per corner-circle font size
Private UsePictures As Boolean          ' Corners are pictures (see CornerPicturesWork)

' Size grip: where in the grip the drag started, and whether a drag is resizing
Private WithEvents lblGrip As MSForms.Label
Private GripStartX As Single
Private GripStartY As Single
Private IsDragResizing As Boolean

' Controls with events: the filter box, the off-screen field that takes the
' keyboard while the list is active, the transparent layer that takes clicks
' on the rows, and the list scrollbar
Private WithEvents txtFilter As MSForms.TextBox
Private WithEvents txtKeys As MSForms.TextBox
Private WithEvents lblListHit As MSForms.Label
Private WithEvents scrList As MSForms.ScrollBar

' Resizer functionality
Private FormResizer As clResizer
Private MinWidth As Single
Private MinHeight As Single
Private LastWidth As Single
Private LastHeight As Single

' Minimum form size
Private Const FORM_WIDTH As Long = 340
Private Const FORM_HEIGHT As Long = 268

' Opening size and position
Private Const DEFAULT_SIZE_FACTOR As Single = 2  ' Open at twice the minimum size
Private Const WINDOW_MARGIN As Long = 20        ' Gap to the Excel window edge (clears the grid scrollbar)

' Layout
Private Const PADDING As Long = 12
Private Const TOP_BAR_HEIGHT As Long = 38
Private Const CARD_PADDING As Long = 12
Private Const CARD_CAPTION_HEIGHT As Long = 24  ' "FORMULA" caption and the value
Private Const FORMULA_LINE_HEIGHT As Long = 24
Private Const MAX_FORMULA_LINES As Long = 6     ' Longer formulas end in "..."
Private Const COLUMN_HEADER_HEIGHT As Long = 20
Private Const FOOTER_HEIGHT As Long = 20
Private Const ROW_HEIGHT As Long = 24
Private Const CHIP_HEIGHT As Long = 18
Private Const CHIP_PADDING As Long = 7          ' Text inset inside a chip
Private Const INDENT As Long = 22               ' Per tree level
Private Const RAIL_OFFSET As Long = 8           ' Tree line x within its parent's indent
Private Const CHEVRON_WIDTH As Long = 12
Private Const VALUE_COLUMN_WIDTH As Long = 96
Private Const FORMULA_COLUMN_SHARE As Single = 0.34  ' Formula column starts at this share of the list width
Private Const SCROLLBAR_WIDTH As Long = 12
Private Const CAPTION_HEIGHT As Long = 14
Private Const ESC_WIDTH As Long = 28            ' The "esc" hint box in the top bar
Private Const ESC_HEIGHT As Long = 16

' Corner radii of the rounded boxes (see PlaceRoundBox)
Private Const CHIP_RADIUS As Single = 4
Private Const ROW_RADIUS As Single = 6
Private Const CARD_RADIUS As Single = 9

' Rounded boxes are drawn on the device pixel grid (PX_POINTS points per
' pixel), with a picture of each corner (see CornerPicture in TraceUtils)
#If Mac Then
Private Const PX_POINTS As Single = 1           ' A standard (non-Retina) display
#Else
Private Const PX_POINTS As Single = 0.75        ' 96 dpi
#End If
Private Const BOX_PARTS As Long = 8             ' Controls per rounded box (see PlaceRoundBox)

' If corner pictures can't be used (see CornerPicturesWork), each corner is a
' filled circle character instead: Arial's black circle (U+25CF), whose ink is
' a circle 881/2048 of the font size across. Excel draws label text a little
' off the font's own metrics; CIRCLE_SHIFT_X (points) and CIRCLE_DROP (a
' share of the font size) put the circle back where it belongs, as measured
' on screen.
Private Const CIRCLE_FONT As String = "Arial"
Private Const CIRCLE_INK As Single = 0.4302
#If Mac Then
Private Const CIRCLE_SHIFT_X As Single = 1.24
Private Const CIRCLE_DROP As Single = 0.094
#Else
Private Const CIRCLE_SHIFT_X As Single = 0
Private Const CIRCLE_DROP As Single = 0.065
#End If

' Size grip in the bottom-right corner, dragged to resize the window
Private Const GRIP_SIZE As Long = 14
Private Const LEFT_BUTTON As Integer = 1        ' MouseMove Button value (MSForms has no named constant)

' Font sizes in points
Private Const TEXT_SIZE As Single = 11
Private Const CHIP_SIZE As Single = 10
Private Const CAPTION_SIZE As Single = 9
Private Const VALUE_SIZE As Single = 18

' Height of one line of text as a share of its font size, used to centre
' text vertically: a label draws its text from the top
Private Const LINE_HEIGHT_FACTOR As Single = 1.25

' Excel for Mac draws label text about 1.5 points higher and 1 point further
' left than the font's metrics say (measured on screen); text centred in a box
' is moved back by this much
#If Mac Then
Private Const TEXT_NUDGE_X As Single = 1
Private Const TEXT_NUDGE_Y As Single = 1.5
#Else
Private Const TEXT_NUDGE_X As Single = 0
Private Const TEXT_NUDGE_Y As Single = 0
#End If

' Characters that end a word in the formula card, so long formulas can wrap
Private Const FORMULA_BREAK_CHARS As String = ",;+-*/^&=<>() " & vbTab & vbCr & vbLf

' Shift bit of the KeyDown Shift argument (MSForms has no named constant for it)
Private Const SHIFT_KEY_MASK As Integer = 1

' Which formula the formula card shows for a selected row. Precedents: the
' formula that uses the selected cell (its nearest ancestor with a formula).
' Dependents: the selected cell's own formula, which uses its parent row.
' frmDependents is generated from this file with this constant set to False.
Private Const FORMULA_FROM_ANCESTOR As Boolean = True

' Colours (&HBBGGRR) of the theme chosen in Settings > Formula Tracing; set
' by LoadTheme when the window opens
Private ColorWindow As Long
Private ColorCard As Long
Private ColorCardBorder As Long     ' Card and esc outline
Private ColorRule As Long
Private ColorChip As Long
Private ColorGreen As Long          ' Traced cell, selected row, highlighted reference
Private ColorOnGreen As Long        ' Text on a green chip
Private ColorSelectedRow As Long
Private ColorRail As Long           ' Tree lines
Private ColorText As Long
Private ColorSecondary As Long
Private ColorCaption As Long


' ---------------------------------------------------------------------------
' Set-up (called by TraceUtils)
' ---------------------------------------------------------------------------

' Record the traced cell and show it in the top bar. Call before
' ConvertToTreeView so list addresses on the same sheet can be shortened.
Public Sub SetRoot(rootCell As Range)
    Set TracedCell = rootCell
    RootBook = rootCell.Worksheet.Parent.Name
    RootSheet = rootCell.Worksheet.Name
    Me.Controls("lblRootChip").Caption = rootCell.address(False, False)
    Me.Controls("lblSheetName").Caption = RootSheet
End Sub

' Continue the trace-into history of the window this one was opened from.
' Nothing starts a fresh history.
Public Sub SetHistory(history As Collection)
    If history Is Nothing Then
        Set TraceHistory = New Collection
    Else
        Set TraceHistory = history
    End If
End Sub

' Build the tree from the rows TraceUtils put in the designer list box and
' show it with the traced cell selected
Public Sub ConvertToTreeView()
    ParseListBoxToTree
    SelectedRow = 0
    TopRow = 0
    RefreshTree
    ShowFormulaContext 1
    UpdateStatus
End Sub

Private Sub UserForm_Initialize()
    LoadTheme
    Set Shown = New Collection
    Set CircleSizes = New Collection
    UsePictures = CornerPicturesWork()

    ' Set form caption and size
    Me.Caption = "Trace Precedents"
    Me.BackColor = ColorWindow
    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT

    ' Store minimum and current dimensions
    MinWidth = Me.Width
    MinHeight = Me.Height
    LastWidth = Me.Width
    LastHeight = Me.Height

    ' Initialize resizer
    Set FormResizer = New clResizer
    FormResizer.NewForm Me.Caption

    ' Initialize tree data
    ReDim TreeData(1 To 100)
    TreeCount = 0
    ReDim VisibleNodes(1 To 1)
    ReDim VisibleRails(1 To 1)
    VisibleCount = 0
    SelectedRow = -1
    Set TraceHistory = New Collection

    ' The designer list box only carries the trace rows from TraceUtils
    lstPrecedents.Visible = False

    ' Backgrounds and rules first, so everything added after them is on top
    AddPanel "lblTopRule", ColorRule
    AddPanel "lblFooterRule", ColorRule
    EnsureRoundBox "lblCard"
    EnsureRoundBox "lblEsc"
    EnsureRoundBox "lblSelection"

    ' Top bar: "Precedents of <cell>", the filter box, the sheet and an esc hint
    AddTextLabel("lblTopTitle", "Precedents of", TEXT_SIZE, False, ColorSecondary).AutoSize = True
    AddChip "lblRootChip"
    AddTextLabel "lblFilterHint", "Filter cells or formulas", TEXT_SIZE, False, ColorCaption
    Set txtFilter = Me.Controls.Add("Forms.TextBox.1", "txtFilter")
    With txtFilter
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .ForeColor = ColorText
        .Font.Name = UiFontName()
        .Font.Size = TEXT_SIZE
    End With
    AddTextLabel("lblSheetName", "", CAPTION_SIZE + 1, False, ColorSecondary).TextAlign = fmTextAlignRight
    AddTextLabel("lblEscHint", "esc", CAPTION_SIZE, False, ColorSecondary).TextAlign = fmTextAlignCenter

    ' Formula card: caption and value; the formula itself is laid out as tokens
    AddTextLabel "lblCardCaption", "FORMULA", CAPTION_SIZE, True, ColorCaption
    With AddTextLabel("lblCardValue", "", VALUE_SIZE, True, ColorText)
        .Font.Name = MonoFontName()
        .TextAlign = fmTextAlignRight
    End With

    ' Column captions
    AddTextLabel "lblColCell", "CELL", CAPTION_SIZE, True, ColorCaption
    AddTextLabel "lblColFormula", "FORMULA", CAPTION_SIZE, True, ColorCaption
    AddTextLabel("lblColValue", "VALUE", CAPTION_SIZE, True, ColorCaption).TextAlign = fmTextAlignRight

    ' Footer: count on the left, keyboard hints on the right
    AddTextLabel("lblCount", "", CAPTION_SIZE + 1, True, ColorText).AutoSize = True
    AddTextLabel("lblStatus", "", CAPTION_SIZE + 1, False, ColorSecondary).AutoSize = True
    With AddTextLabel("lblKeys", KeyHintsText(), CAPTION_SIZE + 1, False, ColorSecondary)
        .TextAlign = fmTextAlignRight
        .Font.Name = SystemFontName()  ' Has the arrow characters
    End With

    ' List scrollbar, the click layer over the rows, and the field that takes
    ' the keyboard while the list is active (kept off-screen)
    Set scrList = Me.Controls.Add("Forms.ScrollBar.1", "scrList")
    With scrList
        .Orientation = fmOrientationVertical
        .BackColor = ColorWindow
        .ForeColor = ColorCaption
        .Min = 0
        .Max = 0
        .SmallChange = 1
    End With
    Set lblListHit = Me.Controls.Add("Forms.Label.1", "lblListHit")
    With lblListHit
        .Caption = ""
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
    End With
    Set txtKeys = Me.Controls.Add("Forms.TextBox.1", "txtKeys")
    txtKeys.Move -200, -200, 20, 12

    ' Size grip: six dots in a triangle, under a transparent layer that takes
    ' the drag
    Dim dot As Long
    For dot = 0 To 5
        AddPanel "lblGripDot" & dot, ColorCaption
    Next dot
    Set lblGrip = Me.Controls.Add("Forms.Label.1", "lblGrip")
    With lblGrip
        .Caption = ""
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .MousePointer = fmMousePointerSizeNWSE
    End With

    ' Ensure form dimensions are properly applied before resizing controls
    DoEvents  ' Allow form to fully initialize

    ' Now resize controls to match the actual form dimensions
    ResizeControls
End Sub

' Dark: a green-tinted near-black window with a raised formula card, grey
' chips and a lime accent. Light: the same on a warm white with a deeper green.
Private Sub LoadTheme()
    If GetSetting("Breakdown", "FormulaTracing", "Theme", "Dark") = "Light" Then
        ColorWindow = &HF8FAFA          ' RGB(250, 250, 248)
        ColorCard = &HFFFFFF            ' RGB(255, 255, 255)
        ColorCardBorder = &HDFE3DF      ' RGB(223, 227, 223)
        ColorRule = &HE6E9E6            ' RGB(230, 233, 230)
        ColorChip = &HEAEEEA            ' RGB(234, 238, 234)
        ColorGreen = &H3E7A1F           ' RGB(31, 122, 62)
        ColorOnGreen = &HFFFFFF         ' RGB(255, 255, 255)
        ColorSelectedRow = &HE0F3E2     ' RGB(226, 243, 224)
        ColorRail = &HCDD2CC            ' RGB(204, 210, 205)
        ColorText = &H1B1D1A            ' RGB(26, 29, 27)
        ColorSecondary = &H5A6058       ' RGB(88, 96, 90)
        ColorCaption = &H787D76         ' RGB(118, 125, 120)
    Else
        ColorWindow = &H121412          ' RGB(18, 20, 18)
        ColorCard = &H1A1C19            ' RGB(25, 28, 26)
        ColorCardBorder = &H2D312C      ' RGB(44, 49, 45)
        ColorRule = &H252824            ' RGB(36, 40, 37)
        ColorChip = &H2A2D28            ' RGB(40, 45, 42)
        ColorGreen = &H70DA98           ' RGB(152, 218, 112)
        ColorOnGreen = &HE2210          ' RGB(16, 34, 14)
        ColorSelectedRow = &H1E2A1E     ' RGB(30, 42, 30)
        ColorRail = &H3E423C            ' RGB(60, 66, 62)
        ColorText = &HEBEEEA            ' RGB(234, 238, 235)
        ColorSecondary = &HA7ACA5       ' RGB(165, 172, 167)
        ColorCaption = &H757A73         ' RGB(115, 122, 117)
    End If
End Sub

' ---------------------------------------------------------------------------
' Control helpers
' ---------------------------------------------------------------------------

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

' Add a flat coloured rectangle used as a background, rule or tree line
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

' Add a cell chip: a rounded box under a transparent label with bold
' monospaced text. The box is made first so the text is drawn on top of it.
Private Function AddChip(ctlName As String) As MSForms.Label
    EnsureRoundBox ctlName
    Set AddChip = AddTextLabel(ctlName, "", CHIP_SIZE, True, ColorText)
    AddChip.TextAlign = fmTextAlignCenter
    AddChip.Font.Name = MonoFontName()
End Function

' Place a chip with its text centred in it, on a background of bgColor. Green
' chips mark the traced cell, the selected row and the highlighted reference.
Private Sub PlaceChip(lbl As MSForms.Label, ByVal x As Single, ByVal y As Single, ByVal w As Single, _
                      ByVal isGreen As Boolean, ByVal bgColor As Long)
    ' On the pixel grid, so the text is centred on the box as drawn
    w = SnapToPixel(x + w) - SnapToPixel(x)
    x = SnapToPixel(x)
    y = SnapToPixel(y)
    PlaceRoundBox lbl.Name, x, y, w, CHIP_HEIGHT, IIf(isGreen, ColorGreen, ColorChip), CHIP_RADIUS, bgColor
    lbl.ForeColor = IIf(isGreen, ColorOnGreen, ColorText)
    lbl.Move x + TEXT_NUDGE_X, TextTop(y, CHIP_HEIGHT, CHIP_SIZE), w, TextLineHeight(CHIP_SIZE)
    lbl.Visible = True
End Sub

Private Sub HideChip(ctlName As String)
    HideControl ctlName
    HideRoundBox ctlName
    Remember ctlName & "#token", ""
End Sub

' ---------------------------------------------------------------------------
' Rounded boxes. MSForms labels are always square, so a rounded box is eight
' controls (boxName_0 to boxName_7): two overlapping rectangles that leave the
' corners out (and, for a box with a border, two more one pixel in, in the
' fill colour), and a picture of each corner with its anti-aliased curve
' blended into the background.
' ---------------------------------------------------------------------------

' Create a rounded box's controls, hidden, if they don't exist yet. Call this
' before adding the controls that must be drawn on top of the box.
Private Sub EnsureRoundBox(boxName As String)
    Dim j As Long
    For j = 0 To BOX_PARTS - 1
        RoundBoxPart boxName, j
    Next j
End Sub

' Place a rounded box filled with fillColor on a background of bgColor, with a
' one-pixel border of borderColor unless that is -1 (NO_BORDER)
Private Sub PlaceRoundBox(boxName As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
                          ByVal fillColor As Long, ByVal radius As Single, ByVal bgColor As Long, _
                          Optional ByVal borderColor As Long = -1)
    Dim state As String
    Dim r As Long               ' Corner radius in pixels
    Dim rp As Single            ' ... and in points
    Dim outerColor As Long
    Dim j As Long

    ' On the pixel grid, so the rectangles and the corner pictures meet exactly
    w = SnapToPixel(x + w) - SnapToPixel(x)
    h = SnapToPixel(y + h) - SnapToPixel(y)
    x = SnapToPixel(x)
    y = SnapToPixel(y)
    r = Int(radius / PX_POINTS + 0.5)
    If r > Int(h / PX_POINTS / 2) Then r = Int(h / PX_POINTS / 2)
    If r > Int(w / PX_POINTS / 2) Then r = Int(w / PX_POINTS / 2)
    If r < 1 Then r = 1
    rp = r * PX_POINTS

    ' Nothing to do if the box already shows exactly this
    state = x & "|" & y & "|" & w & "|" & h & "|" & r & "|" & fillColor & "|" & bgColor & "|" & borderColor
    If ShownAs(boxName) = state Then Exit Sub

    ' Full width without the corner rows, and full height without the corner
    ' columns. A bordered box draws these in the border colour, and a second
    ' pair, one pixel in, in the fill colour.
    If borderColor = NO_BORDER Then outerColor = fillColor Else outerColor = borderColor
    PlacePart boxName, 0, x, y + rp, w, h - 2 * rp, outerColor
    PlacePart boxName, 1, x + rp, y, w - 2 * rp, h, outerColor
    If borderColor = NO_BORDER Then
        RoundBoxPart(boxName, 2).Visible = False
        RoundBoxPart(boxName, 3).Visible = False
    Else
        PlacePart boxName, 2, x + PX_POINTS, y + rp, w - 2 * PX_POINTS, h - 2 * rp, fillColor
        PlacePart boxName, 3, x + rp, y + PX_POINTS, w - 2 * rp, h - 2 * PX_POINTS, fillColor
    End If

    ' The corners: 0 and 1 along the top, 2 and 3 along the bottom
    For j = 0 To 3
        PlaceCorner boxName, j, IIf(j Mod 2 = 0, x, x + w - rp), IIf(j < 2, y, y + h - rp), r, _
                    fillColor, bgColor, borderColor
    Next j
    Remember boxName, state
End Sub

Private Sub PlacePart(boxName As String, ByVal part As Long, ByVal x As Single, ByVal y As Single, _
                      ByVal w As Single, ByVal h As Single, ByVal fillColor As Long)
    With RoundBoxPart(boxName, part)
        .BackColor = fillColor
        .Move x, y, w, h
        .Visible = True
    End With
End Sub

' One corner of a box, in the r x r pixel square whose top-left is (x, y)
Private Sub PlaceCorner(boxName As String, ByVal corner As Long, ByVal x As Single, ByVal y As Single, ByVal r As Long, _
                        ByVal fillColor As Long, ByVal bgColor As Long, ByVal borderColor As Long)
    Dim part As Object
    Dim side As Single
    Dim cx As Single
    Dim cy As Single
    Dim fontSize As Single
    Dim glyphWidth As Single
    Dim glyphHeight As Single

    side = r * PX_POINTS
    Set part = RoundBoxPart(boxName, 4 + corner)
    If UsePictures Then
        Set part.Picture = TraceUtils.CornerPicture(r, bgColor, borderColor, fillColor, corner)
        part.Move x, y, side, side
    Else
        ' A filled circle character of the corner's radius, centred on the
        ' square's inner corner (no border ring)
        If corner Mod 2 = 0 Then cx = x + side Else cx = x
        If corner < 2 Then cy = y + side Else cy = y
        fontSize = 2 * side / CIRCLE_INK
        MeasureCircle fontSize, glyphWidth, glyphHeight
        part.ForeColor = fillColor
        part.Font.Size = fontSize
        part.Move cx - glyphWidth / 2 + CIRCLE_SHIFT_X, cy - glyphHeight / 2 - CIRCLE_DROP * fontSize, _
                  glyphWidth, glyphHeight
    End If
    part.Visible = True
End Sub

Private Sub HideRoundBox(boxName As String)
    Dim j As Long
    If Len(ShownAs(boxName)) = 0 Then Exit Sub
    For j = 0 To BOX_PARTS - 1
        RoundBoxPart(boxName, j).Visible = False
    Next j
    Remember boxName, ""
End Sub

' The nearest whole device pixel, in points
Private Function SnapToPixel(ByVal v As Single) As Single
    SnapToPixel = PX_POINTS * Int(v / PX_POINTS + 0.5)
End Function

' Whether corner pictures can be made and shown here: try one on a temporary
' Image control. If not, corners fall back to circle characters.
Private Function CornerPicturesWork() As Boolean
    Dim pic As Object
    Dim probe As Object

    Set pic = TraceUtils.CornerPicture(4, ColorWindow, NO_BORDER, ColorChip, 0)
    If pic Is Nothing Then Exit Function

    On Error Resume Next
    Set probe = Me.Controls.Add("Forms.Image.1", "imgCornerProbe")
    If probe Is Nothing Then Exit Function
    Set probe.Picture = pic
    CornerPicturesWork = (Err.Number = 0)
    Me.Controls.Remove "imgCornerProbe"
End Function

' Size of an auto-sized label holding the corner circle at this font size,
' measured the first time it is needed
Private Sub MeasureCircle(ByVal fontSize As Single, glyphWidth As Single, glyphHeight As Single)
    Dim key As String
    Dim probe As MSForms.Label
    Dim measured As Variant

    key = "circle " & fontSize
    On Error Resume Next
    measured = CircleSizes(key)
    On Error GoTo 0
    If IsEmpty(measured) Then
        Set probe = Me.Controls.Add("Forms.Label.1", "lblCircleProbe")
        With probe
            .Move -1000, -1000
            .WordWrap = False
            .Font.Name = CIRCLE_FONT
            .Font.Size = fontSize
            .Caption = ChrW(&H25CF)
            .AutoSize = True
            measured = Array(.Width, .Height)
        End With
        Me.Controls.Remove "lblCircleProbe"
        CircleSizes.Add measured, key
    End If
    glyphWidth = measured(0)
    glyphHeight = measured(1)
End Sub

' What a box or token shows (its geometry and colour), or "" if hidden
Private Function ShownAs(itemName As String) As String
    On Error Resume Next
    ShownAs = Shown(itemName)
End Function

Private Sub Remember(itemName As String, state As String)
    On Error Resume Next
    Shown.Remove itemName
    On Error GoTo 0
    If Len(state) > 0 Then Shown.Add state, itemName
End Sub

' One part of a rounded box, created hidden the first time: 0 to 3 are
' rectangles, 4 to 7 the corners (pictures, or circle characters if pictures
' can't be used)
Private Function RoundBoxPart(boxName As String, ByVal part As Long) As Object
    Dim ctlName As String
    ctlName = boxName & "_" & part
    On Error Resume Next
    Set RoundBoxPart = Me.Controls(ctlName)
    On Error GoTo 0
    If Not RoundBoxPart Is Nothing Then Exit Function

    If part < 4 Then
        Set RoundBoxPart = AddPanel(ctlName, ColorWindow)
    ElseIf UsePictures Then
        Set RoundBoxPart = Me.Controls.Add("Forms.Image.1", ctlName)
        RoundBoxPart.BorderStyle = fmBorderStyleNone
        RoundBoxPart.SpecialEffect = fmSpecialEffectFlat
        RoundBoxPart.BackStyle = fmBackStyleTransparent
        RoundBoxPart.PictureSizeMode = fmPictureSizeModeStretch
        RoundBoxPart.PictureAlignment = fmPictureAlignmentTopLeft
    Else
        Set RoundBoxPart = AddTextLabel(ctlName, ChrW(&H25CF), CHIP_SIZE, False, ColorWindow)
        RoundBoxPart.Font.Name = CIRCLE_FONT
        RoundBoxPart.TextAlign = fmTextAlignCenter
    End If
    RoundBoxPart.Visible = False
    ' New controls go under the click layer
    If Not lblListHit Is Nothing Then lblListHit.ZOrder fmZOrderFront
End Function

' Top of a label whose text is centred vertically in a box
Private Function TextTop(ByVal boxTop As Single, ByVal boxHeight As Single, ByVal fontSize As Single) As Single
    TextTop = boxTop + (boxHeight - TextLineHeight(fontSize)) / 2 + TEXT_NUDGE_Y
End Function

Private Function TextLineHeight(ByVal fontSize As Single) As Single
    TextLineHeight = fontSize * LINE_HEIGHT_FACTOR
End Function

' Width of a chip for this text
Private Function ChipWidth(captionText As String) As Single
    ChipWidth = Len(captionText) * MonoCharWidth(CHIP_SIZE) + 2 * CHIP_PADDING
End Function

' A list label by name, created the first time it is needed
Private Function EnsureLabel(ctlName As String, kind As String) As MSForms.Label
    On Error Resume Next
    Set EnsureLabel = Me.Controls(ctlName)
    On Error GoTo 0
    If Not EnsureLabel Is Nothing Then Exit Function

    Select Case kind
        Case "chip"
            Set EnsureLabel = AddChip(ctlName)
        Case "rail"
            Set EnsureLabel = AddPanel(ctlName, ColorRail)
        Case "chevron"
            Set EnsureLabel = AddTextLabel(ctlName, "", TEXT_SIZE, False, ColorCaption)
            EnsureLabel.TextAlign = fmTextAlignCenter
            EnsureLabel.Font.Name = SystemFontName()  ' Has the chevron characters
        Case "formula"
            Set EnsureLabel = AddTextLabel(ctlName, "", TEXT_SIZE - 0.5, False, ColorSecondary)
            EnsureLabel.Font.Name = MonoFontName()
        Case "value"
            Set EnsureLabel = AddTextLabel(ctlName, "", TEXT_SIZE - 0.5, True, ColorText)
            EnsureLabel.Font.Name = MonoFontName()
            EnsureLabel.TextAlign = fmTextAlignRight
    End Select

    ' New list labels go under the click layer
    lblListHit.ZOrder fmZOrderFront
End Function

' Uber Move when it is installed, else the platform's UI font. Chosen once
' per window; the add-in cannot carry a font of its own.
Private Function UiFontName() As String
    If Len(UiFont) = 0 Then UiFont = PickUiFont()
    UiFontName = UiFont
End Function

' The first preferred font that is installed. A font that is not installed is
' drawn in the same fallback font as a made-up name, so it measures the same.
Private Function PickUiFont() As String
    Dim probe As MSForms.Label
    Dim fallbackWidth As Single
    Dim candidateWidth As Single
    Dim candidate As Variant

    PickUiFont = SystemFontName()
    On Error Resume Next
    Set probe = Me.Controls.Add("Forms.Label.1", "lblFontProbe")
    If probe Is Nothing Then Exit Function

    With probe
        .Move -1000, -1000
        .WordWrap = False
        .Caption = "Breakdown traces 0123456789 WMwm"
        .Font.Size = 24
        .Font.Name = "Breakdown Missing Font"
        .AutoSize = True
        fallbackWidth = .Width
        For Each candidate In Array("Uber Move Text", "Uber Move")
            candidateWidth = fallbackWidth
            .AutoSize = False
            .Font.Name = candidate
            .AutoSize = True
            candidateWidth = .Width
            If Abs(candidateWidth - fallbackWidth) > 0.5 Then
                PickUiFont = candidate
                Exit For
            End If
        Next candidate
    End With
    Me.Controls.Remove "lblFontProbe"
End Function

' Fonts that ship with the OS on each platform
Private Function SystemFontName() As String
#If Mac Then
    SystemFontName = "Helvetica Neue"
#Else
    SystemFontName = "Segoe UI"
#End If
End Function

Private Function MonoFontName() As String
#If Mac Then
    MonoFontName = "Menlo"
#Else
    MonoFontName = "Consolas"
#End If
End Function

' Advance width of one character of the monospaced font, in points
Private Function MonoCharWidth(fontSize As Single) As Single
#If Mac Then
    MonoCharWidth = fontSize * 0.602  ' Menlo
#Else
    MonoCharWidth = fontSize * 0.55   ' Consolas
#End If
End Function

Private Function KeyHintsText() As String
    Dim sep As String
    sep = "  " & ChrW(&HB7) & "  "
    KeyHintsText = ChrW(&H2191) & ChrW(&H2193) & " move" & sep & _
                   ChrW(&H2192) & ChrW(&H2190) & " expand" & sep & _
                   "Enter trace into" & sep & "Shift+Enter back"
End Function

Private Function MaxOf(a As Long, b As Long) As Long
    If a > b Then MaxOf = a Else MaxOf = b
End Function

' ---------------------------------------------------------------------------
' Keyboard and mouse
' ---------------------------------------------------------------------------

Private Sub txtKeys_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If HandleNavigationKey(KeyCode, Shift, True) Then KeyCode = 0
End Sub

' Typing while the list is active goes to the filter box
Private Sub txtKeys_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii >= 32 Then
        txtFilter.SetFocus
        txtFilter.text = txtFilter.text & Chr$(KeyAscii)
        txtFilter.SelStart = Len(txtFilter.text)
    End If
    KeyAscii = 0
End Sub

' In the filter box, Left/Right move the text cursor; the other keys work the list
Private Sub txtFilter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If HandleNavigationKey(KeyCode, Shift, False) Then KeyCode = 0
End Sub

' Filtering shows matching rows and the rows above them; the first match is
' selected and its formula shown, without moving Excel on every keystroke
Private Sub txtFilter_Change()
    Me.Controls("lblFilterHint").Visible = (Len(txtFilter.text) = 0)
    TopRow = 0
    RefreshTree
    SelectedRow = FirstMatchingRow()
    RenderRows
    UpdateStatus
    If VisibleCount > 0 Then ShowFormulaContext VisibleNodes(SelectedRow + 1)
End Sub

' Returns True if the key was used
Private Function HandleNavigationKey(ByVal keyValue As Integer, ByVal Shift As Integer, allowExpandKeys As Boolean) As Boolean
    Dim nodeIndex As Integer
    HandleNavigationKey = True

    Select Case keyValue
        Case vbKeyUp
            MoveSelection -1
        Case vbKeyDown
            MoveSelection 1
        Case vbKeyPageUp
            MoveSelection -SlotCount
        Case vbKeyPageDown
            MoveSelection SlotCount
        Case vbKeyRight
            If allowExpandKeys Then SetExpanded SelectedNodeIndex(), True Else HandleNavigationKey = False
        Case vbKeyLeft
            If allowExpandKeys Then SetExpanded SelectedNodeIndex(), False Else HandleNavigationKey = False
        Case vbKeyEscape
            ' Esc clears the filter first, then closes the window
            If Len(txtFilter.text) > 0 Then
                txtFilter.text = ""
                txtKeys.SetFocus
            Else
                Unload Me
            End If
        Case vbKeyReturn
            If (Shift And SHIFT_KEY_MASK) <> 0 Then
                ' Shift+Enter: go back up to the trace we came from
                TraceBack
            Else
                ' Enter: go to the cell and open a new trace rooted at it, so
                ' the user can drill deeper one level at a time. This window
                ' is kept, hidden, so Shift+Enter can bring it back as it is.
                nodeIndex = SelectedNodeIndex()
                If nodeIndex > 0 Then
                    If NavigateToCell(NodeAddress(nodeIndex)) Then
                        RememberWindowSize
                        TraceHistory.Add Me
                        TraceUtils.ShowTracePrecedents TraceHistory
                        Me.Hide
                    End If
                End If
            End If
        Case Else
            HandleNavigationKey = False
    End Select
End Function

' Click on the list: the chevron expands or collapses, anywhere else selects
Private Sub lblListHit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim listRow As Long
    Dim nodeIndex As Integer
    Dim chevronLeft As Single

    listRow = TopRow + Int(Y / ROW_HEIGHT)
    If listRow < 0 Or listRow >= VisibleCount Then Exit Sub
    nodeIndex = VisibleNodes(listRow + 1)

    chevronLeft = ChipLeft(TreeData(nodeIndex).level) + ChipWidth(ShortAddress(TreeData(nodeIndex).text)) + 3 - lblListHit.Left
    If TreeData(nodeIndex).hasChildren And X >= chevronLeft And X <= chevronLeft + CHEVRON_WIDTH + 4 Then
        SelectRow listRow, False
        SetExpanded nodeIndex, Not TreeData(nodeIndex).IsExpanded
    Else
        SelectRow listRow, True
    End If

    On Error Resume Next
    txtKeys.SetFocus
End Sub

Private Sub scrList_Change()
    If IsRendering Then Exit Sub
    TopRow = scrList.Value
    RenderRows
End Sub

Private Sub scrList_Scroll()
    scrList_Change
End Sub

' ---------------------------------------------------------------------------
' Selection
' ---------------------------------------------------------------------------

' Tree node shown on the selected list row, or 0 if none
Private Function SelectedNodeIndex() As Integer
    If SelectedRow >= 0 And SelectedRow < VisibleCount Then
        SelectedNodeIndex = VisibleNodes(SelectedRow + 1)
    End If
End Function

Private Sub MoveSelection(ByVal delta As Long)
    If VisibleCount = 0 Then Exit Sub
    Dim target As Long
    target = SelectedRow + delta
    If target < 0 Then target = 0
    If target > VisibleCount - 1 Then target = VisibleCount - 1
    If target <> SelectedRow Then SelectRow target, True
End Sub

' Select a list row, scroll it into view and optionally go to its cell
Private Sub SelectRow(ByVal listRow As Long, ByVal goToCell As Boolean)
    Dim previousRow As Long
    Dim previousTop As Long
    previousRow = SelectedRow
    previousTop = TopRow

    SelectedRow = listRow
    If SelectedRow < TopRow Then TopRow = SelectedRow
    If SlotCount > 0 And SelectedRow >= TopRow + SlotCount Then TopRow = SelectedRow - SlotCount + 1

    If TopRow = previousTop Then
        ' No scrolling: only the rows losing and gaining the selection change
        RenderSlotOf previousRow
        RenderSlotOf SelectedRow
        PlaceSelectionHighlight
    Else
        RenderRows
    End If
    If goToCell Then SelectNode VisibleNodes(listRow + 1)
End Sub

' Row selected: go to the cell, with the formula card showing the formula
' that links it into the tree
Private Sub SelectNode(ByVal nodeIndex As Integer)
    If NavigateToCell(NodeAddress(nodeIndex)) Then ShowFormulaContext nodeIndex
End Sub

' First list row matching the filter (the traced cell when there is no filter)
Private Function FirstMatchingRow() As Long
    Dim i As Long
    Dim filterText As String
    filterText = LCase$(Trim$(txtFilter.text))
    If Len(filterText) = 0 Then Exit Function
    For i = 1 To VisibleCount
        If NodeMatches(VisibleNodes(i), filterText) Then
            FirstMatchingRow = i - 1
            Exit Function
        End If
    Next i
End Function

Private Sub SetExpanded(ByVal nodeIndex As Integer, ByVal expanded As Boolean)
    If nodeIndex = 0 Then Exit Sub
    If Not TreeData(nodeIndex).hasChildren Then Exit Sub
    If TreeData(nodeIndex).IsExpanded = expanded Then Exit Sub
    If expanded And TreeData(nodeIndex).copyPending Then
        If Not CopyFirstAppearance(nodeIndex) Then Exit Sub
    End If
    TreeData(nodeIndex).IsExpanded = expanded
    RefreshTree
End Sub

' Opening a repeat row for the first time: copy in, under it, the rows under
' its cell's first appearance, moved to its level. The copies count as
' repeats (they add no new precedents), and repeat rows among them are copied
' in the same way when opened. False if there is nothing to copy.
Private Function CopyFirstAppearance(ByVal nodeIndex As Integer) As Boolean
    Dim cellAddress As String
    Dim source As Integer
    Dim sourceEnd As Integer
    Dim copyCount As Integer
    Dim levelShift As Integer
    Dim copies() As TreeNode
    Dim i As Integer

    cellAddress = NodeAddress(nodeIndex)
    For i = 1 To TreeCount
        If Not TreeData(i).isRepeat And TreeData(i).hasChildren Then
            If NodeAddress(i) = cellAddress Then
                source = i
                Exit For
            End If
        End If
    Next i
    If source = 0 Then Exit Function

    ' The first appearance's rows: those after it that are deeper
    sourceEnd = source
    Do While sourceEnd < TreeCount
        If TreeData(sourceEnd + 1).level <= TreeData(source).level Then Exit Do
        sourceEnd = sourceEnd + 1
    Loop
    copyCount = sourceEnd - source
    If copyCount = 0 Or CLng(TreeCount) + copyCount > 32000 Then Exit Function

    ' Take the copies first (the first appearance may sit among the rows that
    ' move), then make room for them after the repeat row
    levelShift = TreeData(nodeIndex).level - TreeData(source).level
    ReDim copies(1 To copyCount)
    For i = 1 To copyCount
        copies(i) = TreeData(source + i)
        copies(i).level = copies(i).level + levelShift
        copies(i).isRepeat = True
        copies(i).IsExpanded = False
    Next i
    If TreeCount + copyCount > UBound(TreeData) Then ReDim Preserve TreeData(1 To TreeCount + copyCount + 50)
    For i = TreeCount To nodeIndex + 1 Step -1
        TreeData(i + copyCount) = TreeData(i)
    Next i
    For i = 1 To copyCount
        TreeData(nodeIndex + i) = copies(i)
    Next i
    TreeCount = TreeCount + copyCount
    TreeData(nodeIndex).copyPending = False
    CopyFirstAppearance = True
End Function

' Shift+Enter: bring back the window this one was drilled into from, exactly
' as it was left. No re-trace, so it is instant.
Private Sub TraceBack()
    If TraceHistory.Count = 0 Then Exit Sub

    Dim previousWindow As frmPrecedents
    Set previousWindow = TraceHistory(TraceHistory.Count)
    TraceHistory.Remove TraceHistory.Count

    previousWindow.ReturnHere
    SwitchingWindows = True
    Unload Me
End Sub

' Show this hidden window again and select the cell its highlighted row is on
Public Sub ReturnHere()
    Dim nodeIndex As Integer
    nodeIndex = SelectedNodeIndex()
    If nodeIndex > 0 Then
        SelectNode nodeIndex
    ElseIf Not TracedCell Is Nothing Then
        On Error Resume Next
        TracedCell.Worksheet.Activate
        TracedCell.Select
        On Error GoTo 0
    End If
    Me.Show vbModeless
End Sub

' Unload a hidden history window without closing the rest of its chain
Public Sub CloseQuietly()
    SwitchingWindows = True
    Unload Me
End Sub

' ---------------------------------------------------------------------------
' Tree
' ---------------------------------------------------------------------------

' Parse existing ListBox indented structure into tree nodes
Private Sub ParseListBoxToTree()
    TreeCount = 0

    Dim i As Integer
    For i = 0 To lstPrecedents.ListCount - 1
        Dim addressText As String
        addressText = lstPrecedents.List(i, 0)

        ' Extract level from indentation
        Dim level As Integer
        level = GetIndentationLevel(addressText)

        ' Extract clean text (remove indentation and level markers)
        Dim cleanText As String
        cleanText = GetCleanText(addressText)

        ' Check if this node has children (next item has higher level)
        Dim hasChildren As Boolean
        hasChildren = False
        If i < lstPrecedents.ListCount - 1 Then
            Dim nextLevel As Integer
            nextLevel = GetIndentationLevel(lstPrecedents.List(i + 1, 0))
            hasChildren = (nextLevel > level)
        End If

        ' Store original data for this node
        Dim originalData As String
        originalData = lstPrecedents.List(i, 0) & "|" & _
                      lstPrecedents.List(i, 1) & "|" & _
                      lstPrecedents.List(i, 2)

        ' The traced cell starts expanded to show the first level immediately
        AddTreeNode cleanText, level, (level = 0), hasChildren, originalData
    Next i

    UniqueCount = 0
    For i = 2 To TreeCount
        If Not TreeData(i).isRepeat Then UniqueCount = UniqueCount + 1
    Next i

    ' A repeat row opens like its cell's first appearance: its rows are copied
    ' in the first time it is expanded (see CopyFirstAppearance)
    Dim expandable As New Collection
    Dim canExpand As Boolean
    For i = 2 To TreeCount
        If Not TreeData(i).isRepeat And TreeData(i).hasChildren Then
            On Error Resume Next
            expandable.Add True, NodeAddress(i)
            On Error GoTo 0
        End If
    Next i
    For i = 2 To TreeCount
        If TreeData(i).isRepeat Then
            canExpand = False
            On Error Resume Next
            canExpand = expandable(NodeAddress(i))
            On Error GoTo 0
            TreeData(i).hasChildren = canExpand
            TreeData(i).copyPending = canExpand
        End If
    Next i
End Sub

' Add a tree node
Private Sub AddTreeNode(nodeText As String, level As Integer, expanded As Boolean, hasChildren As Boolean, originalData As String)
    ' Expand array if needed
    If TreeCount >= UBound(TreeData) Then
        ReDim Preserve TreeData(1 To UBound(TreeData) + 50)
    End If

    TreeCount = TreeCount + 1
    With TreeData(TreeCount)
        .text = nodeText
        .level = level
        .IsExpanded = expanded
        .hasChildren = hasChildren
        .IsVisible = True
        .originalData = originalData
        .refIndex = TraceUtils.TraceRowRef(nodeText)
        .isRepeat = TraceUtils.TraceRowIsRepeat(nodeText)
    End With
End Sub

' Rebuild the visible rows (expand state, or the filter while one is typed)
' and their tree lines, then redraw
Private Sub RefreshTree()
    Dim i As Integer
    UpdateVisibility

    ReDim VisibleNodes(1 To TreeCount + 1)
    VisibleCount = 0
    For i = 1 To TreeCount
        If TreeData(i).IsVisible Then
            VisibleCount = VisibleCount + 1
            VisibleNodes(VisibleCount) = i
        End If
    Next i
    ComputeRails

    If SelectedRow > VisibleCount - 1 Then SelectedRow = VisibleCount - 1
    If SelectedRow < 0 And VisibleCount > 0 Then SelectedRow = 0
    RenderRows
End Sub

' Update visibility from the expand state, or from the filter while one is typed
Private Sub UpdateVisibility()
    Dim i As Integer, j As Integer
    Dim parentLevel As Integer
    Dim filterText As String
    If Not txtFilter Is Nothing Then filterText = LCase$(Trim$(txtFilter.text))

    If Len(filterText) > 0 Then
        ' Matching rows and every row above them in the tree
        For i = 1 To TreeCount
            TreeData(i).IsVisible = False
        Next i
        For i = 1 To TreeCount
            If NodeMatches(i, filterText) Then
                j = i
                Do While j > 0
                    TreeData(j).IsVisible = True
                    j = ParentNode(j)
                Loop
            End If
        Next i
        Exit Sub
    End If

    ' First, make all root nodes (level 0) visible
    For i = 1 To TreeCount
        TreeData(i).IsVisible = (TreeData(i).level = 0)
    Next i

    ' Then, make child nodes visible if their parent is expanded
    For i = 1 To TreeCount
        If TreeData(i).level > 0 Then
            parentLevel = TreeData(i).level - 1

            ' Look backwards for the parent
            For j = i - 1 To 1 Step -1
                If TreeData(j).level = parentLevel Then
                    If TreeData(j).IsVisible And TreeData(j).IsExpanded Then
                        TreeData(i).IsVisible = True
                    End If
                    Exit For
                ElseIf TreeData(j).level < parentLevel Then
                    Exit For
                End If
            Next j
        End If
    Next i
End Sub

' True if the cell name or formula of a node contains the filter text
Private Function NodeMatches(ByVal nodeIndex As Integer, filterText As String) As Boolean
    If InStr(1, LCase$(ShortAddress(TreeData(nodeIndex).text)), filterText, vbBinaryCompare) > 0 Then
        NodeMatches = True
    Else
        NodeMatches = InStr(1, LCase$(OneLine(NodeColumn(nodeIndex, 2))), filterText, vbBinaryCompare) > 0
    End If
End Function

' Tree lines per visible row: one character per level, "1" where the line of
' that level continues below this row (the node at that level has a later
' sibling in the list)
Private Sub ComputeRails()
    Dim i As Long
    Dim j As Long
    Dim level As Integer
    Dim hasNext As Boolean
    Dim parentRails As String

    ReDim VisibleRails(1 To VisibleCount + 1)
    For i = 1 To VisibleCount
        level = TreeData(VisibleNodes(i)).level
        hasNext = False
        For j = i + 1 To VisibleCount
            If TreeData(VisibleNodes(j)).level < level Then Exit For
            If TreeData(VisibleNodes(j)).level = level Then
                hasNext = True
                Exit For
            End If
        Next j

        parentRails = ""
        For j = i - 1 To 1 Step -1
            If TreeData(VisibleNodes(j)).level < level Then
                parentRails = VisibleRails(j)
                Exit For
            End If
        Next j
        If level = 0 Then
            VisibleRails(i) = ""
        Else
            VisibleRails(i) = Left$(parentRails & String(level, "0"), level - 1) & IIf(hasNext, "1", "0")
        End If
    Next i
End Sub

' Get indentation level from text
Private Function GetIndentationLevel(text As String) As Integer
    ' Count leading spaces and convert to level
    Dim spaceCount As Integer
    Dim i As Integer

    For i = 1 To Len(text)
        If Mid(text, i, 1) = " " Then
            spaceCount = spaceCount + 1
        Else
            Exit For
        End If
    Next i

    ' Convert spaces to level (assuming 2 spaces per level from original format)
    GetIndentationLevel = spaceCount \ 2
End Function

' Get clean text without indentation and level markers
Private Function GetCleanText(text As String) As String
    Dim cleanText As String
    cleanText = text

    ' Remove leading spaces
    cleanText = LTrim(cleanText)

    ' Remove level markers like "L1: ", "L2: ", etc.
    Dim colonPos As Long
    colonPos = InStr(cleanText, ": ")
    If colonPos > 0 Then
        cleanText = Mid(cleanText, colonPos + 2)
    End If

    GetCleanText = cleanText
End Function

' A formula on one line, as the list shows it: line breaks and tabs become
' spaces, and runs of spaces one space
Private Function OneLine(ByVal formulaText As String) As String
    formulaText = Replace(Replace(Replace(formulaText, vbCr, " "), vbLf, " "), vbTab, " ")
    Do While InStr(formulaText, "  ") > 0
        formulaText = Replace(formulaText, "  ", " ")
    Loop
    OneLine = formulaText
End Function

' Column of a node's original row: 1 = value, 2 = formula
Private Function NodeColumn(ByVal nodeIndex As Integer, ByVal column As Integer) As String
    Dim dataParts() As String
    dataParts = Split(TreeData(nodeIndex).originalData, "|")
    If UBound(dataParts) >= column Then NodeColumn = dataParts(column)
End Function

' Nearest row above this one with a lower level, or 0 for the traced cell
Private Function ParentNode(ByVal nodeIndex As Integer) As Integer
    Dim j As Integer
    For j = nodeIndex - 1 To 1 Step -1
        If TreeData(j).level < TreeData(nodeIndex).level Then
            ParentNode = j
            Exit Function
        End If
    Next j
End Function

' Cell address of a tree node, without the notes after it
Private Function NodeAddress(ByVal nodeIndex As Integer) As String
    NodeAddress = TraceUtils.TraceRowAddress(TreeData(nodeIndex).text)
End Function

' The cell or range a tree row stands for, or Nothing if it can't be resolved
Private Function NodeRange(ByVal nodeIndex As Integer) As Range
    On Error Resume Next
    If nodeIndex = 1 Then
        Set NodeRange = TracedCell
    Else
        Set NodeRange = Application.Range(NodeAddress(nodeIndex))
    End If
End Function

' True for a single cell that holds a formula
Private Function NodeHasFormula(ByVal nodeIndex As Integer) As Boolean
    On Error Resume Next
    Dim target As Range
    Set target = NodeRange(nodeIndex)
    If target Is Nothing Then Exit Function
    If target.Cells.Count = 1 Then NodeHasFormula = target.HasFormula
End Function

' Split "'[Book.xlsx]Sheet 1'!$A$1", "Sheet1!$A$1" or "$A$1" into its parts.
' The cell reference loses its $ signs.
Private Sub SplitAddress(ByVal addr As String, bookName As String, sheetName As String, cellRef As String)
    Dim bangPos As Long
    Dim closePos As Long
    Dim sheetPart As String

    bookName = ""
    sheetName = ""
    bangPos = InStrRev(addr, "!")
    If bangPos = 0 Then
        cellRef = Replace(addr, "$", "")
        Exit Sub
    End If

    cellRef = Replace(Mid(addr, bangPos + 1), "$", "")
    sheetPart = Left(addr, bangPos - 1)
    If Len(sheetPart) >= 2 And Left(sheetPart, 1) = "'" And Right(sheetPart, 1) = "'" Then
        sheetPart = Replace(Mid(sheetPart, 2, Len(sheetPart) - 2), "''", "'")
    End If

    If Left(sheetPart, 1) = "[" Then
        closePos = InStr(sheetPart, "]")
        If closePos > 0 Then
            bookName = Mid(sheetPart, 2, closePos - 2)
            sheetPart = Mid(sheetPart, closePos + 1)
        End If
    End If
    sheetName = sheetPart
End Sub

' Address as shown in the list: just the cell on the traced sheet, with the
' sheet on another sheet, and with workbook and sheet in another workbook
Private Function ShortAddress(ByVal nodeText As String) As String
    Dim bookName As String
    Dim sheetName As String
    Dim cellRef As String

    SplitAddress TraceUtils.TraceRowAddress(nodeText), bookName, sheetName, cellRef

    Dim sameBook As Boolean
    sameBook = (Len(bookName) = 0 Or bookName = RootBook)
    If Len(sheetName) = 0 Or (sameBook And sheetName = RootSheet) Then
        ShortAddress = cellRef
    ElseIf sameBook Then
        ShortAddress = sheetName & "!" & cellRef
    Else
        ShortAddress = "[" & bookName & "]" & sheetName & "!" & cellRef
    End If
End Function

' Navigate to a cell address. Returns True if the cell was selected.
Private Function NavigateToCell(precedentAddress As String) As Boolean
    On Error Resume Next

    Dim target As Range

    ' Parse out sheet name and cell address
    Dim exclamationPosition As Integer
    exclamationPosition = InStr(precedentAddress, "!")

    If exclamationPosition > 0 Then
        Dim sheetName As String
        Dim cellAddress As String

        ' Handle external references like [Workbook]Sheet!Address
        If InStr(precedentAddress, "[") > 0 Then
            ' External reference format: [Workbook]Sheet!Address
            sheetName = Mid(precedentAddress, InStrRev(precedentAddress, "]") + 1, exclamationPosition - InStrRev(precedentAddress, "]") - 1)
        Else
            ' Simple format: Sheet!Address
            sheetName = Left(precedentAddress, exclamationPosition - 1)
        End If

        cellAddress = Mid(precedentAddress, exclamationPosition + 1)

        ' Clean up sheet name
        If Right(sheetName, 1) = "'" Then
            sheetName = Left(sheetName, Len(sheetName) - 1)
        End If
        If Left(sheetName, 1) = "'" Then
            sheetName = Mid(sheetName, 2)
        End If

        ' Navigate to the cell
        Worksheets(sheetName).Activate
        Set target = Worksheets(sheetName).Range(cellAddress)
    Else
        ' No sheet specified - assume current sheet
        Set target = Range(precedentAddress)
    End If
    target.Select

    NavigateToCell = (Err.Number = 0)
    On Error GoTo 0
End Function

' ---------------------------------------------------------------------------
' Formula card
' ---------------------------------------------------------------------------

' Show in the formula card the formula that relates this row to the tree,
' with the connecting reference as a green chip. Precedents: the parent's
' formula with the selected cell highlighted (the traced cell shows its own
' formula). Dependents: the selected cell's formula with its parent highlighted.
Private Sub ShowFormulaContext(ByVal nodeIndex As Integer)
    On Error Resume Next

    Dim formulaNode As Integer
    Dim relationNode As Integer
    formulaNode = nodeIndex
    If FORMULA_FROM_ANCESTOR Then
        Dim ancestor As Integer
        ancestor = ParentNode(nodeIndex)
        Do While ancestor > 0
            If NodeHasFormula(ancestor) Then Exit Do
            ancestor = ParentNode(ancestor)
        Loop
        If ancestor > 0 Then
            formulaNode = ancestor
            relationNode = nodeIndex
        End If
    Else
        relationNode = ParentNode(nodeIndex)
    End If

    Dim formulaCell As Range
    Set formulaCell = NodeRange(formulaNode)
    If formulaCell Is Nothing Then Exit Sub
    Set formulaCell = formulaCell.Cells(1, 1)

    CardFormula = formulaCell.formula
    CardHighlightStart = 0
    Me.Controls("lblCardCaption").Caption = "FORMULA  " & ChrW(&HB7) & "  " & formulaCell.address(False, False)
    Me.Controls("lblCardValue").Caption = "= " & TraceUtils.GetCellValueAsString(formulaCell)

    If relationNode > 0 Then
        ' A row for a reference in the formula shown highlights that very
        ' reference, so each use of a cell used twice can be highlighted
        If TreeData(relationNode).refIndex > 0 And ParentNode(relationNode) = formulaNode Then
            Dim spans As Collection
            Dim span As Variant
            Set spans = TraceUtils.ReferenceSpans(CardFormula)
            If TreeData(relationNode).refIndex <= spans.Count Then
                span = spans(TreeData(relationNode).refIndex)
                CardHighlightStart = span(0)
            End If
        End If

        ' Otherwise the first reference to the cell, or to a range holding it
        If CardHighlightStart = 0 Then
            Dim relationCell As Range
            Dim spanStart As Long
            Dim spanLength As Long
            Set relationCell = NodeRange(relationNode)
            If Not relationCell Is Nothing Then
                If TraceUtils.FindReferenceSpan(CardFormula, formulaCell.Worksheet, relationCell, spanStart, spanLength) Then
                    CardHighlightStart = spanStart
                End If
            End If
        End If
    End If

    ' The card grows with the formula; re-lay out the window if its height changed
    Dim previousLines As Long
    previousLines = CardLines
    RenderFormulaCard
    If CardLines <> previousLines Then ResizeControls
End Sub

' Lay out the card formula as a flow of words, with each cell reference as a
' chip. Wraps at operators and commas; after MAX_FORMULA_LINES it ends in "...".
Private Sub RenderFormulaCard()
    Dim areaLeft As Single
    Dim areaRight As Single
    Dim areaTop As Single
    Dim x As Single
    Dim lineIndex As Long
    Dim used As Long
    Dim tokenText As String
    Dim isChip As Boolean
    Dim tokenWidth As Single
    Dim i As Long
    Dim wordStart As Long
    Dim nextRef As Long
    Dim refLength As Long
    Dim spans As Collection
    Dim spanIndex As Long
    Dim ellipsisWidth As Single
    Dim afterSpace As Boolean

    areaLeft = PADDING + CARD_PADDING
    areaRight = CardRight() - CARD_PADDING
    areaTop = CardTop() + CARD_PADDING + CARD_CAPTION_HEIGHT
    ellipsisWidth = 3 * MonoCharWidth(TEXT_SIZE)
    x = areaLeft

    Set spans = TraceUtils.ReferenceSpans(CardFormula)
    spanIndex = 1
    i = 1
    Do While i <= Len(CardFormula)
        ' Next token: a reference chip, or a word of plain text
        If spanIndex <= spans.Count Then nextRef = spans(spanIndex)(0) Else nextRef = Len(CardFormula) + 1
        If i = nextRef Then
            refLength = spans(spanIndex)(1)
            tokenText = Mid$(CardFormula, i, refLength)
            isChip = True
            spanIndex = spanIndex + 1
        Else
            wordStart = i
            Do While i < nextRef And i <= Len(CardFormula)
                i = i + 1
                If InStr(1, FORMULA_BREAK_CHARS, Mid$(CardFormula, i - 1, 1)) > 0 Then Exit Do
            Loop
            ' Line breaks and tabs show as spaces, and a run of spaces as one
            tokenText = Replace(Replace(Replace(Mid$(CardFormula, wordStart, i - wordStart), vbCr, " "), vbLf, " "), vbTab, " ")
            If Len(Trim$(tokenText)) = 0 And (x = areaLeft Or afterSpace) Then GoTo NextToken
            isChip = False
        End If

        If isChip Then tokenWidth = ChipWidth(tokenText) Else tokenWidth = Len(tokenText) * MonoCharWidth(TEXT_SIZE)
        afterSpace = (Not isChip And Right$(tokenText, 1) = " ")

        ' Wrap; past the last line, finish with an ellipsis
        If x + tokenWidth > areaRight And x > areaLeft Then
            lineIndex = lineIndex + 1
            x = areaLeft
        End If
        If lineIndex >= MAX_FORMULA_LINES Then
            lineIndex = MAX_FORMULA_LINES - 1
            used = used + 1
            PlaceToken used, "...", False, False, areaRight - ellipsisWidth, areaTop + lineIndex * FORMULA_LINE_HEIGHT, ellipsisWidth
            Exit Do
        End If

        used = used + 1
        PlaceToken used, tokenText, isChip, (isChip And i = CardHighlightStart), x, areaTop + lineIndex * FORMULA_LINE_HEIGHT, tokenWidth
        If isChip Then
            x = x + tokenWidth + 2
            i = i + refLength
        Else
            x = x + tokenWidth
        End If
NextToken:
    Loop

    ' Hide tokens left over from a longer formula
    For i = used + 1 To TokenCount
        If Len(ShownAs("lblTok" & i & "#token")) > 0 Then HideChip "lblTok" & i
    Next i
    CardLines = lineIndex + 1
End Sub

Private Sub PlaceToken(ByVal tokenNumber As Long, tokenText As String, ByVal isChip As Boolean, ByVal isGreen As Boolean, _
                       ByVal x As Single, ByVal y As Single, ByVal w As Single)
    ' Every token can be a chip, so each has its rounded box under it
    Dim lbl As MSForms.Label
    Dim state As String
    state = tokenText & "|" & isChip & "|" & isGreen & "|" & x & "|" & y & "|" & w
    If ShownAs("lblTok" & tokenNumber & "#token") = state Then Exit Sub

    If tokenNumber > TokenCount Then
        Set lbl = AddChip("lblTok" & tokenNumber)
        TokenCount = tokenNumber
    Else
        Set lbl = Me.Controls("lblTok" & tokenNumber)
    End If

    lbl.Caption = tokenText
    If isChip Then
        lbl.TextAlign = fmTextAlignCenter
        lbl.Font.Size = CHIP_SIZE
        lbl.Font.Bold = True
        PlaceChip lbl, x, y + (FORMULA_LINE_HEIGHT - CHIP_HEIGHT) / 2, w, isGreen, ColorCard
    Else
        HideRoundBox lbl.Name
        lbl.TextAlign = fmTextAlignLeft
        lbl.Font.Size = TEXT_SIZE
        lbl.Font.Bold = False
        lbl.ForeColor = ColorText
        lbl.Move x, TextTop(y, FORMULA_LINE_HEIGHT, TEXT_SIZE), w + 2, TextLineHeight(TEXT_SIZE)
        lbl.Visible = True
    End If
    Remember "lblTok" & tokenNumber & "#token", state
End Sub

' ---------------------------------------------------------------------------
' List rows
' ---------------------------------------------------------------------------

' Fill the row slots from VisibleNodes, starting at TopRow
Private Sub RenderRows()
    Dim s As Long
    Dim listRow As Long

    If SlotCount = 0 Then Exit Sub
    If TopRow > VisibleCount - SlotCount Then TopRow = VisibleCount - SlotCount
    If TopRow < 0 Then TopRow = 0

    For s = 0 To SlotCount - 1
        listRow = TopRow + s
        If listRow < VisibleCount Then
            RenderRow s, VisibleNodes(listRow + 1), VisibleRails(listRow + 1), ListTop + s * ROW_HEIGHT, (listRow = SelectedRow)
        Else
            HideRow s
        End If
    Next s
    PlaceSelectionHighlight

    ' Scrollbar, without reacting to our own change
    IsRendering = True
    scrList.Max = MaxOf(0, VisibleCount - SlotCount)
    scrList.LargeChange = MaxOf(1, SlotCount - 1)
    scrList.Value = TopRow
    scrList.Visible = (VisibleCount > SlotCount)
    IsRendering = False
End Sub

' The one rounded highlight behind the selected row, if it is on screen
Private Sub PlaceSelectionHighlight()
    If SelectedRow >= TopRow And SelectedRow < TopRow + SlotCount And SelectedRow < VisibleCount Then
        PlaceRoundBox "lblSelection", PADDING - 4, ListTop + (SelectedRow - TopRow) * ROW_HEIGHT + 1, _
                      ListWidth + 8, ROW_HEIGHT - 2, ColorSelectedRow, ROW_RADIUS, ColorWindow
    Else
        HideRoundBox "lblSelection"
    End If
End Sub

' Redraw the slot that shows list row listRow, if it is on screen
Private Sub RenderSlotOf(ByVal listRow As Long)
    If listRow < 0 Or listRow >= VisibleCount Then Exit Sub
    If listRow < TopRow Or listRow >= TopRow + SlotCount Then Exit Sub
    RenderRow listRow - TopRow, VisibleNodes(listRow + 1), VisibleRails(listRow + 1), _
              ListTop + (listRow - TopRow) * ROW_HEIGHT, (listRow = SelectedRow)
End Sub

Private Sub RenderRow(ByVal s As Long, ByVal nodeIndex As Integer, ByVal rails As String, ByVal rowTop As Single, ByVal isSelected As Boolean)
    Dim level As Integer
    Dim shortName As String
    Dim chipLeftX As Single
    Dim chipW As Single
    Dim midY As Single
    Dim formulaLeft As Single
    Dim valueLeft As Single
    Dim a As Integer
    Dim rail As MSForms.Label
    Dim chip As MSForms.Label

    level = TreeData(nodeIndex).level
    shortName = ShortAddress(TreeData(nodeIndex).text)
    chipLeftX = ChipLeft(level)
    chipW = ChipWidth(shortName)
    midY = rowTop + ROW_HEIGHT / 2
    valueLeft = PADDING + ListWidth - VALUE_COLUMN_WIDTH

    ' Tree lines: full-height lines for ancestors whose branch continues, and
    ' an elbow into this row's chip
    For a = 1 To level - 1
        If Mid$(rails, a, 1) = "1" Then
            Set rail = EnsureLabel("lblRail" & s & "_" & a, "rail")
            If a > MaxRailDepth Then MaxRailDepth = a
            rail.Move RailX(a), rowTop, 1, ROW_HEIGHT
            rail.Visible = True
        Else
            HideControl "lblRail" & s & "_" & a
        End If
    Next a
    For a = IIf(level < 1, 1, level) To MaxRailDepth
        HideControl "lblRail" & s & "_" & a
    Next a
    If level >= 1 Then
        With EnsureLabel("lblElbowV" & s, "rail")
            .Move RailX(level), rowTop, 1, IIf(Mid$(rails, level, 1) = "1", ROW_HEIGHT, ROW_HEIGHT / 2)
            .Visible = True
        End With
        With EnsureLabel("lblElbowH" & s, "rail")
            .Move RailX(level), midY, chipLeftX - RailX(level) - 1, 1
            .Visible = True
        End With
    Else
        HideControl "lblElbowV" & s
        HideControl "lblElbowH" & s
    End If

    ' Cell chip: green for the traced cell and the selected row
    Set chip = EnsureLabel("lblChip" & s, "chip")
    chip.Caption = shortName
    PlaceChip chip, chipLeftX, midY - CHIP_HEIGHT / 2, chipW, (nodeIndex = 1 Or isSelected), _
              IIf(isSelected, ColorSelectedRow, ColorWindow)

    ' Expand/collapse chevron after the chip
    With EnsureLabel("lblChev" & s, "chevron")
        If TreeData(nodeIndex).hasChildren Then
            .Caption = IIf(TreeData(nodeIndex).IsExpanded, ChrW(&H2C7), ChrW(&H203A))
            .Move chipLeftX + chipW + 3, TextTop(rowTop, ROW_HEIGHT, TEXT_SIZE), CHEVRON_WIDTH, TextLineHeight(TEXT_SIZE)
            .Visible = True
        Else
            .Visible = False
        End If
    End With

    ' Formula and value columns
    formulaLeft = PADDING + Int(ListWidth * FORMULA_COLUMN_SHARE)
    If formulaLeft < chipLeftX + chipW + CHEVRON_WIDTH + 6 Then formulaLeft = chipLeftX + chipW + CHEVRON_WIDTH + 6
    With EnsureLabel("lblFormula" & s, "formula")
        .Caption = OneLine(NodeColumn(nodeIndex, 2))
        .Move formulaLeft, TextTop(rowTop, ROW_HEIGHT, TEXT_SIZE - 0.5), IIf(valueLeft - 8 > formulaLeft, valueLeft - 8 - formulaLeft, 0), TextLineHeight(TEXT_SIZE - 0.5)
        .Visible = True
    End With
    With EnsureLabel("lblValue" & s, "value")
        .Caption = NodeColumn(nodeIndex, 1)
        .ForeColor = IIf(isSelected, ColorGreen, ColorText)
        .Move valueLeft, TextTop(rowTop, ROW_HEIGHT, TEXT_SIZE - 0.5), VALUE_COLUMN_WIDTH, TextLineHeight(TEXT_SIZE - 0.5)
        .Visible = True
    End With
End Sub

Private Sub HideRow(ByVal s As Long)
    Dim a As Long
    HideControl "lblElbowV" & s
    HideControl "lblElbowH" & s
    HideChip "lblChip" & s
    HideControl "lblChev" & s
    HideControl "lblFormula" & s
    HideControl "lblValue" & s
    For a = 1 To MaxRailDepth
        HideControl "lblRail" & s & "_" & a
    Next a
End Sub

' Hide a control if it has been created
Private Sub HideControl(ctlName As String)
    On Error Resume Next
    Me.Controls(ctlName).Visible = False
End Sub

' Left edge of the chip for a tree level
Private Function ChipLeft(ByVal level As Integer) As Single
    ChipLeft = PADDING + 4 + level * INDENT
End Function

' x of the tree line that runs down from a level's parent chip
Private Function RailX(ByVal level As Integer) As Single
    RailX = PADDING + 4 + (level - 1) * INDENT + RAIL_OFFSET
End Function

' ---------------------------------------------------------------------------
' Status and layout
' ---------------------------------------------------------------------------

Private Sub UpdateStatus()
    Dim shownCount As Long
    Dim i As Long
    If Len(txtFilter.text) > 0 Then
        For i = 1 To VisibleCount
            If VisibleNodes(i) > 1 Then
                If Not TreeData(VisibleNodes(i)).isRepeat Then shownCount = shownCount + 1
            End If
        Next i
        Me.Controls("lblCount").Caption = CStr(shownCount)
        Me.Controls("lblStatus").Caption = " of " & UniqueCount & " precedents shown"
    Else
        Me.Controls("lblCount").Caption = CStr(UniqueCount)
        Me.Controls("lblStatus").Caption = IIf(UniqueCount = 1, " precedent", " precedents")
    End If
    LayoutFooterText
End Sub

' Show the form in the top-right corner of the Excel window, at the size the
' last trace window had, or twice its minimum size the first time. Call after the list is populated so the final layout pass
' sizes everything for the larger form.
Public Sub ShowAtDefaultPosition()
    On Error Resume Next

    Dim targetWidth As Single
    Dim targetHeight As Single
    targetWidth = Val(GetSetting("Breakdown", "TraceWindow", "Width", "0"))
    targetHeight = Val(GetSetting("Breakdown", "TraceWindow", "Height", "0"))
    If targetWidth = 0 Then targetWidth = MinWidth * DEFAULT_SIZE_FACTOR
    If targetHeight = 0 Then targetHeight = MinHeight * DEFAULT_SIZE_FACTOR

    ' Stay inside the Excel window, but never below the minimum size
    If targetWidth > Application.Width - 2 * WINDOW_MARGIN Then targetWidth = Application.Width - 2 * WINDOW_MARGIN
    If targetHeight > Application.UsableHeight Then targetHeight = Application.UsableHeight
    If targetWidth < MinWidth Then targetWidth = MinWidth
    If targetHeight < MinHeight Then targetHeight = MinHeight
    Me.Width = targetWidth
    Me.Height = targetHeight

    ' Top-right corner, just below the ribbon and formula bar
    Dim chromeHeight As Single
    chromeHeight = Application.Height - Application.UsableHeight
    If chromeHeight < 0 Or chromeHeight > Application.Height / 2 Then chromeHeight = 0

    Me.StartUpPosition = 0  ' Manual
    Me.Left = Application.Left + Application.Width - Me.Width - WINDOW_MARGIN
    Me.Top = Application.Top + chromeHeight
    If Me.Top + Me.Height > Application.Top + Application.Height Then
        Me.Top = Application.Top + Application.Height - Me.Height
    End If
    If Me.Left < Application.Left Then Me.Left = Application.Left
    If Me.Top < Application.Top Then Me.Top = Application.Top

    ResizeControls
    LastWidth = Me.Width
    LastHeight = Me.Height

    On Error GoTo 0
    Me.Show vbModeless
End Sub

' Resizer event handlers
Private Sub UserForm_Activate()
    ' Make the form resizable
    FormResizer.Activate
    ' Lay out again now the form is visible; Excel for Mac may not apply
    ' control size changes made before the form is shown
    ResizeControls
    On Error Resume Next
    txtKeys.SetFocus
End Sub

Private Sub UserForm_Resize()
    If IsDragResizing Then Exit Sub

    ' Prevent resizing below minimum dimensions
    If Me.Width < MinWidth Then Me.Width = MinWidth
    If Me.Height < MinHeight Then Me.Height = MinHeight

    ' Only resize controls if dimensions actually changed
    If Me.Width <> LastWidth Or Me.Height <> LastHeight Then
        ResizeControls
        LastWidth = Me.Width
        LastHeight = Me.Height
    End If
End Sub

' Closing the window (Esc or the close button) ends the drill-down: close
' the hidden windows kept for Shift+Enter as well
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    RememberWindowSize
    If SwitchingWindows Then Exit Sub

    Dim hiddenWindow As Variant
    For Each hiddenWindow In TraceHistory
        hiddenWindow.CloseQuietly
    Next hiddenWindow
    Set TraceHistory = New Collection
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next

    ' Clean up the resizer object to prevent crashes
    Set FormResizer = Nothing

    On Error GoTo 0
End Sub

Private Function FormInnerWidth() As Single
    FormInnerWidth = Me.InsideWidth
    If FormInnerWidth <= 0 Then FormInnerWidth = Me.Width - 4
End Function

Private Function FormInnerHeight() As Single
    FormInnerHeight = Me.InsideHeight
    If FormInnerHeight <= 0 Then FormInnerHeight = Me.Height - 22  ' approx. title bar
End Function

Private Function CardTop() As Single
    CardTop = TOP_BAR_HEIGHT + 1 + PADDING
End Function

Private Function CardRight() As Single
    CardRight = FormInnerWidth() - PADDING
End Function

' Lay out the top bar, formula card, column captions, list and footer to fill
' the client area
Private Sub ResizeControls()
    On Error Resume Next

    Dim w As Single
    Dim h As Single
    w = FormInnerWidth()
    h = FormInnerHeight()
    If w <= 4 * PADDING + VALUE_COLUMN_WIDTH Then Exit Sub

    ' Top bar
    Dim barTextTop As Single
    Dim x As Single
    Dim filterRight As Single
    barTextTop = Int((TOP_BAR_HEIGHT - CAPTION_HEIGHT - 2) / 2)
    Me.Controls("lblTopRule").Move 0, TOP_BAR_HEIGHT, w, 1
    With Me.Controls("lblTopTitle")
        .Left = PADDING
        .Top = barTextTop
        x = .Left + .Width + 6
    End With
    Dim rootChip As MSForms.Label
    Set rootChip = Me.Controls("lblRootChip")
    PlaceChip rootChip, x, Int((TOP_BAR_HEIGHT - CHIP_HEIGHT) / 2), ChipWidth(rootChip.Caption), True, ColorWindow
    x = x + ChipWidth(rootChip.Caption) + 12

    ' "esc" hint: an outlined rounded box at the right
    Dim escLeft As Single
    Dim escTop As Single
    escLeft = w - PADDING - ESC_WIDTH
    escTop = Int((TOP_BAR_HEIGHT - ESC_HEIGHT) / 2)
    PlaceRoundBox "lblEsc", escLeft, escTop, ESC_WIDTH, ESC_HEIGHT, ColorCard, CHIP_RADIUS, ColorWindow, ColorCardBorder
    Me.Controls("lblEscHint").Move escLeft + TEXT_NUDGE_X, TextTop(escTop, ESC_HEIGHT, CAPTION_SIZE), ESC_WIDTH, TextLineHeight(CAPTION_SIZE)
    Me.Controls("lblSheetName").Move escLeft - 8 - 120, barTextTop, 120, CAPTION_HEIGHT + 2
    filterRight = escLeft - 8 - 120 - 8
    If filterRight < x + 40 Then filterRight = x + 40
    Me.Controls("lblFilterHint").Move x + 2, barTextTop, filterRight - x - 2, CAPTION_HEIGHT + 2
    txtFilter.Move x, barTextTop - 3, filterRight - x, CAPTION_HEIGHT + 8

    ' Formula card: caption and value, then the formula laid out below them
    Dim cardHeight As Single
    RenderFormulaCard
    cardHeight = 2 * CARD_PADDING + CARD_CAPTION_HEIGHT + CardLines * FORMULA_LINE_HEIGHT
    PlaceRoundBox "lblCard", PADDING, CardTop(), w - 2 * PADDING, cardHeight, ColorCard, CARD_RADIUS, ColorWindow, ColorCardBorder
    Me.Controls("lblCardCaption").Move PADDING + CARD_PADDING, CardTop() + CARD_PADDING, w / 2, CAPTION_HEIGHT
    Me.Controls("lblCardValue").Move w / 2, CardTop() + CARD_PADDING - 4, w / 2 - PADDING - CARD_PADDING, CARD_CAPTION_HEIGHT + 2

    ' Column captions
    Dim headerTop As Single
    headerTop = CardTop() + cardHeight + PADDING
    ListWidth = w - 2 * PADDING - SCROLLBAR_WIDTH - 4
    Me.Controls("lblColCell").Move PADDING + 4, headerTop, 100, CAPTION_HEIGHT
    Me.Controls("lblColFormula").Move PADDING + Int(ListWidth * FORMULA_COLUMN_SHARE), headerTop, 100, CAPTION_HEIGHT
    Me.Controls("lblColValue").Move PADDING + ListWidth - VALUE_COLUMN_WIDTH, headerTop, VALUE_COLUMN_WIDTH, CAPTION_HEIGHT

    ' Footer along the bottom edge
    Dim footerTop As Single
    footerTop = h - FOOTER_HEIGHT
    Me.Controls("lblFooterRule").Move 0, footerTop, w, 1
    Me.Controls("lblKeys").Move w * 0.35, footerTop + 4, w * 0.65 - PADDING - GRIP_SIZE, CAPTION_HEIGHT
    LayoutGrip w, h
    LayoutFooterText

    ' List rows fill the space between the captions and the footer; rows that
    ' no longer fit are hidden
    Dim listHeight As Single
    Dim s As Long
    Dim newSlotCount As Long
    ListTop = headerTop + COLUMN_HEADER_HEIGHT
    listHeight = footerTop - 4 - ListTop
    newSlotCount = Int(listHeight / ROW_HEIGHT)
    If newSlotCount < 1 Then newSlotCount = 1
    For s = newSlotCount To SlotCount - 1
        HideRow s
    Next s
    SlotCount = newSlotCount

    scrList.Move w - PADDING - SCROLLBAR_WIDTH, ListTop, SCROLLBAR_WIDTH, SlotCount * ROW_HEIGHT
    lblListHit.Move PADDING - 4, ListTop, ListWidth + 8, SlotCount * ROW_HEIGHT
    lblListHit.ZOrder fmZOrderFront

    RenderRows
    Me.Repaint
    On Error GoTo 0
End Sub

' Six dots in a triangle in the bottom-right corner, with the drag layer over
' them
Private Sub LayoutGrip(ByVal w As Single, ByVal h As Single)
    Dim dotCol As Variant
    Dim dotRow As Variant
    Dim dot As Long
    dotCol = Array(2, 1, 2, 0, 1, 2)
    dotRow = Array(0, 1, 1, 2, 2, 2)
    For dot = 0 To 5
        Me.Controls("lblGripDot" & dot).Move w - 4.5 - (2 - dotCol(dot)) * 3, h - 4.5 - (2 - dotRow(dot)) * 3, 1.5, 1.5
    Next dot
    lblGrip.Move w - GRIP_SIZE, h - GRIP_SIZE, GRIP_SIZE, GRIP_SIZE
    lblGrip.ZOrder fmZOrderFront
End Sub

' Dragging the grip resizes the window by as much as the mouse moved. The grip
' moves with the corner, so the mouse stays over it.
Private Sub lblGrip_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GripStartX = X
    GripStartY = Y
End Sub

Private Sub lblGrip_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim newWidth As Single
    Dim newHeight As Single
    If Button <> LEFT_BUTTON Then Exit Sub

    newWidth = Me.Width + X - GripStartX
    newHeight = Me.Height + Y - GripStartY
    If newWidth < MinWidth Then newWidth = MinWidth
    If newHeight < MinHeight Then newHeight = MinHeight
    If Abs(newWidth - Me.Width) < 2 And Abs(newHeight - Me.Height) < 2 Then Exit Sub

    ' One layout pass for both changes, instead of one per UserForm_Resize
    IsDragResizing = True
    Me.Width = newWidth
    Me.Height = newHeight
    IsDragResizing = False
    ResizeControls
    LastWidth = Me.Width
    LastHeight = Me.Height
End Sub

Private Sub lblGrip_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RememberWindowSize
End Sub

' The next trace window opens at this size (see ShowAtDefaultPosition)
Private Sub RememberWindowSize()
    SaveSetting "Breakdown", "TraceWindow", "Width", CStr(Int(Me.Width))
    SaveSetting "Breakdown", "TraceWindow", "Height", CStr(Int(Me.Height))
End Sub

' Put the "precedents" caption right after the count; both size to their text
Private Sub LayoutFooterText()
    On Error Resume Next
    Dim textTop As Single
    textTop = Me.Controls("lblFooterRule").Top + 4
    With Me.Controls("lblCount")
        .Left = PADDING
        .Top = textTop
    End With
    With Me.Controls("lblStatus")
        .Left = Me.Controls("lblCount").Left + Me.Controls("lblCount").Width
        .Top = textTop
    End With
    On Error GoTo 0
End Sub

