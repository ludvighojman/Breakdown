VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Makes a UserForm resizable on Windows by adding a sizing border to its
' window. Excel for Mac has no user32: the Mac branch declares nothing and
' Activate does nothing there (forms keep their fixed size on Mac).
#If Mac Then
    ' No API declarations on Mac
#ElseIf VBA7 Then
    Private Declare PtrSafe Function GetWindowLongW Lib "user32.dll" Alias "GetWindowLongPtrW" ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long) _
        As LongPtr

    Private Declare PtrSafe Function SetWindowLongW Lib "user32.dll" Alias "SetWindowLongPtrW" ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As LongPtr) _
        As LongPtr

    Private Declare PtrSafe Function FindWindowW Lib "user32" ( _
        ByVal lpClassName As LongPtr, _
        ByVal lpWindowName As LongPtr) _
        As LongPtr
#Else
    Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long) _
        As Long

    Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) _
        As Long

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) _
        As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZEBOX As Long = &H20000

Private FormCaption As String

' Remember the caption the form's window is found by
Public Sub NewForm(Caption As String)
    FormCaption = Caption
End Sub

' Add a sizing border and maximize/minimize buttons to the form's window
Public Sub Activate()
#If Mac Then
    ' Window-style changes need user32; not available on Mac
    Exit Sub
#Else
    Dim hWnd As LongPtr
    Dim lStyle As LongPtr

    #If VBA7 Then
        hWnd = FindWindowW(0, StrPtr(FormCaption))
    #Else
        hWnd = FindWindow(vbNullString, FormCaption)
    #End If
    If hWnd = 0 Then Exit Sub

    #If VBA7 Then
        lStyle = GetWindowLongW(hWnd, GWL_STYLE)
        SetWindowLongW hWnd, GWL_STYLE, lStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
    #Else
        lStyle = GetWindowLong(hWnd, GWL_STYLE)
        SetWindowLong hWnd, GWL_STYLE, lStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
    #End If
#End If
End Sub
