VERSION 5.00
Begin VB.UserControl ctlByteArrayEditor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   ScaleHeight     =   503
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   Begin VB.VScrollBar vscr 
      Height          =   7455
      LargeChange     =   75
      Left            =   8400
      Max             =   -32768
      Min             =   -32768
      MousePointer    =   1  'Arrow
      SmallChange     =   50
      TabIndex        =   0
      Top             =   0
      Value           =   -32768
      Width           =   255
   End
End
Attribute VB_Name = "ctlByteArrayEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const DRAW_BORDER_FOCUSED As Long = &H6A4300
Const DRAW_BORDER_UNFOCUSED As Long = &HFFEADD

Dim ba As ByteArray

Dim draw_IsFocused As Boolean

Dim drawRect_LastFullWidth As Long
Dim drawRect_LastFullHeight As Long
Dim drawRect_Width As Long
Dim drawRect_Height As Long
Dim drawRect_Top As Long
Dim drawRect_Left As Long
Dim drawRect_Bottom As Long
Dim drawRect_Right As Long

Dim i_selStart As Long
Dim i_selLength As Long
Dim i_selEnd As Long


'Private Sub UserControl_Initialize()
'    Call UserControl_Resize   'ResizeNow
'    Call DrawBorderGrids
'End Sub
'Private Sub UserControl_GotFocus()
'    draw_IsFocused = True
'    Call DrawBorderGrids
'End Sub
'Private Sub UserControl_LostFocus()
'    draw_IsFocused = False
'    Call DrawBorderGrids
'End Sub
'Private Sub UserControl_Paint()
'    Call ReDraw
'End Sub
'Public Sub UserControl_Resize() 'ResizeNow
'    On Error Resume Next
'    Call vscr.Move(ScaleWidth - vscr.Width, 0, vscr.Width, ScaleHeight)
'    Call ReDraw
'End Sub

Public Property Get ScrollValue() As Long
    ScrollValue = vscr.Value + 32768
End Property
Public Property Let ScrollValue(Value As Long)
    If Value < 0 Then throw Exceptions.NegativeArgumentException("Negative Scroll Value.")
    vscr.Value = Value - 32768
End Property
Public Property Get ScrollMax() As Long
    ScrollMax = vscr.Max + 32768
End Property
Public Property Let ScrollMax(Value As Long)
    If Value < 0 Then throw Exceptions.NegativeArgumentException("Negative Scroll Max Value.")
    vscr.Max = Value - 32768
    If Value = 0 Then vscr.Visible = False
End Property


Public Sub ReDraw()
    Call Cls
    If drawRect_LastFullWidth <> ScaleWidth Or _
       drawRect_LastFullHeight <> ScaleHeight Then Call CalculateVars
    Call DrawGrids
End Sub
Private Sub CalculateVars()
    drawRect_LastFullWidth = ScaleWidth
    drawRect_LastFullHeight = ScaleHeight
    drawRect_Width = drawRect_LastFullWidth - 1
    If vscr.Visible Then drawRect_Width = drawRect_Width - vscr.Width - 1
    drawRect_Height = drawRect_LastFullHeight - 1
    drawRect_Top = 0
    drawRect_Left = 0
    drawRect_Bottom = drawRect_Top + drawRect_Height
    drawRect_Right = drawRect_Left + drawRect_Width
End Sub
Private Sub DrawBorderGrids()
    Dim Color As Long
    If draw_IsFocused Then
        Color = DRAW_BORDER_FOCUSED
    Else
        Color = DRAW_BORDER_UNFOCUSED
    End If
    Line (drawRect_Left, drawRect_Top)-(drawRect_Right, drawRect_Top), Color
    Line (drawRect_Left, drawRect_Top)-(drawRect_Left, drawRect_Bottom), Color
    Line (drawRect_Left, drawRect_Bottom)-(drawRect_Right, drawRect_Bottom), Color
    Line (drawRect_Right, drawRect_Top)-(drawRect_Right, drawRect_Bottom), Color
End Sub
Private Sub DrawGrids()
    Call DrawBorderGrids
End Sub

Private Sub vscr_Validate(Cancel As Boolean)
    'Cancel = True
End Sub
