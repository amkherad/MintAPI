VERSION 5.00
Begin VB.UserControl baEditor 
   BackColor       =   &H00C88B0D&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MousePointer    =   3  'I-Beam
   ScaleHeight     =   8115
   ScaleWidth      =   9630
   Begin MintAPI2ndLayer.ctlByteArrayEditor Editor 
      Height          =   7395
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   13044
   End
   Begin VB.Frame frmResizer 
      BackColor       =   &H00C88B0D&
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   7200
      MousePointer    =   9  'Size W E
      TabIndex        =   1
      Top             =   0
      Width           =   30
   End
   Begin MintAPI2ndLayer.ByteArrayTypedEditor te 
      Height          =   7455
      Left            =   7320
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   13150
   End
End
Attribute VB_Name = "baEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim i_sizerX As Long


Private Sub frmResizer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    i_sizerX = X / 15
End Sub
Private Sub frmResizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        Dim lft As Single, teWidth As Long
        lft = frmResizer.Left + (X / 15) - i_sizerX
        frmResizer.Left = lft
        teWidth = ScaleWidth - lft
        Call te.Move(lft + 30, 0, ScaleWidth - lft)
        Call Editor.Move(0, 0, frmResizer.Left)
    End If
End Sub
Private Sub frmResizer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Call Editor.ResizeNow
    'Call te.ResizeNow
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Dim l As Long
    If te.Visible Then
        l = te.Width
        Call te.Move(ScaleWidth - l, 0, l, ScaleHeight)
        Call frmResizer.Move(te.Left - frmResizer.Width, 0, frmResizer.Width, ScaleHeight)
        'Call te.ResizeNow
        l = l + frmResizer.Width
    End If
    Call Editor.Move(0, 0, frmResizer.Left, ScaleHeight)
    'Call Editor.ResizeNow
End Sub
