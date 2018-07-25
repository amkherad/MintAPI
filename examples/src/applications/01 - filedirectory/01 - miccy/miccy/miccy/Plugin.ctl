VERSION 5.00
Begin VB.UserControl Plugin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   ScaleHeight     =   1665
   ScaleWidth      =   11700
   Begin VB.Line lnFocus 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   0
      X2              =   12225
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label info 
      BackStyle       =   0  'Transparent
      Caption         =   "[LICENCE]"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   7
      Left            =   5100
      TabIndex        =   16
      Tag             =   "[LICENCE]"
      Top             =   450
      Width           =   2235
   End
   Begin VB.Label captionLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Licence:"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   5
      Left            =   3900
      TabIndex        =   15
      Top             =   450
      Width           =   1170
   End
   Begin VB.Image ico 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   150
      Stretch         =   -1  'True
      Top             =   60
      Width           =   390
   End
   Begin VB.Label info 
      BackStyle       =   0  'Transparent
      Caption         =   "[EMAIL]"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   6
      Left            =   5100
      TabIndex        =   14
      Tag             =   "[EMAIL]"
      Top             =   900
      Width           =   2235
   End
   Begin VB.Label captionLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   7
      Left            =   3900
      TabIndex        =   13
      Top             =   900
      Width           =   1170
   End
   Begin VB.Label info 
      BackStyle       =   0  'Transparent
      Caption         =   "[COMPANY]"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   5
      Left            =   5100
      TabIndex        =   12
      Tag             =   "[COMPANY]"
      Top             =   675
      Width           =   2235
   End
   Begin VB.Label captionLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   6
      Left            =   3900
      TabIndex        =   11
      Top             =   675
      Width           =   1170
   End
   Begin VB.Label info 
      BackStyle       =   0  'Transparent
      Caption         =   "[PLUGINS]"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   4
      Left            =   1275
      TabIndex        =   10
      Tag             =   "[PLUGINS]"
      Top             =   1350
      Width           =   2235
   End
   Begin VB.Label info 
      BackStyle       =   0  'Transparent
      Caption         =   "[TOOLKIT]"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   1275
      TabIndex        =   9
      Tag             =   "[TOOLKIT]"
      Top             =   1125
      Width           =   2235
   End
   Begin VB.Label info 
      BackStyle       =   0  'Transparent
      Caption         =   "[PROGRAMMER]"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   1275
      TabIndex        =   8
      Tag             =   "[PROGRAMMER]"
      Top             =   900
      Width           =   2235
   End
   Begin VB.Label info 
      BackStyle       =   0  'Transparent
      Caption         =   "[PUBLISHER]"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   1275
      TabIndex        =   7
      Tag             =   "[PUBLISHER]"
      Top             =   675
      Width           =   2235
   End
   Begin VB.Label info 
      BackStyle       =   0  'Transparent
      Caption         =   "[VERSION]"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   1275
      TabIndex        =   6
      Tag             =   "[VERSION]"
      Top             =   450
      Width           =   2235
   End
   Begin VB.Label captionLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Plugins:"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   4
      Left            =   75
      TabIndex        =   5
      Top             =   1350
      Width           =   1170
   End
   Begin VB.Label captionLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Toolkit:"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   4
      Top             =   1125
      Width           =   1170
   End
   Begin VB.Label captionLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Programmer:"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   75
      TabIndex        =   3
      Top             =   900
      Width           =   1170
   End
   Begin VB.Label captionLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Publisher:"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   2
      Top             =   675
      Width           =   1170
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FBD2AE&
      X1              =   0
      X2              =   12225
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Label captionLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   450
      Width           =   1170
   End
   Begin VB.Label title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Plugin Name"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B58004&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   75
      Width           =   1485
   End
End
Attribute VB_Name = "Plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const WWIDTH As Long = 11930
Const WHEIGHT As Long = 1665

Const MINPIXMOVTOSCROLL As Long = 10

Public Event LeftClick(X As Long, Y As Long)
Public Event RightClick(X As Long, Y As Long)
Public Event StartScrolling(ToTop As Boolean, Power As Double)
Public Event magneticScrolling(Value As Long)

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Dim mDY As Long, secondDY As Long
'Dim downScroll As Boolean, magneticScrolling As Boolean
'Dim isScrolling As Boolean

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer): RaiseEvent KeyDown(KeyCode, Shift): End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer): RaiseEvent KeyPress(KeyAscii): End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer): RaiseEvent KeyUp(KeyCode, Shift): End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        mDY = Y / 15
'        downScroll = True
'        If isScrolling Then
'            magneticScrolling = True
'            RaiseEvent LeftClick(CLng(X) / 15, CLng(Y) / 15)
'            Exit Sub
'        End If
'        'scroller.Enabled = True
'        magneticScrolling = False
'        RaiseEvent LeftClick(CLng(X) / 15, CLng(Y) / 15)
'    ElseIf Button = 2 Then
'        RaiseEvent RightClick(CLng(X) / 15, CLng(Y) / 15)
'    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        secondDY = Y / 15
'        RaiseEvent magneticScrolling(secondDY - mDY)
'    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        'If downScroll Then
'            'If scroller.Enabled Then
'                If (secondDY - mDY) > MINPIXMOVTOSCROLL Then
'                    RaiseEvent StartScrolling(True, CDbl(secondDY - mDY) / CDbl(5))
'                ElseIf (secondDY - mDY) < -MINPIXMOVTOSCROLL Then
'                    RaiseEvent StartScrolling(False, CDbl(mDY - secondDY) / CDbl(5))
'                End If
'            'End If
'        'End If
'        'downScroll = False
'        magneticScrolling = False
'    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
'Private Sub scroller_Timer()
'    If magneticScrolling Then GoTo exitsub
'    If downScroll Then
'        magneticScrolling = True
''    Else
''        If (secondDY - mDY) > MINPIXMOVTOSCROLL Then
''            RaiseEvent StartScrolling(True, CDbl(secondDY - mDY) / CDbl(5))
''        ElseIf (secondDY - mDY) < -MINPIXMOVTOSCROLL Then
''            RaiseEvent StartScrolling(False, CDbl(mDY - secondDY) / CDbl(5))
''        End If
'    End If
'exitsub:
'    'scroller.Enabled = False
'End Sub
Private Sub info_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single): Call UserControl_MouseDown(Button, Shift, info(Index).Left + X, info(Index).top + Y): End Sub
Private Sub info_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single): Call UserControl_MouseMove(Button, Shift, info(Index).Left + X, info(Index).top + Y): End Sub
Private Sub captionLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single): Call UserControl_MouseDown(Button, Shift, captionLabel(Index).Left + X, captionLabel(Index).top + Y): End Sub
Private Sub captionLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single): Call UserControl_MouseMove(Button, Shift, captionLabel(Index).Left + X, captionLabel(Index).top + Y): End Sub
Private Sub title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): Call UserControl_MouseDown(Button, Shift, title.Left + X, title.top + Y): End Sub
Private Sub title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call UserControl_MouseMove(Button, Shift, title.Left + X, title.top + Y): End Sub
Private Sub ico_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): Call UserControl_MouseDown(Button, Shift, ico.Left + X, ico.top + Y): End Sub
Private Sub ico_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call UserControl_MouseMove(Button, Shift, ico.Left + X, ico.top + Y): End Sub

Private Sub UserControl_Initialize()
    Width = WWIDTH
    Height = WHEIGHT
End Sub
Private Sub UserControl_InitProperties()
    Width = WWIDTH
    Height = WHEIGHT
End Sub

Private Sub UserControl_Resize()
    Width = WWIDTH
    Height = WHEIGHT
End Sub

Private Sub UserControl_EnterFocus()
    BackColor = &HF9EFE3
    lnFocus.Visible = True
End Sub
Private Sub UserControl_ExitFocus()
    BackColor = vbWhite
    lnFocus.Visible = False
End Sub
Private Sub UserControl_GotFocus()
    BackColor = &HF9EFE3
    lnFocus.Visible = True
End Sub
Private Sub UserControl_LostFocus()
    BackColor = vbWhite
    lnFocus.Visible = False
End Sub

Private Property Get VSTR(Name As String) As String
    Dim i As Long
    For i = info.LBound To info.UBound
        If CStr(info(i).tag) = Name Then
            VSTR = info(i).Caption
            Exit Property
        End If
    Next
    throw ItemNotExistsException
End Property
Private Property Let VSTR(Name As String, Value As String)
    Dim i As Long
    For i = info.LBound To info.UBound
        If CStr(info(i).tag) = Name Then
            info(i).Caption = Replace(info(i).tag, Name, Value)
            Exit Property
        End If
    Next
    throw ItemNotExistsException
End Property

Public Property Get Caption() As String
    Caption = title.Caption
End Property
Public Property Let Caption(Value As String)
    title.Caption = Value
End Property
Public Property Get Version() As String
    Version = VSTR("[VERSION]")
End Property
Public Property Let Version(Value As String)
    VSTR("[VERSION]") = Value
End Property
Public Property Get Publisher() As String
    Publisher = VSTR("[PUBLISHER]")
End Property
Public Property Let Publisher(Value As String)
    VSTR("[PUBLISHER]") = Value
End Property
Public Property Get Programmer() As String
    Programmer = VSTR("[PROGRAMMER]")
End Property
Public Property Let Programmer(Value As String)
    VSTR("[PROGRAMMER]") = Value
End Property
Public Property Get Toolkit() As String
    Toolkit = VSTR("[TOOLKIT]")
End Property
Public Property Let Toolkit(Value As String)
    VSTR("[TOOLKIT]") = Value
End Property
Public Property Get plugins() As String
    plugins = VSTR("[PLUGINS]")
End Property
Public Property Let plugins(Value As String)
    VSTR("[PLUGINS]") = Value
End Property
Public Property Get licence() As String
    licence = VSTR("[LICENCE]")
End Property
Public Property Let licence(Value As String)
    VSTR("[LICENCE]") = Value
End Property
Public Property Get Company() As String
    Company = VSTR("[COMPANY]")
End Property
Public Property Let Company(Value As String)
    VSTR("[COMPANY]") = Value
End Property
Public Property Get Email() As String
    Email = VSTR("[EMAIL]")
End Property
Public Property Let Email(Value As String)
    VSTR("[EMAIL]") = Value
End Property
