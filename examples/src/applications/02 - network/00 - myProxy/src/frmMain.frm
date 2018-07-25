VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "myProxy"
   ClientHeight    =   4605
   ClientLeft      =   9030
   ClientTop       =   6255
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   Begin VB.Frame frm 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   3435
      Left            =   540
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   5550
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   45
         TabIndex        =   9
         Text            =   "localhost\myProxy\index.php?rq=GET"
         Top             =   765
         Width           =   5460
      End
      Begin VB.CommandButton saveBtn 
         Caption         =   "Save"
         Height          =   375
         Left            =   4500
         TabIndex        =   7
         Top             =   3015
         Width           =   870
      End
      Begin VB.TextBox portNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1530
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "12648"
         Top             =   90
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proxy Server Address:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   540
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Port Number:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   135
         Width           =   1440
      End
   End
   Begin VB.TextBox log 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   2490
      Left            =   630
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   5280
   End
   Begin VB.Image settings 
      Height          =   240
      Left            =   5670
      Picture         =   "frmMain.frx":1082
      Top             =   3555
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informations and Tips"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   900
      TabIndex        =   6
      Top             =   360
      Width           =   1890
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   3030
      Left            =   585
      Shape           =   4  'Rounded Rectangle
      Top             =   450
      Width           =   5370
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   36
      X2              =   408
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Image voteDown 
      Height          =   240
      Index           =   4
      Left            =   5445
      Picture         =   "frmMain.frx":1306
      Top             =   4185
      Width           =   240
   End
   Begin VB.Image voteUp 
      Height          =   240
      Index           =   3
      Left            =   5715
      Picture         =   "frmMain.frx":1703
      Top             =   4095
      Width           =   240
   End
   Begin VB.Image status 
      Height          =   240
      Index           =   2
      Left            =   1890
      Picture         =   "frmMain.frx":196D
      Top             =   4140
      Width           =   240
   End
   Begin VB.Image connection_statusPIC 
      Height          =   240
      Index           =   0
      Left            =   1260
      Picture         =   "frmMain.frx":1BC3
      Top             =   3870
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image status 
      Height          =   240
      Index           =   0
      Left            =   1260
      Picture         =   "frmMain.frx":1E1A
      Top             =   4140
      Width           =   240
   End
   Begin VB.Label StatusLBL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   585
      TabIndex        =   3
      Top             =   4185
      Width           =   630
   End
   Begin VB.Image connection_statusPIC 
      Height          =   240
      Index           =   2
      Left            =   1800
      Picture         =   "frmMain.frx":2243
      Top             =   3870
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image connection_statusPIC 
      Height          =   240
      Index           =   1
      Left            =   1530
      Picture         =   "frmMain.frx":266C
      Top             =   3870
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image status 
      Height          =   240
      Index           =   1
      Left            =   1575
      Picture         =   "frmMain.frx":28CB
      Top             =   4140
      Width           =   240
   End
   Begin VB.Image closeBTN 
      Height          =   240
      Left            =   5760
      Picture         =   "frmMain.frx":2A1B
      Top             =   90
      Width           =   240
   End
   Begin VB.Image closeBTNH 
      Height          =   240
      Left            =   5760
      Picture         =   "frmMain.frx":2AB8
      Top             =   315
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image closeBTNNH 
      Height          =   240
      Left            =   5760
      Picture         =   "frmMain.frx":4E8A
      Top             =   495
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image iconT 
      Height          =   360
      Left            =   75
      Picture         =   "frmMain.frx":4F27
      Top             =   90
      Width           =   360
   End
   Begin VB.Shape border 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3075
      Left            =   0
      Top             =   0
      Width           =   6045
   End
   Begin VB.Label lblMove 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4605
      Left            =   225
      TabIndex        =   2
      Top             =   0
      Width           =   510
   End
   Begin VB.Shape leftP 
      BackColor       =   &H00C88B0D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000F0C6&
      BorderWidth     =   2
      Height          =   4560
      Left            =   0
      Top             =   0
      Width           =   510
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum ConnectionStatus
    CSFailure = 0
    CSWarning = 1
    CSSuccessfull = 2
End Enum
Dim CS As ConnectionStatus

Dim cwp As Boolean
Dim mdX As Long, mdY As Long


Private Sub Form_Load()
    Call Me.Move(Screen.Width - Me.Width - 30, Screen.Height - Me.Height - 500)
    frm.BackColor = vbWhite
    Call Sort
    ConnectionStatus = CS
    Serving = False
    UpdateAvailable = False
End Sub
Private Sub iconT_DblClick(): Call Unload(Me): End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Call border.Move(1, 1, ScaleWidth - 1, ScaleHeight - 1)
    Call leftP.Move(1, 1, leftP.Width, ScaleHeight - 1)
    Call lblMove.Move(0, 0, lblMove.Width, ScaleHeight)
End Sub


Private Sub lblMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        mdX = X
        mdY = Y
    End If
End Sub

Private Sub lblMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Me.Move(Me.Left + X - mdX, Me.Top + Y - mdY)
    End If
    If cwp Then _
        Set closeBTN.Picture = closeBTNNH.Picture
    cwp = False
End Sub


Private Sub closeBTN_Click(): Call Unload(Me): End Sub

Private Sub closeBTN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set closeBTN.Picture = closeBTNH.Picture
    cwp = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cwp Then _
        Set closeBTN.Picture = closeBTNNH.Picture
    cwp = False
End Sub

Private Sub Sort()
    Dim C As Control, iLeft As Long
    iLeft = StatusLBL.Left + StatusLBL.Width + 6
    For Each C In status
        If C.Visible Then
            C.Left = iLeft
            iLeft = iLeft + C.Width + 3
        End If
    Next
End Sub

Public Property Get ConnectionStatus() As ConnectionStatus
    ConnectionStatus = CS
End Property
Public Property Let ConnectionStatus(Value As ConnectionStatus)
    Set status(0).Picture = connection_statusPIC(Value).Picture
    CS = Value
End Property

Public Property Get Serving() As Boolean
    Serving = status(1).Visible
End Property
Public Property Let Serving(Value As Boolean)
    status(1).Visible = Value
    Call Sort
End Property
Public Property Get UpdateAvailable() As Boolean
    UpdateAvailable = status(2).Visible
End Property
Public Property Let UpdateAvailable(Value As Boolean)
    status(2).Visible = Value
    Call Sort
End Property

Private Sub saveBtn_Click()
    frm.Visible = False
    settings.Visible = True
End Sub

Private Sub settings_Click()
    frm.Visible = True
    settings.Visible = False
End Sub
