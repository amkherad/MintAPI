VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   0  'None
   Caption         =   "About miccy Picture Tools"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "About.frx":57E2
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Register(free)"
      Height          =   315
      Left            =   3675
      TabIndex        =   6
      Top             =   4050
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Donate"
      Height          =   315
      Left            =   2700
      TabIndex        =   5
      Top             =   4050
      Width           =   915
   End
   Begin VB.CommandButton okButton 
      Caption         =   "&OK"
      Height          =   315
      Left            =   5325
      TabIndex        =   4
      Top             =   4050
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by Pink Fluid"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1500
      TabIndex        =   8
      Top             =   3675
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright by deftro.com"
      ForeColor       =   &H00FBD2AE&
      Height          =   240
      Left            =   300
      TabIndex        =   7
      Top             =   4125
      Width           =   2115
   End
   Begin VB.Image closeBTN 
      Height          =   240
      Left            =   6225
      Picture         =   "About.frx":A5C9
      Top             =   225
      Width           =   240
   End
   Begin VB.Label versionCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "version 1.0.0.2012"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1500
      TabIndex        =   2
      Top             =   1050
      Width           =   1620
   End
   Begin VB.Label licence 
      BackStyle       =   0  'Transparent
      Caption         =   "LICENCE"
      ForeColor       =   &H00FFFFFF&
      Height          =   2340
      Left            =   1500
      TabIndex        =   3
      Top             =   1350
      Width           =   4740
   End
   Begin VB.Label title 
      BackStyle       =   0  'Transparent
      Height          =   1065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6765
   End
   Begin VB.Label titleCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "miccy Ultimate Tools"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   1500
      TabIndex        =   1
      Top             =   675
      Width           =   3000
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FORMWIDTH As Long = 450
Const FORMHEIGHT As Long = 300

Dim fX As Long, fY As Long

Private Sub closeBTN_Click(): Call Unload(Me): End Sub
Private Sub okButton_Click(): Call Unload(Me): End Sub

Private Sub Form_Load()
    Call Me.Move(Me.Left, Me.top, FORMWIDTH * 15, FORMHEIGHT * 15)
    Dim hRGN As Long
    hRGN = API_CreateRoundRectRgn(0, 0, FORMWIDTH, FORMHEIGHT, 7, 7)
    Call API_SetWindowRgn(hwnd, hRGN, True)
    Call API_DeleteObject(hRGN)
End Sub

Private Sub title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fX = X
    fY = Y
End Sub
Private Sub title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Me.Move(Me.Left + X - fX, Me.top + Y - fY)
    End If
End Sub
