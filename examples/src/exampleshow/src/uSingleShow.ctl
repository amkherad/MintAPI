VERSION 5.00
Begin VB.UserControl uSingleShow 
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   ScaleHeight     =   7005
   ScaleWidth      =   9480
   Begin VB.Frame frm 
      BackColor       =   &H00F8D5BA&
      BorderStyle     =   0  'None
      Height          =   6090
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   8025
   End
   Begin VB.Label lblBack 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "< Back"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   540
   End
   Begin VB.Shape frmBorder 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   6495
      Left            =   0
      Top             =   360
      Width           =   9195
   End
End
Attribute VB_Name = "uSingleShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub lblBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblBack.ForeColor = vbBlue
    lblBack.FontUnderline = True
End Sub

