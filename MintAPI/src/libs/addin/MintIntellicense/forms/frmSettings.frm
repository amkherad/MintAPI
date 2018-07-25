VERSION 5.00
Begin VB.Form frmSettings 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "MintIntellicense Settings"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   501
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   683
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Frame frame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   7455
   End
   Begin VB.Shape bkFooter 
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   6960
      Width           =   9615
   End
   Begin VB.Shape bkColor 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   7815
      Left            =   0
      Top             =   -240
      Width           =   2895
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONTROLS_PADDING As Long = 5

Private Sub Form_Resize()
On Error Resume Next
Call bkColor.Move(-2, -2, bkColor.Width, Height:=ScaleHeight)
Call bkFooter.Move(-2, ScaleHeight - bkFooter.Height + 2, ScaleWidth + 4)
Call frame.Move(bkColor.Width, 0, ScaleWidth - bkColor.Width, ScaleHeight - bkFooter.Height)
Call btnSave.Move(ScaleWidth - btnSave.Width - CONTROLS_PADDING, ScaleHeight - btnSave.Height - CONTROLS_PADDING)
End Sub
