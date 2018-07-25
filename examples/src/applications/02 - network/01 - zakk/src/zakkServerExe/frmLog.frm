VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "zakk Server Log Form"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt 
      Height          =   6000
      Left            =   1980
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   90
      Width           =   7620
   End
   Begin VB.Frame frm 
      Caption         =   "Tools"
      Height          =   6000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1950
      Begin VB.CommandButton btnClear 
         Caption         =   "Clear Logs"
         Height          =   420
         Left            =   90
         TabIndex        =   3
         Top             =   810
         Width           =   1770
      End
      Begin VB.CommandButton btnShutdown 
         Caption         =   "Shutdown Server"
         Height          =   420
         Left            =   90
         TabIndex        =   2
         Top             =   315
         Width           =   1770
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Ali Mousavi Kherad"
         Height          =   960
         Left            =   90
         TabIndex        =   4
         Top             =   1350
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClear_Click(): txt.Text = "": End Sub
Private Sub btnShutdown_Click(): Call Unload(Me): End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call frm.Move(0, 0, frm.Width, ScaleHeight)
    Call txt.Move(frm.Width, 0, ScaleWidth - frm.Width, ScaleHeight)
End Sub
