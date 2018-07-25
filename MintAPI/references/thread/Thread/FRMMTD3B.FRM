VERSION 5.00
Begin VB.Form frmMTDemo3B 
   Caption         =   "Background Progress"
   ClientHeight    =   570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleWidth      =   3030
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblThreadID 
      Height          =   195
      Left            =   1170
      TabIndex        =   1
      Top             =   90
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Thread ID:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   825
   End
End
Attribute VB_Name = "frmMTDemo3B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   lblThreadID.Caption = Str$(App.threadid)
End Sub

