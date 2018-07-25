VERSION 5.00
Begin VB.Form frmMTDemo2 
   Caption         =   "Multithreading Demo2"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLongOp 
      Caption         =   "Long Op"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   1005
   End
   Begin VB.CommandButton cmdLaunch2 
      Caption         =   "Launch Create"
      Height          =   375
      Left            =   1620
      TabIndex        =   3
      Top             =   630
      Width           =   1365
   End
   Begin VB.CommandButton cmdLaunch1 
      Caption         =   "Launch New"
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   630
      Width           =   1365
   End
   Begin VB.Label lblThread 
      Height          =   285
      Left            =   1530
      TabIndex        =   1
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Thread ID:"
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   915
   End
End
Attribute VB_Name = "frmMTDemo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MTDemo2 - Multithreading demo program
' Copyright © 1997 by Desaware Inc.  All Rights Reserved

Option Explicit


Private Sub cmdLaunch1_Click()
   Dim c As New clsMTDemo2
   c.DisplayObjPtr Nothing
End Sub

Private Sub cmdLaunch2_Click()
   Dim c As clsMTDemo2
   Set c = CreateObject("MTDemo2.clsMTDemo2")
End Sub


Private Sub cmdLongOp_Click()
   Dim l&
   Dim s$
   For l = 1 To 1000000
      s = Chr$(l And &H7F)
   Next l
End Sub

Private Sub Form_Load()
   lblThread.Caption = Str$(App.ThreadID)
End Sub
