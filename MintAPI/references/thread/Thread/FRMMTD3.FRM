VERSION 5.00
Begin VB.Form frmMTDemo3 
   Caption         =   "Multithreading Demo 3"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreateFree 
      Caption         =   "Create Free"
      Height          =   375
      Left            =   2070
      TabIndex        =   3
      Top             =   720
      Width           =   1635
   End
   Begin VB.CommandButton cmdCreateApt 
      Caption         =   "Create Apartment"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   1635
   End
   Begin VB.Label lblThreadID 
      Height          =   195
      Left            =   1170
      TabIndex        =   1
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Thread ID:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   825
   End
End
Attribute VB_Name = "frmMTDemo3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmMTDemo3.frm
' MTDemo3 multithreading example
' Copyright (c) 1997 by Desaware Inc.
' All Rights Reserved

Option Explicit

Dim WithEvents c As clsBackground
Attribute c.VB_VarHelpID = -1

Private Sub c_DoneCounting()
   MsgBox "DoneCounting event has been raised"
   Set c = Nothing
End Sub

Private Sub cmdCreateApt_Click()
   Set c = New clsBackground
   StartBackgroundThreadApt c
End Sub

Private Sub cmdCreateFree_Click()
   Set c = New clsBackground
   StartBackgroundThreadFree c
End Sub

Private Sub Form_Load()
   lblThreadID.Caption = Str$(App.threadid)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set c = Nothing
End Sub
