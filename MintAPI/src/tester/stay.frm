VERSION 5.00
Begin VB.Form stay 
   BackColor       =   &H00EFAE00&
   Caption         =   "stay"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1665
      Top             =   1260
   End
End
Attribute VB_Name = "stay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim t As Long
Private Sub Timer1_Timer()
t = t + 1
Caption = t
End Sub
