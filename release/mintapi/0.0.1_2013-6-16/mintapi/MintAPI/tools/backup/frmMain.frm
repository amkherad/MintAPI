VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "MintAPI Tools"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   7665
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    Call Me.Hide
End Sub
