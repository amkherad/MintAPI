VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMTDemo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' MTDemo2 - Multithreading demo program
' Copyright © 1997 by Desaware Inc.  All Rights Reserved

Option Explicit

Private Sub Class_Initialize()
   Dim f As New frmMTDemo2
   f.Show
   Set f = Nothing
End Sub

Public Sub DisplayObjPtr(cls As clsMTDemo2)
   If Not (cls Is Nothing) Then
      MsgBox "Object pointer is: " & ObjPtr(cls)
   End If
End Sub
