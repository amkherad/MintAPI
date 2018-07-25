Attribute VB_Name = "modMTDemo2"
' MTDemo2 - Multithreading demo program
' Copyright © 1997 by Desaware Inc.  All Rights Reserved
Option Explicit
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Sub Main()
   Dim f As frmMTDemo2
   ' We need this because Main is called on each new thread
   Dim hwnd As Long
   hwnd = FindWindow(vbNullString, "Multithreading Demo2")
   If hwnd = 0 Then
      Set f = New frmMTDemo2
      f.Show
      Set f = Nothing
   End If
End Sub
