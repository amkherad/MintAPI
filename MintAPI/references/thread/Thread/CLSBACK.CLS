VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Class clsBackground
' MTDemo 3 - Multithreading example
' Copyright © 1997 by Desaware Inc. All Rights Reserved
Option Explicit

Event DoneCounting()

Dim l As Long

Public Function DoTheCount(ByVal finalval&) As Boolean
   Dim s As String
   If l = 0 Then
      s$ = "In Thread " & App.threadid
      Call MessageBox(0, s$, "", 0)
   End If
   l = l + 1
   If l >= finalval Then
      l = 0
      DoTheCount = True
      Call MessageBox(0, "Done with counting", "", 0)
      RaiseEvent DoneCounting
   End If
End Function

Public Function ShowAForm()
   Dim f As New frmMTDemo3B
   f.Show
End Function
