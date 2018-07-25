Attribute VB_Name = "mint_static_mngr"
Option Explicit

Dim stt_AssemblyStatic As AssemblyStatic

Public Function Assembly() As AssemblyStatic
    If stt_AssemblyStatic Is Nothing Then _
        Set stt_AssemblyStatic = New AssemblyStatic
    Set Assembly = stt_AssemblyStatic
End Function
