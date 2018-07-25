Attribute VB_Name = "baseAPI"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "baseAPI"

Public Declare Sub API_ZeroMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long)
Public Declare Function API_VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub API_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
