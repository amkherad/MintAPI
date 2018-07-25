VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Friend Sub Initialize(Path As String)
    Call SetPath(Path)
End Sub

Public Sub SetPath(Path As String)
    
End Sub



Public Function Exists(Optional OptionalPath As String) As Boolean
    
End Function
Public Function Create(Optional OptionalPath As String) As FileMethods
    Call MakeTreeDirectories(GetFilePath(OptionalPath))
    Call CreateFile(OptionalPath)
    Set Create = Me
End Function
Public Sub Delete(Optional OptionalPath As String)
    
End Sub
