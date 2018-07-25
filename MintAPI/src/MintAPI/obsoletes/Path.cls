VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const NOT_FOUND             As Long = &H0
Private Const vbBackSlashBackSlash  As Long = &H5C005C



Public Property Get MaxPathLength() As Long
    
End Property

Public Property Get AltDirectorySeparatorChar() As Integer
    
End Property

Public Property Get AltDirectorySeparatorString() As String
    
End Property

Public Property Get DirectorySeparatorChar() As Integer
    
End Property

Public Property Get DirectorySeparatorString() As String
    
End Property

Public Property Get InvalidPathChars() As Integer()
    
End Property

Public Property Get PathSeparator() As Integer
    
End Property

Public Property Get PathSeparatorString() As String
    
End Property

Public Property Get VolumeSeparatorChar() As Integer
    
End Property

Public Property Get VolumeSeparatorString() As String
    
End Property

Public Function ChangeExtension(ByVal Path As String, ByVal Extension As String) As String
    
End Function

Public Function IsPathRooted(ByVal Path As String) As Boolean
    
End Function

Public Function HasExtension(ByVal Path As String) As Boolean
    
End Function

Public Function Combine(ByVal Path1 As String, ByVal Path2 As String) As String
    
End Function

Public Function GetDirectoryName(ByVal Path As String) As String
    
End Function

Public Function GetExtension(ByVal Path As String) As String
    
End Function

Public Function GetFileName(ByVal Path As String) As String
    
End Function

Public Function GetFileNameWithoutExtension(ByVal Path As String) As String
    
End Function

Public Function GetFullPath(ByVal Path As String) As String
    
End Function

Public Function GetPathRoot(ByVal Path As String) As String
    
End Function

Public Function GetTempFileName() As String
    
End Function

Public Function GetTempPath() As String
    
End Function

Public Function GetInvalidPathChars() As Integer()
    
End Function

Public Function GetInvalidFileNameChars() As Integer()
    
End Function
