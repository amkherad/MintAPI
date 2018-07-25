VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "masterProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISettings

Public Function getInclude() As String()
    Dim args() As String
    
    getInclude = args()
End Function
Public Function getExclude() As String()
    Dim args() As String
    
    getInclude = args()
End Function

Public Function SubFolders() As Boolean
    SubFolders = False
End Function
