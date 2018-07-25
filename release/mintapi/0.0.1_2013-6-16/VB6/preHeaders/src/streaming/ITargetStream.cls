VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITargetStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@PROJECT_LICENSE
Option Explicit
Option Base 0
Const CLASSID As String = "ITargetStream"

Public Function OutStream(ByVal Data As IData) As Boolean
    throw NotImplementedException
End Function
Public Function InStream(ParamArray args() As Variant) As IData
    throw NotImplementedException
End Function

Public Function InStatus() As Boolean
    throw NotImplementedException
End Function
Public Function OutStatus() As Boolean
    throw NotImplementedException
End Function

