VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@PROJECT_LICENSE
Option Explicit
Option Base 0
Const CLASSID As String = "IData"

Public Sub getEntireDataByteArray(targetArray() As Byte)
    throw NotImplementedException
End Sub
Public Sub getCopyOfEntireDataByteArray(targetArray() As Byte)
    throw NotImplementedException
End Sub

Public Sub ReadByteArray(targetArray() As Byte)
    throw NotImplementedException
End Sub
Public Sub WriteByteArray(targetArray() As Byte)
    throw NotImplementedException
End Sub

Public Sub SetData(targetData)
    throw NotImplementedException
End Sub
Public Sub ReadData(targetData)
    throw NotImplementedException
End Sub

Public Sub Defragmentation()
    throw NotImplementedException
End Sub

Public Sub Clear()
    throw NotImplementedException
End Sub

Public Property Get IsEmpty() As Boolean
    throw NotImplementedException
End Property
Public Property Get Length() As Long
    throw NotImplementedException
End Property
