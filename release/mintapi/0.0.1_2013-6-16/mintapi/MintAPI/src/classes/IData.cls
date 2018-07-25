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
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "IData"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Public Enum DataType
    BinaryData
    SequentialData
    TableData
End Enum

Public Property Get Length() As Long
    throw NotImplementedException
End Property
Public Sub SetByteArray(B() As Byte)
    throw NotImplementedException
End Sub
Public Sub CopyByteArray(targetByteArray() As Byte)
    throw NotImplementedException
End Sub
Public Function GetByteArray() As Byte()
    throw NotImplementedException
End Function
Public Sub Append(target)
    throw NotImplementedException
End Sub
Public Sub Clear()
    throw NotImplementedException
End Sub
