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
    throw Exps.NotImplementedException
End Property
Public Sub BoundByteArray(B() As Byte)
    throw Exps.NotImplementedException
End Sub
Public Sub CloneByteArray(B() As Byte)
    throw Exps.NotImplementedException
End Sub
Public Function GetByteArray() As Byte()
    throw Exps.NotImplementedException
End Function
Public Sub Append(target)
    throw Exps.NotImplementedException
End Sub
Public Sub Clear()
    throw Exps.NotImplementedException
End Sub
