VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IClassTexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "IClassTexer"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Public Enum mint_TexingType
    tt_Binary
    tt_String
End Enum

Public Function toString() As String
    throw NotImplementedException
End Function
Public Sub fromString(Value As String)
    throw NotImplementedException
End Sub
Public Function toByteArray() As Byte()
    throw NotImplementedException
End Function
Public Sub fromByteArray(Value() As Byte)
    throw NotImplementedException
End Sub
Public Property Get texHeader(texType As mint_TexingType) As String
    throw NotImplementedException
End Property
