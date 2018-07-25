VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TableData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "TableData"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements ITargetStream
Implements IAliasable


Private Type Column
    Value As Variant
End Type
Private Type Row
    cols() As Column
    colsCount As Long
End Type
Private Type Table
    colTypes() As VbVarType
    colNames() As String
    colNamesCount As Long

    rows() As Row
    rowsCount As Long
End Type

Dim static_column_types As Boolean

Private Sub Class_Initialize()
    static_column_types = True
End Sub

Public Property Get StaticColumnTypes() As Boolean
    StaticColumnTypes = static_column_types
End Property
Public Property Let StaticColumnTypes(Value As Boolean)
    static_column_types = Value
    If Not Value Then

    End If
End Property

Public Function CheckIfNotStaticColumnTypes() As Boolean

End Function

Public Sub Sort()

End Sub
Public Sub Restructure()

End Sub
Public Sub reserve()

End Sub

Private Function ITargetStream_Available(Dir As StreamDirection) As Long

End Function

Private Sub ITargetStream_CloseStream(Optional direction As StreamDirection = 3&)
    
End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean

End Function

Private Sub ITargetStream_inStream(Data As Variant, Optional Length As Long = -1&)

End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    
End Function

Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = StreamDirection.sdBoth)

End Sub

Private Sub ITargetStream_outStream(Data As Variant, Optional Length As Long = -1&)

End Sub

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved)

End Sub

Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved)

End Sub


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
