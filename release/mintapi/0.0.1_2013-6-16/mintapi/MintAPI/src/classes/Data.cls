VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "Data"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IData
Implements IClassTexer
Implements IClassClonable
Implements ITargetStream
Implements IAliasable


Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String


Private Function IClassClonable_Clone() As Object

End Function

Private Sub IClassExceptionManager_setAllExceptionsState(Optional States As Boolean = True)

End Sub

Private Sub IClassExceptionManager_setExceptionState(exp As ExceptionType, Optional State As Boolean = True)

End Sub

Private Sub IClassTexer_fromByteArray(Value() As Byte)

End Sub

Private Sub IClassTexer_fromString(Value As String)

End Sub

Private Function IClassTexer_toByteArray() As Byte()

End Function

Private Function IClassTexer_toString() As String

End Function
Private Property Get IClassTexer_texHeader(texType As mint_TexingType) As String

End Property

Private Sub IData_Append(target As Variant)

End Sub

Private Sub IData_Clear()

End Sub

Private Sub IData_CopyByteArray(targetByteArray() As Byte)

End Sub

Private Function IData_GetByteArray() As Byte()

End Function

Private Property Get IData_Length() As Long

End Property

Private Sub IData_SetByteArray(B() As Byte)

End Sub

Private Function ITargetStream_Available(direction As StreamDirection) As Long

End Function

Private Sub ITargetStream_CloseStream(Optional direction As StreamDirection = 3&)

End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean

End Function

Private Sub ITargetStream_inStream(Data As Variant, Optional Length As Long = -1&)

End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    
End Function

Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = 3&)

End Sub

Private Sub ITargetStream_outStream(Data As Variant, Optional Length As Long = -1&)

End Sub

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, inState, inState_LOCK, Reserved)
End Sub
Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, outState, outState_LOCK, Reserved)
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
