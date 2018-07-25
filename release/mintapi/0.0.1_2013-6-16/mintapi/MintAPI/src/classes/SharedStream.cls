VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SharedStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "SharedStream"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements ITargetStream
Implements IAliasable


Friend Sub Initialize(Name As String)

End Sub

Private Function ITargetStream_Available(Dir As StreamDirection) As Long

End Function

Private Sub ITargetStream_CloseStream(Optional direction As StreamDirection = 3&)

End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean
    ITargetStream_getState = True
End Function

Private Sub ITargetStream_inStream(Data, Optional Length As Long = -1&)

End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    
End Function

Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = StreamDirection.sdBoth)

End Sub

Private Sub ITargetStream_outStream(Data, Optional Length As Long = -1&)

End Sub

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved)

End Sub

Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved)

End Sub


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
