VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Response"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITargetStream
Implements IClassClonable
Implements IAliasable
Implements ISupportSignal
Implements ISupportSlot

Public Enum ResponseBufferMode
    rbmBuffer
    rbmStream
    rbmDefault = rbmStream
End Enum

Dim r_buffMode As ResponseBufferMode
Dim r_buff As String
Dim r_isBuffering As Boolean
Dim r_headers As New Headers
Dim r_enableHtmlErrorChecks As Boolean

Public Sub SendHeaders()
    
End Sub
Public Sub WriteData(Data)
    
End Sub
Public Sub Flush()
    If isBuffering Then _
        stdout buff
End Sub

Public Property Get HtmlCompression() As Boolean
    
End Property
Public Property Let HtmlCompression(Value As Boolean)
    
End Property


Public Function Request() As Request
    
End Function
Public Function Socket() As Socket
    
End Function
Public Property Get BufferMode() As ResponseBufferMode
    
End Property
Public Function Session() As Session
    
End Function
Public Function Application() As WebApplication
    
End Function

Private Property Get IAliasable_Alias() As String

End Property

Private Function IClassClonable_Clone() As Object

End Function

Private Function ISupportSignal_Emitter() As MintAPI.SignalEmitter

End Function

Private Property Get ISupportSlot_AutoManage() As Boolean

End Property

Private Sub ISupportSlot_ManualManageCall(SlotName As String)

End Sub

Private Function ISupportSlot_Methods() As String

End Function

Private Function ITargetStream_Available(direction As MintAPI.StreamDirection) As Long

End Function

Private Sub ITargetStream_CloseStream(Optional direction As MintAPI.StreamDirection = 3&)

End Sub

Private Function ITargetStream_getState(Dir As MintAPI.StreamDirection) As Boolean

End Function

Private Sub ITargetStream_inStream(Data As Variant, Optional Length As Long = -1&)

End Sub

Private Sub ITargetStream_OpenStream(Optional direction As MintAPI.StreamDirection = 3&)

End Sub

Private Sub ITargetStream_outStream(Data As Variant, Optional Length As Long = -1&)

End Sub

Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0&) As Byte()

End Function

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)

End Sub

Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)

End Sub
