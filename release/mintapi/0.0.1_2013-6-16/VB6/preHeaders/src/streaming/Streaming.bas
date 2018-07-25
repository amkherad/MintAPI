Attribute VB_Name = "Streaming"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "Streaming"

Public Sub out(ByVal target As ITargetStream, streamBuffer)
        OutStream target, streamBuffer
End Sub
Public Sub OutStream(ByVal target As ITargetStream, streamBuffer)
    Dim idt As IData
    If Not target.OutStatus Then throw TargetNotReadyException
    If VarType(streamBuffer) = vbObject Then
        If (TypeOf streamBuffer Is IData) Then
            Set idt = streamBuffer
        Else
            Set idt = New BinaryData
            Call idt.SetData(streamBuffer)
        End If
    Else
        Set idt = New BinaryData
        Call idt.SetData(streamBuffer)
    End If
    Call target.OutStream(idt)
End Sub
Public Sub InStream(ByVal target As ITargetStream, streamBuffer)
    Dim idt As IData
    If Not target.InStatus Then throw TargetNotReadyException
    If VarType(streamBuffer) = vbObject Then
        If (TypeOf streamBuffer Is IData) Then
            Set idt = streamBuffer
        Else
            Set idt = New BinaryData
            Call idt.SetData(streamBuffer)
        End If
    Else
        Set idt = New BinaryData
        Call idt.SetData(streamBuffer)
    End If
    Call target.InStream(idt)
End Sub
