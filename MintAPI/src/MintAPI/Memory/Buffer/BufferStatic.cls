VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BufferStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ClassStreamBuffer"
'@PROJECT_LICENSE

''<class name="BufferStatic" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "BufferStatic" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Public Function BinaryBuffer(Expression, Optional ByVal Length As Long = -1) As ByteArray
Attribute BinaryBuffer.VB_UserMemId = 0
    Set BinaryBuffer = New ByteArray
    Call BinaryBuffer.Constructor1(Expression, True, Length)
End Function

Public Function ClassStreamBuffer(ByVal BufferType As ClassStreamBufferType, Optional ByRef Value As Variant, Optional ByVal Length As Long = -1) As ClassStreamBuffer
    Set ClassStreamBuffer = New ClassStreamBuffer
    If (BufferType And csbtTypeMask) = csbtData Then
        If IsObject(Value) Then
            If TypeOf Value Is ClassStreamBuffer Then
                Set ClassStreamBuffer = Value
                Exit Function
            End If
        End If
        Call ClassStreamBuffer.Constructor0(BufferType, Value, Length)
    Else
        Call ClassStreamBuffer.Constructor0(BufferType, vbNullPtr, 0)
    End If
End Function
Public Function ClassStreamBufferV(ByRef Value As Variant, Optional ByVal Length As Long = -1) As ClassStreamBuffer
    If IsObject(Value) Then
        If TypeOf Value Is ClassStreamBuffer Then
            Set ClassStreamBufferV = Value
            Exit Function
        End If
    End If
    Set ClassStreamBufferV = New ClassStreamBuffer
    Call ClassStreamBufferV.Constructor1(Value, Length)
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

'All your methods here...

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
