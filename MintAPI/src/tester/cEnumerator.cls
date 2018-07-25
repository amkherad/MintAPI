VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEnumerator

Dim X As Long

Public Property Get Current() As Variant
    Current = X
End Property
Public Function MoveNext() As Boolean
    MoveNext = IEnumerator_MoveNext
End Function
Private Sub Reset()
    X = 0
End Sub

Private Property Get IEnumerator_Current() As Variant
    VariantCopy IEnumerator_Current, Current
End Property

Private Function IEnumerator_MoveNext() As Boolean
    If X < 5 Then
        IEnumerator_MoveNext = True
        X = X + 1
    End If
End Function

Private Sub IEnumerator_Reset()
    Reset
End Sub

Public Function NewEnum() As stdole.IUnknown
    Set [NewEnum] = Enumerator.CreateEnumerator(Me)
End Function

