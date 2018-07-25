VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEnumerable

Private Function IEnumerable_GetEnumerator() As MintAPI.IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function
Private Function GetEnumerator() As MintAPI.IEnumerator
    Set GetEnumerator = New cEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = Enumerator.CreateEnumerator(GetEnumerator)
End Function
Public Function NewEnum() As stdole.IUnknown
    Set [NewEnum] = Enumerator.CreateEnumerator(GetEnumerator)
End Function

