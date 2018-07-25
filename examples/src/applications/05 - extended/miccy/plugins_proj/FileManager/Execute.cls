VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Execute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IAction

Private Sub IAction_proccessAction(ByVal provider As miccyUltimateTools.provider)
    Dim actBuf As ActionBuffer
    Set actBuf = provider.ActionBuffer
    Dim i As Long
    For i = 0 To actBuf.Count - 1
        If (actBuf.getVarType(i) And vbString) = vbString Then
            Call TryExec(actBuf(i))
        End If
    Next
End Sub
Private Function IAction_provideSettings() As miccyUltimateTools.ISettings
    IAction_provideSettings = Nothing
End Function
Private Sub TryExec(Str As String)
    If Dir(Str) <> "" Then
        'exec here.
    End If
End Sub

