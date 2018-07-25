VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Delete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IAction


Private Sub IAction_proccessAction(ByVal provider As provider)
    If Not provider.isInstanceRegistered Then Exit Sub
    Dim i As Long, b As ActionBuffer, Instance As Instance
    Set Instance = provider.Instance
    Set b = Instance.Proccess.buf
    On Error GoTo err
    For i = 0 To b.Count - 1
        If Dir(b(i)) = "" Then
            Call Instance.cerr("Value Not Such File Or Directory.")
        Else
            Call checkAndDelFileFolder(b(i))
        End If
    Next
    Exit Sub
err:
    Call Instance.cerr("Some Errors Occured Message:" & err.Description)
End Sub
Private Sub checkAndDelFileFolder(Path As String)
    
End Sub

Private Function IAction_provideSettings() As miccyUltimateTools.ISettings

End Function
