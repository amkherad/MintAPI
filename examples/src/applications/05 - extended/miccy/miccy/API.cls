VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "API"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub i_Main(args() As String)
    If App.PrevInstance Then
        If Not gp.MultiInstance Then
            throw Exps.InvalidCallException
        End If
    End If
    Set provider = New provider
    Set Iwindows = New Collection
    Call Load(Settings)
    Call mForm.Show
End Sub
Public Function CanLoadNewInstance() As Boolean
    CanLoadNewInstance = ((gp.MultiInstance) Imp (Not App.PrevInstance))
End Function
Public Sub Initialize(ByVal unloadObject As Object)
    Set callerLibrary_unloadObject = unloadObject
End Sub

Public Function Instance() As Instance
    Dim i As New Instance
    Set Instance = i
End Function
