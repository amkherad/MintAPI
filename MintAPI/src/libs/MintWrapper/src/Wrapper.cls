VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
    totalCreatedInstances = totalCreatedInstances + 1
End Sub


Public Sub LoadHost()
    Call InitializeLoop
End Sub
Public Sub UnloadHost()
    Call FinilizeLoop
    End
End Sub


Public Function Startup(Optional Arguments) As Startup
    Dim s As New Startup
    Call s.Initialize(Arguments)
    Set Startup = s
End Function


Public Sub RegisterVariable(Name As String, var)
    If IsObject(var) Then throw "InvalidStatusException", "var can't be object."
On Error GoTo Err
    Call MintHostAPI_Variables.Add(var, Name)
    Exit Sub
Err:
    Call Err.Raise(Err.Number, Err.Source, "A variable with the same name already registered.")
End Sub
Public Sub UnregisterVariable(Name As String)
    Call MintHostAPI_Variables.Add(Name)
End Sub
