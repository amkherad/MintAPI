Attribute VB_Name = "modMain"
Option Explicit

Public totalCreatedInstances As Long

Dim MintAPI As Object
Dim ComArg As String
Dim instanceCounter As Long

Public MintHostAPI_Variables As New Collection

Public Sub Main()
    ComArg = Command$
    If IsCommandIncludes("startup") Then
        Call StartupActions
    End If
    If IsCommandIncludes("stay") Then _
        Call InitializeLoop
End Sub
Public Function IsCommandIncludes(strCheck As String) As Boolean
    IsCommandIncludes = (InStr(1, ComArg, " /" & strCheck) > 0)
End Function

Public Sub InitializeLoop()
    If instanceCounter = 0 Then _
        Call Load(frmStay)
    instanceCounter = instanceCounter + 1
End Sub
Public Sub FinilizeLoop()
    instanceCounter = instanceCounter - 1
    If instanceCounter = 0 Then _
    Call Unload(frmStay)
End Sub

Public Sub StartupActions()
    
End Sub

Public Function API() As Object
On Error GoTo Err
    If MintAPI Is Nothing Then _
        Set MintAPI = CreateObject("MintAPI.API")
    API = MintAPI
Err:
End Function
Public Sub throw(exp As String, Optional arg As String)
    
End Sub
Public Function MintRegister() As Object
    
End Function
