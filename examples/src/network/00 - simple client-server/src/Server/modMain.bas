Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    Call Application.trigConsole(sdBoth)
    
#If DEVMODE Then
    Call Application.EnterDeveloperMode
#End If

    Dim t As New tApplication
    Call Application.StartApplication(t, ApplicationExecutable, App, 0)
End Sub