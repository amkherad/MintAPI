Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    Call Application.trigConsole(sdBoth)

    Dim t As New tApplication
    Call Application.StartApplication(t, ApplicationExecutable, App, 0)
End Sub
