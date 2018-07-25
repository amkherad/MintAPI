Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    'Call Application.EnableVisualStyles(True)
    'Call Application.EnterDeveloperMode
    Dim t As New tApplication
    Call Application.StartApplication(t, ApplicationExecutable, App, 0)
End Sub
