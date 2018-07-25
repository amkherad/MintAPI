Attribute VB_Name = "modMain"
Option Explicit
Public Sub Main()
    'Call Application.EnableVisualStyles(True)
'    Call Application.trigConsole(sdBoth)
    Dim t As New tApplication
    Call Application.Run(t, ApplicationExecutable Or OwnConsole, App, 0)
End Sub
