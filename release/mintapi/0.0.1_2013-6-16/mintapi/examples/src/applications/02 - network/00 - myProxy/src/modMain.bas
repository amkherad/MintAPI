Attribute VB_Name = "modMain"
Option Explicit
Public Const PREVINSTANCEMESSAGEID_SEND As Long = 12
Public Const PREVINSTANCEMESSAGEID_RECIEVE As Long = -15300

Public trayHandle As Long

Public Sub Main()
    If App.PrevInstance Then Exit Sub
    Call Application.EnableVisualStyles(True)
    'Call Application.EnterDeveloperMode
    Dim t As New tApplication
    Call Application.StartApplication(t, ApplicationExecutable, App, 0)
End Sub

Public Sub ExitApplication()
    Call Environment.DestroyTrayIcon(trayHandle)
    Call Application.TerminateApplication
    Call Unload(frmMain)
    Call Unload(callbackHandler)
End Sub
