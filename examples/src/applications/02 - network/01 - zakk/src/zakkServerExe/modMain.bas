Attribute VB_Name = "modMain"
Option Explicit

Public lCallback As New LogCallback

Public Sub Main()
    Call Load(frmLog)
    
    Call Log("Initializing zakk server...")
    
    Call zakkServer.InitializeServer(Command$, lCallback)
    
    Call Log("zakk server initializing successfull...")
    
    If Not (IsArg("-log") Or IsArg("/nolog")) Then _
        Call frmLog.Show
End Sub

Public Function IsArg(Arg As String) As Boolean
    IsArg = (InStr(1, Command$, " " & Arg) > 0)
End Function


Public Sub Log(Message As String)
    frmLog.txt.Text = frmLog.txt.Text & Format(Now, "dd/mm/yyyy hh:MM:ss:ms") & "  " & Message & vbCrLf
End Sub
