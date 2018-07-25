Attribute VB_Name = "modMain"
Option Explicit

'commands:
'about                          about
'help                           help about
'install
'uninstall
'startup
'mount
'unmount
Public Declare Function API_ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function API_WaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Const STRINGTABLEID_ABOUT As Long = 100
Public Const STRINGTABLEID_HELP As Long = 101
Public Const STRINGTABLEID_LICENSE As Long = 103

Public Sub Main()
    Call Application.trigConsole(sdBoth)
    Dim ContinueLoad As Boolean
    Dim appCommandArguments As String
    
    appCommandArguments = VBA.Command$
    
    'break on install and uninstall to prevent load assemblies.
    If appCommandArguments = "install" Then
        
        ContinueLoad = False
    ElseIf appCommandArguments = "uninstall" Then
        ContinueLoad = False
    Else
        ContinueLoad = True
    End If
    
    If ContinueLoad Then
        If Not SpecialMethods.RegisterLibraryLicense("", "") Then throw Exceptions.Exception("Unable to register application license.")
        Dim tApp As New tApplication
        Call Application.StartApplication(tApp, ApplicationExecutable, App, VBA.Command$)
    End If
End Sub


Public Function GetStringTable(ID As Long) As String
    Dim str As String
    str = LoadResString(ID)
    str = Replace(str, "\n", vbCrLf)
    str = Replace(str, "\t", vbTab)
    GetStringTable = str
End Function
