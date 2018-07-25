Attribute VB_Name = "Module1"
Public Const CTF_COINIT = &H8
Public Const CTF_INSIST = &H1
Public Const CTF_PROCESS_REF = &H4
Public Const CTF_THREAD_REF = &H2

Declare Function SHCreateThread Lib "shlwapi.dll" (ByVal pfnThreadProc As Long, pData As Any, ByVal dwFlags As Long, ByVal pfnCallback As Long) As Long

Public stopThread As Boolean
Public beginTime As Long


'Our thread
Public Sub myThread()
    'Set start time...
    beginTime = Timer
    
    'This would normally lock the form
    Do While Not stopThread
        'blah
        'blah
        'blah
        MsgBox "test"
    Loop
    
    'Lets see how long the thread was running...
    MsgBox "The thread ran for: " & Timer - beginTime
    
End Sub
