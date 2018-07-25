Attribute VB_Name = "mint_callbacks"
Option Explicit

Public Sub Dummy_Void()
    
End Sub

'Public Sub CallBack_Timer_Procedure(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
'On Error GoTo ErrorHandler
'    Const API_WM_TIMER = &H113
'    Dim objTimer As Timer
'    ' Make sure that the message is WM_TIMER.
'    If uMsg = API_WM_TIMER Then
'        ' It is a timer event.
'        'Debug.Print "Timer: ", hwnd, uMsg, idEvent, dwTime
'
'        For Each objTimer In modMain_AllTimers
'            ' Execute the callback method in the class.
'            Call objTimer.HandleCallBack(idEvent)
'        Next objTimer
'    End If
'    Exit Sub
'ErrorHandler:
'    throw Exps.Exception(Err.Description)
'End Sub
'
'Public Function CallBack_Thread_Procedure(ByVal clsHandle As Long) As Long
'    Dim Method As New Method
'    'Dim Thread As New Thread
'    Call Method.Constructor0("", clsHandle)
''    Call Thread.Initialize(targetFuncHandle:=Method)
''    Call Thread.Invoke
'End Function

Public Function mint_callback_ConsoleCtrlEventHandler(ByVal CtrlType As Long) As Long
'Const CTRL_C_EVENT = 0
'Const CTRL_BREAK_EVENT = 1
'Const CTRL_CLOSE_EVENT = 2
''  3 is reserved!
''  4 is reserved!
'Const CTRL_LOGOFF_EVENT = 5
'Const CTRL_SHUTDOWN_EVENT = 6
'    If CtrlType = CTRL_C_EVENT Or CtrlType = CTRL_BREAK_EVENT Then _
'        mint_api_console_is_breaked = True
End Function
Public Function mint_callback_VectoredHandler(ByVal ExceptionInfo As Long) As Long
    
End Function
