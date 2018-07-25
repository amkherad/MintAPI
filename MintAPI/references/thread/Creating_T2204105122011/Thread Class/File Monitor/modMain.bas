Attribute VB_Name = "Monitor"
Option Explicit

'File monitoring functions
Private Declare Function FindNextChangeNotification Lib "kernel32.dll" (ByVal hChangeHandle As Long) As Long
Private Declare Function FindFirstChangeNotification Lib "kernel32.dll" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As Long
Private Declare Function FindCloseChangeNotification Lib "kernel32.dll" (ByVal hChangeHandle As Long) As Long

'Wait Functions
Private Declare Function CreateEvent Lib "kernel32.dll" Alias "CreateEventA" (ByRef lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32.dll" (ByVal nCount As Long, ByRef lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function SetEvent Lib "kernel32.dll" (ByVal hEvent As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long


'Structures
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type


'Private Const INFINITE As Long = &HFFFFFFFF
'Private Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4
'Private Const FILE_NOTIFY_CHANGE_CREATION As Long = &H40
'Private Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2
'Private Const FILE_NOTIFY_CHANGE_FILE_NAME As Long = &H1
'Private Const FILE_NOTIFY_CHANGE_LAST_ACCESS As Long = &H20
'Private Const FILE_NOTIFY_CHANGE_LAST_WRITE As Long = &H10
'Private Const FILE_NOTIFY_CHANGE_SECURITY As Long = &H100
'Private Const FILE_NOTIFY_CHANGE_SIZE As Long = &H8

'This ENUM is derived from above constants
Public Enum NotificationFlags
Attributes = 4
Creation = &H40
Dir_Name = 2
File_Name = 1
Last_Access = &H20
Last_Write = &H10
Size = &H8
End Enum

Public Handle As Long, THandle As Long, EHandle As Long
Public WaitValue As Long
Public Handles(1) As Long
Public MemAddress As Long, Tlsindex As Long, TlsAddress As Long

Public Sub WaitForEventEx(Han As Long)
'We will set the address of memory we created in the Tlsindex retrieved from the
'MSVBVM60.DLL. VB will use this address to store our local variables.
TlsSetValue Tlsindex, ByVal MemAddress

While True
    Res = WaitForMultipleObjects(2, Handles(0), 0, &HFFFFFFFF)
 If Res = 0 Then
    Change_Occured                                       'Change occured in the specified
    FindNextChangeNotification Handle                    'Directory
 ElseIf Res = 1 Then
    Form1.ADDtext "Thread closed (Monitor disposed)"      'The event was signaled
    Exit Sub                                              'Exit this thread
End If
Wend
End Sub


Public Function AddMonitor(ByVal path As String, WatchSubtree As Boolean, Flags As NotificationFlags) As Integer
'Here we create a handle using FindFirstChangeNotification and then create an event
'using CreateEvent. And we allocate a memory for a new thread which will later
'be set in the Tls of the thread. The Tls index that VB creates was stored in the
'memory location &H6610EE7C in my version of MSVBVM60. So it may differ in another
'versions. If it differs the VB IDE crashes. Then we create thread which will be
'suspended and only run after a call to Start.

Handle = FindFirstChangeNotification(path, WatchSubtree, Flags) 'Create wait handle
Form1.ADDtext "Monitor Handle created ==> " & Handle
If Handle <> -1 Then
Dim Y As SECURITY_ATTRIBUTES
EHandle = CreateEvent(Y, 0, 0, "Exit_Thread")                     'Create Event
Handles(0) = Handle
Handles(1) = EHandle
WaitValue = 1000

Form1.ADDtext "Retrieved TlsIndex ==> " & Str(Tlsindex)
THandle = CreateThread(Y, 0, AddressOf WaitForEventEx, ByVal Handle, 4, Tid)
Form1.ADDtext "New Thread Created with ID " & Hex(THandle)
AddMonitor = 1
Else
Form1.ADDtext "ERROR: Monitor handle couldn't be created. IS any checkbox checked?"
End If                                                             'Create a new thread
End Function

Public Sub DisposeMonitor()
If SetEvent(EHandle) = 0 Then Exit Sub                              'Fire the event
Sleep 100
FindCloseChangeNotification Handle                                  'Close wait handle
CloseHandle EHandle                                                 'Close event handle
Form1.Text1.Text = ""
End Sub

Public Sub Start()
ResumeThread THandle                                                 'Resume thread
End Sub

Public Sub Change_Occured()
'Here you decide what to do, I will add a log message in Form1.
Form1.ADDtext "Change occured"
End Sub
