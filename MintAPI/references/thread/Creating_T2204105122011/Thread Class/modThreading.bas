Attribute VB_Name = "modThreading"
'VB6 is not threadsafe. So you must be careful while using this module. It will be
'a lot better if you first read the Readme file provided with this module before
'continuing your work. While there may be other ways to create threads in VB6, I
'prefer this way, because this is a lot comfortable with me due to direct approach.
'But this may not apply in your case. If you don't want to take even a slightest risk
'then you would rather search something else in Internet.

'PRO_SCIENCE108
Option Explicit

'Thread Functions
Private Declare Function CreateThread Lib "kernel32.dll" (ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByRef lpParameter As Any, ByVal dwCreationFlags As Long, ByRef lpThreadId As Long) As Long
Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib "KERNEL32" (ByVal hThread As Long) As Long
Private Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function ReleaseSemaphore Lib "KERNEL32" (ByVal hSemaphore As Long, ByVal lReleaseCount As Long, lpPreviousCount As Long) As Long
Private Declare Function CreateSemaphore Lib "KERNEL32" Alias "CreateSemaphoreA" (lpSemaphoreAttributes As Long, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As Long

'Memory Functions
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function WriteProcessMemory Lib "KERNEL32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
'Windows Function
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetCurrentThreadId Lib "KERNEL32" () As Long
Private Declare Function OpenThread Lib "KERNEL32" (ByVal AccessRights As Long, ByVal Inherit As Long, ByVal ID As Long) As Long

'Constants
Private Const WM_USER = &H400
Public Const WM_PROGRESS = WM_USER + &HFF
Private Const CREATE_SUSPENDED As Long = &H4
Public Const THREADS_MAX = 10               'The number of maximum threads. Can be incremented or decremented

Private Const WAIT_TIMEOUT = &H102
Private Const INFINITE = -1
Public Type CRITICAL_SECTION

    dummy As Long
End Type

'Structures
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type ASMFunction
    Counter As Long
    Buffer(255) As Byte
End Type

'Public Variables.
Public TotalThreads As Long         'Total threads created
Public FormWndProc As Long          'The default WindowProc of a form
Private SemaphoreHandle As Long

'Private Variables
Private startThreadFS18 As Long     'Main thread's TIB address
Private Tlsdata(256) As Byte        'Buffer to store main thread's TLS data.



'Create threads.
Public Function CreateNewThread(ByVal ThreadProcedure As Long, ByVal CreateSuspended As _
                                Boolean, Optional ByVal Param As Long = 0) As Long
                                

Dim Y As SECURITY_ATTRIBUTES, CreationFlags As Long
Dim THandle As Long, Tid As Long

CopyMemory Tlsdata(0), ByVal (ReadFS18 + &HE10), 256

'Initialize FS18 reading and writing procedures
WriteFS18 ReadFS18
'Read thread's TIB start address. This address will be used in the InitThread
'procedure to copy whole TLS data to the created thread's TLS.
If startThreadFS18 = 0 Then _
                startThreadFS18 = ReadFS18

InitializeLock

CreationFlags = IIf(CreateSuspended, CREATE_SUSPENDED, 0)
THandle = CreateThread(Y, 0&, ThreadProcedure, ByVal Param, _
                                     CreationFlags, Tid)      'Create thread.

CreateNewThread = THandle
End Function


'Thread's entry point.
Public Function ThreadProc(ByVal objThread As Thread) As Long
    'MyFS18 = modThreading.PersonateAndReturnFS18
    
    Dim Y As Long
    
    Y = modThreading.InitThread
    objThread.RaiseWorkerEvents         'Raise the DoWork Event
    ZeroMem Y, 256                      'After returning Zero the TLS Block
    ThreadProc = 0                      'Return Exit code.
    
    'modThreading.StopImPersonationThread MyFS18
End Function

Public Function InitThread() As Long
'Copy main thread's TLS Data into our TLS. And return TLS Address i.e. FS[E10]
    CopyMemory ByVal (ReadFS18 + &HE10), Tlsdata(0), 256
    InitThread = ReadFS18 + &HE10
End Function

Public Sub StartThread(ByVal ThreadHandle As Long)
    ResumeThread ThreadHandle
End Sub

Public Sub StopThread(ByVal ThreadHandle As Long)
    SuspendThread ThreadHandle
End Sub

Private Function InitializeLock()
    If Not SemaphoreHandle Then
        SemaphoreHandle = CreateSemaphore(ByVal 0, 1, 1, "Default Semaphore")
        If Not SemaphoreHandle Then
            'Semaphore creation failed. Log the results or do something else
        End If
    End If
End Function

Public Function TryAcquireLock() As Long
   If SemaphoreHandle = 0 Then Exit Function
   TryAcquireLock = IIf(WaitForSingleObject(SemaphoreHandle, 0) = WAIT_TIMEOUT, 0, 1)
End Function

Public Function AcquireLock()
    If SemaphoreHandle = 0 Then Exit Function
    WaitForSingleObject SemaphoreHandle, INFINITE
End Function

Public Function ReleaseLock()
    If SemaphoreHandle = 0 Then Exit Function
    ReleaseSemaphore SemaphoreHandle, 1, ByVal 0
End Function

Public Sub AssignSafely(LHS As Variant, ByVal RHS As Variant)
    AcquireLock
        LHS = RHS
    ReleaseLock
End Sub

'******************************************MESSAGING ROUTINE****************************

Public Function NewWndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long _
                            , ByVal lParam As Long) As Long
'Our thread contacts the main thread via this routine.

    If Msg = WM_PROGRESS Then
'If message is our thread's then
'Must use CallWindowProc because VB won't allow to compile.lParam is a pointer to
'thread object but VB won't know it.
        CallWindowProc AddressOf ThreadWndProc, hwnd, Msg, wParam, lParam
        NewWndProc = 0
        Exit Function
    End If
     NewWndProc = CallWindowProc(FormWndProc, hwnd, Msg, wParam, lParam)
End Function

Private Function ThreadWndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long _
                            , ByVal lParam As Thread) As Long
 'Raise the Event. This event will be either ProgressChanged Event or Work Completed event.
    lParam.RaiseWorkerEvents
End Function

'**********************************************ASSEMBLY SECTION**************************

'These are assembly functions. They basically provide us the data and methods that we
'couldn't directly do via VB or the APIs.

Public Function PersonateAndReturnFS18() As Long
    PersonateAndReturnFS18 = ReadFS18
    WriteFS18 startThreadFS18
End Function

Public Sub StopImPersonationThread(ByVal originalFS18 As Long)
    WriteFS18 originalFS18
End Sub

'Read the Thread's TIB's Linear address
Private Function ReadFS18() As Long
Static IsAsmLoaded As Long, SubFS18 As ASMFunction
'Write the assembly only once. It saves time and prevents error.

'MOV EAX, FS:[18]
'RET

If IsAsmLoaded = False Then WriteAssembly VarPtr(SubFS18.Buffer(0)), "64A118000000C21000"

If IsAsmLoaded = False Then WriteAssembly AddressOf DummyFunction1, "64A118000000C3"
IsAsmLoaded = True
ReadFS18 = DummyFunction1
If ReadFS18 = 0 Then _
    ReadFS18 = CallWindowProc(VarPtr(SubFS18.Buffer(0)), 0, 0, 0, 0)
End Function

'Replace the Thread's TIB's Linear address
Public Function WriteFS18(ByVal lData As Long)
Static IsAsmLoaded As Long

'Mov EAX, SS:[ESP+4]
'Mov FS:[18], EAX
'Ret 4
If IsAsmLoaded = False Then WriteAssembly AddressOf DummyFunction2, "8B44240464A318000000C20400"
IsAsmLoaded = True
WriteFS18 = DummyFunction2(lData)
End Function

'Zeros the given memory. Couldn't use API cause it would have caused crash.
Public Function ZeroMem(ByVal Dest As Long, ByVal noOfBytes As Long)
Static IsAsmLoaded As Long

'PUSH EDI
'XOR Al, AL
'MOV EDI, ES:[ESP+8]
'MOV ECX, SS:[ESP+C]
'CLD
'REP STOS BYTE PTR ES:[EDI]
'POP EDI
'RET 8

If IsAsmLoaded = False Then WriteAssembly AddressOf DummyFunction3, "5732C0268B7C24088B4C240CFCF3AA5FC20800"

IsAsmLoaded = True
DummyFunction3 Dest, noOfBytes
End Function

'Writes the assembly code at given address
Private Sub WriteAssembly(ByVal lAddress As Long, ByVal sData As String)
Dim Bytes() As Byte, i As Integer
ReDim Bytes(Len(sData) / 2)

For i = 0 To UBound(Bytes) - 1
    Bytes(i) = Val("&H" & Mid(sData, i * 2 + 1, 2))
Next
WriteProcessMemory -1, ByVal lAddress, Bytes(0), UBound(Bytes), 0
End Sub


'These are dummy functions. These functions will be replaced by the assembly codes
Private Function DummyFunction2(ByVal lData As Long) As Long
'This is all bogus
Dim i
On Error Resume Next
For i = 1 To 10
Next
End Function

Private Function DummyFunction1() As Long
'This is all bogus
Dim i
On Error Resume Next
For i = 1 To 10
Next
End Function
Private Function DummyFunction3(ByVal Dest As Long, ByVal noOfBytes As Long) As Long
'This is all bogus
Dim i
On Error Resume Next
For i = 1 To 10
Next
End Function




                            



