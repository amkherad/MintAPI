VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectoryWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'********************************************APIS****************************************
Private Declare Function ReadDirectoryChanges Lib "kernel32.dll" Alias "ReadDirectoryChangesW" (ByVal hDirectory As Long, ByVal lpBuffer As Long, ByVal nBufferLength As Long, ByVal bWatchSubTree As Long, ByVal dwNotifyFiler As Long, ByVal lpBytesReturned As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityATTRIBUTES As SECURITY_ATTRIBUTES, ByVal dwCREATIONDisposition As Long, ByVal dwFlagsAndATTRIBUTES As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateEvent Lib "kernel32.dll" Alias "CreateEventA" (ByRef lpEventATTRIBUTES As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'*******************************************EVENTS***************************************
Public Event OnCreated(ByVal Name As String)
Public Event OnDeleted(ByVal Name As String)
Public Event OnRenamed(ByVal OldName As String, ByVal NewName As String)
Public Event OnChanged(ByVal Name As String)
Public Event OnStopped()

'*****************************************EVENT VARIABLES********************************
Private ChangeType As Long
Private OldName As String
Private NewName As String

'*****************************************TYPES******************************************
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Private Type FILE_NOTIFY_INFORMATION
    NextEntryOffSet As Long
    Action As Long
    FileNameLength As Long
    FileName(255 - 1) As Byte
End Type

Private Type MonitorParams
    CurrentDirectory As String
    bWatchSubTree As Boolean
    Flags As NotificationFlags
    ThreadParams As Long
    DirectoryHandle As Long
    hEvent As Long
End Type
'************************************************ENUM***********************************
Public Enum NotificationFlags
    FILEATTRIBUTES = 4
    CREATION = &H40
    DIR_NAME = 2
    FILE_NAME = 1
    LAST_ACCESS = &H20
    LAST_WRITE = &H10
    FILESIZE = &H8
End Enum
'************************************************CONSTANTS********************************
Private Const FILE_LIST_DIRECTORY As Long = &H1
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_FLAG_BACKUP_SEMANTICS As Long = &H2000000 'for opening a directory
Private Const FILE_FLAG_OVERLAPPED = &H40000000

Private Const FILE_ACTION_ADDED As Long = &H1
Private Const FILE_ACTION_MODIFIED As Long = &H3
Private Const FILE_ACTION_REMOVED As Long = &H2
Private Const FILE_ACTION_RENAMED_NEW_NAME As Long = &H5
Private Const FILE_ACTION_RENAMED_OLD_NAME As Long = &H4

'********************************************PRIVATE VARIABLES***************************
Dim Res As Long, DirectoryHandle As Long
Dim SA As SECURITY_ATTRIBUTES, FNI As FILE_NOTIFY_INFORMATION
Private MP As MonitorParams
Private WithEvents WatchingThread As Thread     'The thread
Attribute WatchingThread.VB_VarHelpID = -1


Public Function CreateMonitor(ByVal MonitorParent As Form, ByVal Path As String, ByVal bWatchSubTree As Boolean, ByVal nFlags As NotificationFlags) As Integer
    If nFlags = 0 Then
        Exit Function
    End If
    If MP.DirectoryHandle <> 0 Then Exit Function   'Already Created
                                                    'Create Directory Handle
    DirectoryHandle = CreateFile(Path, FILE_LIST_DIRECTORY, FILE_SHARE_READ _
                        Or FILE_SHARE_WRITE, SA, OPEN_EXISTING, _
                        FILE_FLAG_BACKUP_SEMANTICS Or FILE_FLAG_OVERLAPPED, 0)
                                                    'Set Parametres
    MP.CurrentDirectory = Path
    MP.bWatchSubTree = IIf(bWatchSubTree, 1, 0)
    MP.Flags = nFlags
    MP.DirectoryHandle = DirectoryHandle
    MP.hEvent = CreateEvent(SA, 0, 0, "Completion Event")
                                                    'Start Thread. Now follow DoWork Event
    WatchingThread.StartWorkAsync MonitorParent, DirectoryHandle
    'Here passing DirectoryHandle was not necessary, it could've been accessed by
    'MP.DirectoryHandle but well just showing how to pass a parametre to the Thread
End Function

Public Sub ResumeMonitor()
    WatchingThread.ResumeWork
End Sub

Public Sub PauseMonitor()
    WatchingThread.PauseWork
End Sub

Public Sub StopMonitor()
    If WatchingThread.IsBusy Then WatchingThread.CancelWork
End Sub

Public Function GetFullPath(ByVal FileName As String) As String
    GetFullPath = MP.CurrentDirectory & "\" & FileName
End Function

Private Sub Class_Initialize()
    Set WatchingThread = New Thread
End Sub

Private Sub Class_Terminate()
    If WatchingThread.WorkerStatus <> TERMINATED Then Me.StopMonitor
    Set WatchingThread = Nothing
End Sub

Private Sub WatchingThread_DoWork(ByVal Arg As Variant, Result As Variant)
Dim FName As String, bBytesReturned As Long, Pos As Long, WaitResult As Long
Static LogString As String
Dim Buffer(1024 * 2 - 1) As Byte        'buffer size can be increased.
Dim Y As OVERLAPPED
FName = Space$(255)
Y.hEvent = MP.hEvent
Start:
    Res = ReadDirectoryChanges(CLng(Arg), VarPtr(Buffer(0)), 1024 * 2, _
                        MP.bWatchSubTree, MP.Flags, VarPtr(bBytesReturned), Y, 0)
WaitAgain:
                      
While WatchingThread.CancellationPending
    'WaitResult = WaitForMultipleObjects(2, MP.hEvents(0), 0, INFINITE)
    WaitResult = WaitForSingleObject(MP.hEvent, 10)
 Select Case WaitResult
  Case &H102            'Wait Timeout
        GoTo WaitAgain:
  Case 0            'A change has occured
      While True
        CopyMemory FNI, Buffer(Pos), Len(FNI)
        
        'Extract file name from the FNI.Filename array
        FName = FNI.FileName
        FName = Left(FName, FNI.FileNameLength \ 2)
        
        ChangeType = FNI.Action
        
        If FNI.Action <> FILE_ACTION_RENAMED_NEW_NAME Then
            OldName = FName
        Else
            NewName = FName
        End If
        WatchingThread.ReportProgress 0
        
        Pos = Pos + FNI.NextEntryOffSet
        
        If FNI.NextEntryOffSet = 0 Then
            Pos = 0
            GoTo Start:
        End If
    Wend
  
  Case Else
        
  End Select
Wend

CloseHandle MP.hEvent
CloseHandle MP.DirectoryHandle
MP.DirectoryHandle = 0
End Sub

Private Sub WatchingThread_ProgressChanged(ByVal ProgressPercent As Long)
    Dim Args As NotificationFlags
        Select Case ChangeType
            Case FILE_ACTION_ADDED
                RaiseEvent OnCreated(OldName)
            Case FILE_ACTION_REMOVED
                RaiseEvent OnDeleted(OldName)
            Case FILE_ACTION_MODIFIED
                RaiseEvent OnChanged(OldName)
            Case FILE_ACTION_RENAMED_NEW_NAME
                RaiseEvent OnRenamed(OldName, NewName)
            Case FILE_ACTION_RENAMED_OLD_NAME
        End Select
End Sub

Private Sub WatchingThread_WorkCompleted(ByVal Result As Variant)
    RaiseEvent OnStopped
End Sub
