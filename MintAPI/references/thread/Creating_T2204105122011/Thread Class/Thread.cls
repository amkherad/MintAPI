VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Thread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Synchronization

'Events (MT) = Executed in main thread. (CT) = Executed in created thread
Public Event DoWork(ByVal Arg As Variant, Result As Variant)    '(CT) The main event.
Public Event ProgressChanged(ByVal ProgressPercent As Long)     '(MT) Raised when ReportProgress is called.
Public Event WorkCompleted(ByVal Result As Variant)             '(MT) Raised when DoWork event completes

'Private Variables. Name explains all
Private Status As Thread_Status
Private DidUserCancel As Boolean
Private ParentThreadId As Long
Private PrivateThreadHandle As Long
Private ParentForm As Form
Private ParentFormHwnd As Long
Private Progress As Long

'These booleans control the events.
Private RaiseProgressFlag As Boolean
Private RaiseDoWorkFlag As Boolean
Private RaiseCompletedFlag As Boolean

Private ThreadParam As Variant

Public Enum Thread_Status
    NONE = 0
    SUSPENDED
    RUNNING
    TERMINATED
End Enum




'***********************************PUBLIC FUNCTIONS***********************************

'The main method of this class.
Public Sub StartWorkAsync(ByVal ThreadParent As Form, ByVal lParam As Variant)
        
'logText "Entering StartWorkAsync"
If IsBusy Then
    'Thread already created.
    'logText "IsBusy True, Leaving Sub Procedure."
    Exit Sub
End If

'logText "Resetting all variables"
'Reset all variables first
ResetVariables

'Set thread variables.
RaiseDoWorkFlag = True          'Imp: if this statement is missed the DoWork will never be
Status = RUNNING                'executed making all our efforts futile.
Set ParentForm = ThreadParent

ParentFormHwnd = ParentForm.hwnd
ThreadParam = lParam

'logText "Going to create a thread"

SubClass False

'Create a new thread with the specified parametres
PrivateThreadHandle = modThreading.CreateNewThread(AddressOf ThreadProc, _
                                   False, ObjPtr(Me))

If PrivateThreadHandle = 0 Then
    'Error, Find more details.
    'logText "Error while creating thread. Leaving sub procedure."
    
    GoTo error:
    
End If

'logText "Thread Created. Thread handle = " & Str(PrivateThreadHandle)
Exit Sub

error:
'Deal with problems
End Sub


Public Function IsBusy() As Boolean
    If Status = RUNNING Or Status = SUSPENDED Then IsBusy = True
    'logText "IsBusy called. Returned " & IIf(Status = RUNNING Or Status = SUSPENDED, "True", "False")
End Function

Public Sub ResumeWork()
Attribute ResumeWork.VB_Description = "Resumes the suspended thread"

    If PrivateThreadHandle <> 0 Then _
        modThreading.StartThread PrivateThreadHandle
    Status = RUNNING
    'logText "ResumeWork called. Thread handle = " & Str(PrivateThreadHandle)
End Sub

Public Sub PauseWork()
Attribute PauseWork.VB_Description = "Suspends the thread."
    If PrivateThreadHandle <> 0 Then _
        modThreading.StopThread PrivateThreadHandle
    Status = SUSPENDED
    'logText "PauseWork called. Thread handle = " & Str(PrivateThreadHandle)
End Sub

Public Sub CancelWork()
Attribute CancelWork.VB_Description = "Requests cancellation of work."
    'logText "CancelWork Called."
    DidUserCancel = True
End Sub

Public Sub ReportProgress(ByVal ProgressPercent As Long, Optional ByVal ReportAsync As Boolean = False)
Attribute ReportProgress.VB_Description = "Method that causes the worker to raise ProgressChanged event"
    'logText "ReportProgress called. Progresspercent = " & Str(ProgressPercent)
    
    RaiseProgressFlag = True
    Progress = ProgressPercent
    SendMsg WM_PROGRESS, ReportAsync
    RaiseProgressFlag = False
End Sub


'*******************************************PROPERTIES********************************
Public Property Get Param() As Variant
    Param = ThreadParam
End Property


Public Property Get Handle() As Long
    Handle = PrivateThreadHandle
End Property

Public Property Get CancellationPending() As Boolean
    CancellationPending = Not DidUserCancel
End Property

Public Property Get WorkerStatus() As Thread_Status
    WorkerStatus = Status
End Property

'********************************************PRIVATE FUNCTIONS**************************
Private Sub ResetVariables()
'logText "Resetting Variables"
    Status = NONE
    ParentFormHwnd = 0
    PrivateThreadHandle = 0
    ThreadParam = CVar(0)
    Progress = 0
    RaiseProgressFlag = False
    RaiseDoWorkFlag = False
    RaiseCompletedFlag = False
    DidUserCancel = False
End Sub

'Sends message to the main thread
Private Sub SendMsg(ByVal Msg As Long, Optional ByVal bAsync As Boolean = False)
    'logText "SendMsg Entered. Sending Message Now"
    If bAsync Then _
        PostMessage ParentFormHwnd, Msg, 0, ObjPtr(Me) _
    Else _
           SendMessage ParentFormHwnd, Msg, 0, ByVal ObjPtr(Me)
    'logText "Returning From SendMessage."
End Sub

'*********************************************EVENTS**********************************
'Even though this method is declared public but it is intended to be used only by the
'modThreading module.

Public Sub RaiseWorkerEvents()
Static Res As Variant
'Three booleans control the execution of this method. The Reason for this rather
'complex approach is that, we can't raise events from outside of the object
'Another approach would have been to declare three different public methods
'i.e. RaiseDoWork, RaiseProgress and RaiseCompleted. But since they had to be public,
'it was possible to raise the events manually which was undesirable. Even though
'this method is public, manual raising of events is not possible.

'logText "RaiseWorkerEvents Entered."

    If RaiseProgressFlag Then
        'logText "Raising ProgressChanged Event. Progress = " & CStr(Progress)
        RaiseEvent ProgressChanged(Progress)
    End If
    
    If RaiseDoWorkFlag Then
       
        RaiseDoWorkFlag = False
        'logText "Raising DoWork Event. Param = " & CStr(PAram)
        
        RaiseEvent DoWork(ThreadParam, Res)
        
        'logText "DoWork Event Done. Res = " & CStr(Res)
        
        'After the do work returns we raise the WorkCompleted event.
        'But not from this thread
        RaiseCompletedFlag = True
        SendMsg WM_PROGRESS
        RaiseCompletedFlag = False
        
        'Now the thread ends.
        Status = TERMINATED
        
    End If
    If RaiseCompletedFlag Then
        'logText "Raising RaiseCompletedFlag Event. Res = " & CStr(Res)
        RaiseEvent WorkCompleted(Res)
        SubClass True
        Res = 0
       
    End If
End Sub

'*******************************************LOGGING MODULE*******************************
'You can uncomment all the logging statements and this method to log this object's
'activities. This will help you to debug your program.


'Private Sub logText(ByVal logText As String)
'Open "Thread Log.log" For Append As #9
'Print #9, logText
'Close #9
'End Sub
'****************************************************************************************

Private Sub SubClass(ByVal bUnSubClass As Boolean)
'This routine subclasses and unsubclasses the Parent Form

If bUnSubClass Then
    'Unsubclass only if the calling thread is the last "CREATED" thread.
    
    'NOTE:-There could've been a better place to increment and decrement the
    'TotalThreads instead of this procedure, like maybe the start and end of the
    'modThreading.ThreadProc. But it would have led to RACE CONDITION.
    'This method instead is called from only the main thread so race condition is avoided
    
    modThreading.TotalThreads = TotalThreads - 1
    If modThreading.TotalThreads = 0 Then CleanUp
Else

'Subclass parentform.
    modThreading.TotalThreads = TotalThreads + 1
    
    If GetWindowLong(ParentFormHwnd, -4) = RetAddress(AddressOf NewWndProc) Then
        'Already subclassed
        Exit Sub
    End If
        modThreading.FormWndProc = SetWindowLong(ParentFormHwnd, -4, AddressOf NewWndProc)
End If
End Sub

Private Function RetAddress(ByVal Param As Long) As Long
'This is just a little tweak. VB doesn't allow something like -
' If GetWindowlong(..)= Addressof NewWndProc.
'So I had to use this method instead

    RetAddress = Param
End Function


Private Static Sub CleanUp()
Dim X As Object
'Unsubclass the forms
For Each X In Forms
    If GetWindowLong(X.hwnd, -4) = RetAddress(AddressOf NewWndProc) Then
        SetWindowLong X.hwnd, -4, FormWndProc
    End If
Next
End Sub



