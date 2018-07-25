VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Timer"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Const SIGNAL_TIMERELAPSED As String = "timerelapsed"
Const SIGNAL_TIMERSTART As String = "timerstart"
Const SIGNAL_TIMERSTOP As String = "timerstop"

Implements ISupportSignal
Implements ISupportSlot
Implements IAliasable

Private Declare Function API_SetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function API_KillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long


Public Event TimerStart(E As EventArgs)
Public Event TimerStop(E As EventArgs)
Public Event TimerElapsed(E As EventArgs)

Dim tmHndl As Long

Dim tmInterval As Long
Dim tmEnabled As Boolean
Dim tmInitDelay As Long

Dim tmhWnd As Long

Dim tmCallBackMethod As Method

Dim over_AllTicks As Long
Dim mySignalEmitter As New SignalEmitter
    
Private Sub Class_Initialize()
    Call mySignalEmitter.Initialize(Me)
End Sub

Friend Sub Initialize(Optional CallBack As Method, Optional Interval As Long = 1000, Optional InitialDelay As Long = 0, Optional StartNow As Boolean = False, Optional hWndParent As Long = 0)
    tmInterval = Interval
    tmInitDelay = InitialDelay
    tmEnabled = StartNow
    tmhWnd = hWndParent
    Set tmCallBackMethod = CallBack
    If StartNow Then _
        Call Start
End Sub
Friend Function GetHandle() As Long
    GetHandle = tmHndl
End Function
Friend Sub EnsureUnload()
    If tmHndl Then
        If API_KillTimer(0, tmHndl) <> 0 Then
            ' Timer killed.
            tmHndl = 0
        End If
    End If
End Sub

Public Sub ChangeCallBack(CallBackMethod As Method)
    Set tmCallBackMethod = CallBackMethod
End Sub

Private Sub InitializeTimer()
    On Error GoTo ErrorHandler
    ' Is the timer being started?
    If tmEnabled Then
        ' Is the timer currently running?
        If tmHndl = 0 Then
            ' Start the timer.
            If tmInterval > 0 Then
                ' Everything is okay.
                '------------
                ' Now create the timer.
                tmHndl = API_SetTimer(tmhWnd, 0, tmInterval, AddressOf modMain_Timer_CallBack_Procedure)
                If tmHndl <> 0 Then RaiseEvent TimerStart(EventArgs(Me))
            End If
        Else
            If API_KillTimer(0, tmHndl) <> 0 Then
                ' Timer killed.
                tmHndl = 0
            End If
            If tmInterval > 0 Then
                ' Restart the timer with the new interval.
                tmHndl = API_SetTimer(tmhWnd, 0, tmInterval, AddressOf modMain_Timer_CallBack_Procedure)
            End If
        End If
    ElseIf tmHndl <> 0 Then
        If API_KillTimer(0, tmHndl) <> 0 Then
            ' Timer killed.
            tmHndl = 0
            RaiseEvent TimerStop(EventArgs(Me))
        End If
    End If
    Exit Sub
ErrorHandler:
End Sub

Public Property Get Enabled() As Boolean
    Enabled = (tmHndl <> 0)
End Property
Public Property Let Enabled(Value As Boolean)
    tmEnabled = Value
    Call InitializeTimer
End Property

Public Property Get Interval() As Long
    Interval = tmInterval
End Property
Public Property Let Interval(Value As Long)
    tmInterval = Value
    Call InitializeTimer
End Property

Public Sub Dispose()
    If tmHndl <> 0 Then
        If API_KillTimer(0, tmHndl) <> 0 Then
            ' Timer killed.
            tmHndl = 0
            RaiseEvent TimerStop(EventArgs(Me))
        End If
    End If
    tmEnabled = False
    tmInterval = 1000
    Set tmCallBackMethod = Nothing
End Sub
Friend Sub HandleCallBack(ByVal lngTimerID As Long)
On Error GoTo ErrorHandler
    ' Make sure that the message is WM_TIMER.
    If tmHndl = lngTimerID Then
        over_AllTicks = over_AllTicks + 1
        Call tmCallBackMethod.Invoke
        RaiseEvent TimerElapsed(EventArgs(Me))
    End If
    Exit Sub
ErrorHandler:
End Sub

Public Sub Start(Optional Interval As Long = -1, Optional CallBack As Method)
    
End Sub
Public Sub Pause()
    
End Sub
Public Sub Refresh()
    
End Sub

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter
End Function

Private Property Get ISupportSlot_AutoManage() As Boolean
    
End Property
Private Sub ISupportSlot_ManualManageCall(SlotName As String)
    
End Sub
Private Function ISupportSlot_Methods() As String
    
End Function
Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
