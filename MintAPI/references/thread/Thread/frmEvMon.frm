VERSION 5.00
Begin VB.Form frmEventMonTest 
   Caption         =   "Event Monitor Test"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   375
      Left            =   450
      TabIndex        =   4
      Top             =   1350
      Width           =   1185
   End
   Begin VB.TextBox txtTimer2 
      Height          =   375
      Left            =   2070
      TabIndex        =   3
      Text            =   "1000"
      Top             =   810
      Width           =   1545
   End
   Begin VB.CommandButton cmdTimer2 
      Caption         =   "Start Timer2"
      Height          =   375
      Left            =   450
      TabIndex        =   2
      Top             =   810
      Width           =   1185
   End
   Begin VB.CommandButton cmdTimer1 
      Caption         =   "Start Timer1"
      Height          =   375
      Left            =   450
      TabIndex        =   1
      Top             =   270
      Width           =   1185
   End
   Begin VB.TextBox txtTimer1 
      Height          =   375
      Left            =   2070
      TabIndex        =   0
      Text            =   "1000"
      Top             =   270
      Width           =   1545
   End
End
Attribute VB_Name = "frmEventMonTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Event monitor test
' Copyright © 1998 by Desaware Inc. All Rights Reserved
' Note - Waitable timers are NT4 and later only

Option Explicit

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Declare Function CreateWaitableTimer Lib "kernel32" Alias "CreateWaitableTimerA" (ByVal lpTimerAttributes As Long, ByVal bManualReset As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetWaitableTimer Lib "kernel32" (ByVal hTimer As Long, lpDueTime As FILETIME, ByVal lPeriod As Long, ByVal pfnCompletionRoutine As Long, ByVal lpArgToCompletionRoutine As Long, ByVal fResume As Long) As Long
Private Declare Sub agConvertDoubleToFileTime Lib "apigid32.dll" (ByVal d As Double, f1 As Any)
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long


Dim WithEvents waitobject1 As EventMonitor
Attribute waitobject1.VB_VarHelpID = -1
Dim WithEvents waitobject2 As EventMonitor
Attribute waitobject2.VB_VarHelpID = -1
Dim Timer1 As Long
Dim Timer2 As Long


Private Sub cmdAbort_Click()
   waitobject1.Abort
   waitobject2.Abort
End Sub

Private Sub cmdTimer1_Click()
   Dim totaltime As Double
   Dim ft As FILETIME
   totaltime = Val(txtTimer1.Text)
   Call agConvertDoubleToFileTime(-10000 * totaltime, ft)
   Call SetWaitableTimer(Timer1, ft, 0, 0, 0, 0)
   On Error GoTo ReportError1
   waitobject1.ObjectToSignal = Timer1
   On Error GoTo 0
   Exit Sub
ReportError1:
   MsgBox "Error: " & Err.Description
End Sub

Private Sub cmdTimer2_Click()
   Dim totaltime As Double
   Dim ft As FILETIME
   totaltime = Val(txtTimer2.Text)
   Call agConvertDoubleToFileTime(-10000 * totaltime, ft)
   
   Call SetWaitableTimer(Timer2, ft, 0, 0, 0, 0)
   On Error GoTo ReportError2
   waitobject2.ObjectToSignal = Timer2
   On Error GoTo 0
   Exit Sub
ReportError2:
   MsgBox "Error: " & Err.Description
End Sub


' Create our two timers
Private Sub Form_Load()
   Set waitobject1 = New EventMonitor
   Set waitobject2 = New EventMonitor
   Timer1 = CreateWaitableTimer(0, 1, vbNullString)
   Timer2 = CreateWaitableTimer(0, 1, vbNullString)
End Sub

' Clean up the handles
Private Sub Form_Unload(Cancel As Integer)
   Call CloseHandle(Timer1)
   Call CloseHandle(Timer2)
End Sub

' Why use the API message box? Because the VB message box
' blocks events.
Private Sub waitobject1_WaitCompleted(ByVal ObjectToWaitFor As Long, ByVal WaitResult As Long)
   Call MessageBox(hwnd, "Timer1 expired with flags: " & WaitResult, "", 0)
End Sub

Private Sub waitobject2_WaitCompleted(ByVal ObjectToWaitFor As Long, ByVal WaitResult As Long)
   Call MessageBox(hwnd, "Timer2 expired with flags: " & WaitResult, "", 0)
End Sub
