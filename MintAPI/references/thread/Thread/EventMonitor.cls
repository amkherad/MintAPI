VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Multithreaded event monitor program
' Copyright © 1998 by Desaware Inc.
Option Explicit

' This object takes a Windows object that can be signalled and
' waits until it is signalled.
' It starts a background thread if necessary for the wait.
' Raises the WaitCompleted event when the event is signalled

Dim m_ObjectToWaitFor As Long
Dim m_AbortObject As Long

Event WaitCompleted(ByVal ObjectToWaitFor As Long, ByVal WaitResult As Long)

' Object is a system Waitable timer, event, semaphore, mutex, or other
' system object that can be signalled
Public Property Get ObjectToSignal() As Long
   ObjectToSignal = m_ObjectToWaitFor
End Property

' Set the object to signal, and start the wait
Public Property Let ObjectToSignal(ByVal vNewValue As Long)
   Dim res&
   If m_ObjectToWaitFor <> 0 Then
      Err.Raise vbObjectError + 512, "EventMonitor", "Object wait already in progress"
      Exit Property
   End If
   ' We do a wait to make sure the object isn't
   ' already signalled
   res = WaitForSingleObject(vNewValue, 0)
   If res = WAIT_FAILED Then
      ' Chances are the object's not valid
      Err.Raise vbObjectError + 513, "EventMonitor", "Wait operation failed - LastError:" & Err.LastDllError
      Exit Property
   End If
   If res <> WAIT_TIMEOUT Then
      ' It's timed out
      Signal res
      Exit Property
   End If
   m_ObjectToWaitFor = vNewValue
   ' Start the background thread
   res = StartBackgroundThreadApt(Me)
   If res = 0 Then
      ' Chances are the object's not valid
      Err.Raise vbObjectError + 514, "EventMonitor", "Unable to create background thread"
      m_ObjectToWaitFor = 0
      Exit Property
   End If
End Property

' Outside object shouldn't call this.
' It's public in order to force it to be marshalled
' Result of 1 indicates timer was aborted
Public Sub Signal(ByVal result As Long)
   RaiseEvent WaitCompleted(m_ObjectToWaitFor, result)
   m_ObjectToWaitFor = 0
End Sub

Public Sub Abort()
   Call PulseEvent(m_AbortObject)
End Sub

Private Sub Class_Initialize()
   m_AbortObject = CreateEvent(0, 0, 0, vbNullString)
End Sub

Private Sub Class_Terminate()
   Call CloseHandle(m_AbortObject)
End Sub

' Used by the module to get the event handle for aborting
Public Property Get AbortHandle() As Long
   AbortHandle = m_AbortObject
End Property
