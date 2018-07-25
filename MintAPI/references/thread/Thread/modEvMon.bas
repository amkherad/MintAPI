Attribute VB_Name = "modMTBack"
' modMTBack
' MTDemo3 multithreading example
' Copyright (c) 1997 by Desaware Inc.
' All Rights Reserved

Option Explicit

Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function PulseEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long

Public Const WAIT_FAILED = -1&
Public Const WAIT_OBJECT_0 = 0
Public Const WAIT_ABANDONED = &H80&
Public Const WAIT_ABANDONED_0 = &H80&
Public Const WAIT_TIMEOUT = &H102&
Public Const WAIT_IO_COMPLETION = &HC0&
Public Const STILL_ACTIVE = &H103&
Public Const INFINITE = -1&


' Structure to hold IDispatch GUID
Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Public IID_IDispatch As GUID

Declare Function CoMarshalInterThreadInterfaceInStream Lib "ole32.dll" _
   (riid As GUID, ByVal pUnk As IUnknown, ppStm As Long) As Long
 
Declare Function CoGetInterfaceAndReleaseStream Lib "ole32.dll" _
   (ByVal pStm As Long, riid As GUID, pUnk As IUnknown) As Long
 
Declare Function CoInitialize Lib "ole32.dll" (ByVal pvReserved As Long) As Long
Declare Sub CoUninitialize Lib "ole32.dll" ()
 
Declare Function CreateThread Lib "kernel32" (ByVal lpSecurityAttributes As Long, _
   ByVal dwStackSize As Long, _
   ByVal lpStartAddress As Long, _
   ByVal lpParameter As Long, _
   ByVal dwCreationFlags As Long, _
   lpThreadId As Long) As Long
   
Declare Function GetCurrentThreadId Lib "kernel32" () As Long


' Initialize the GUID structure
Private Sub InitializeIID()
   Static Initialized As Boolean
   If Initialized Then Exit Sub
   With IID_IDispatch
      .Data1 = &H20400
      .Data2 = 0
      .Data3 = 0
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
   Initialized = True
End Sub


' An correctly marshalled apartment model callback.
' This is the correct approch, though slower.
Public Function BackgroundFuncApt(ByVal param As Long) As Long
   Dim ObjList(1) As Long
   Dim qobj As Object
   Dim qobj2 As EventMonitor
   Dim res&
   ' This new thread is a new apartment, we must
   ' initialize OLE for this apartment (VB doesn't seem to do it)
   res = CoInitialize(0)
   ' Proper apartment modeled approach
   res = CoGetInterfaceAndReleaseStream(param, IID_IDispatch, qobj)
   Set qobj2 = qobj
   
   ObjList(0) = qobj2.ObjectToSignal
   ObjList(1) = qobj2.AbortHandle
   
   ' Suspend the thread until one of the objects is signaled.
   ' This will either be the one we are waiting for, or the Abort
   ' event
   res = WaitForMultipleObjects(2, ObjList(0), 0, INFINITE)
   
   qobj2.Signal res
   
   ' All calls to CoInitialize must be balanced
   CoUninitialize
   ' The thread will terminate after this
End Function

' Start the background thread for this object
' using the apartment model
' Returns zero on error
Public Function StartBackgroundThreadApt(ByVal qobj As EventMonitor)
   Dim threadid As Long
   Dim hnd&, res&
   Dim threadparam As Long
   Dim tobj As Object
   Set tobj = qobj
   ' Proper marshalled approach
   InitializeIID
   res = CoMarshalInterThreadInterfaceInStream(IID_IDispatch, qobj, threadparam)
   If res <> 0 Then
      StartBackgroundThreadApt = 0
      Exit Function
   End If
   hnd = CreateThread(0, 8000, AddressOf BackgroundFuncApt, threadparam, 0, threadid)
   If hnd = 0 Then
      ' Return with zero (error)
      Exit Function
   End If
   ' We don't need the thread handle
   CloseHandle hnd
   StartBackgroundThreadApt = threadid
End Function



