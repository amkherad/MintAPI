Attribute VB_Name = "modMTBack"
' modMTBack
' MTDemo3 multithreading example
' Copyright (c) 1997 by Desaware Inc.
' All Rights Reserved

Option Explicit

Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

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

Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Sub Main()
   frmMTDemo3.Show
End Sub

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

' A free threaded callback.
' This is an invalid approach, though it works
' in this case.
Public Function BackgroundFuncFree(ByVal param As IUnknown) As Long
   Dim qobj As clsBackground
   Dim res&
   
   ' Free threaded approach
   Set qobj = param
   
   Do While Not qobj.DoTheCount(100000)
   Loop
   
   ' qobj.ShowAForm ' Crashes!
   
   ' Thread ends on return
End Function

' Start the background thread for this object
' using the invalid free threading approach.
Public Function StartBackgroundThreadFree(ByVal qobj As clsBackground)
   Dim threadid As Long
   Dim hnd&
   Dim threadparam As Long
   ' Free threaded approach
   threadparam = ObjPtr(qobj)
   hnd = CreateThread(0, 2000, AddressOf BackgroundFuncFree, threadparam, 0, threadid)
   If hnd = 0 Then
      ' Return with zero (error)
      Exit Function
   End If
   ' We don't need the thread handle
   CloseHandle hnd
   StartBackgroundThreadFree = threadid
End Function

' An correctly marshalled apartment model callback.
' This is the correct approch, though slower.
Public Function BackgroundFuncApt(ByVal param As Long) As Long
   Dim qobj As Object
   Dim qobj2 As clsBackground
   Dim res&
   ' This new thread is a new apartment, we must
   ' initialize OLE for this apartment (VB doesn't seem to do it)
   res = CoInitialize(0)
   ' Proper apartment modeled approach
   res = CoGetInterfaceAndReleaseStream(param, IID_IDispatch, qobj)
   Set qobj2 = qobj
   Do While Not qobj2.DoTheCount(10000)
   Loop
   qobj2.ShowAForm
   
   ' Alternatively, you can put a wait function here,
   ' then call the qobj function when the wait is
   ' satisfied
   
   ' All calls to CoInitialize must be balanced
   CoUninitialize
End Function

' Start the background thread for this object
' using the apartment model
' Returns zero on error
Public Function StartBackgroundThreadApt(ByVal qobj As clsBackground)
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
   hnd = CreateThread(0, 2000, AddressOf BackgroundFuncApt, threadparam, 0, threadid)
   If hnd = 0 Then
      ' Return with zero (error)
      Exit Function
   End If
   ' We don't need the thread handle
   CloseHandle hnd
   StartBackgroundThreadApt = threadid
   ' This message box can cause problems
   ' MsgBox "New thread created " & threadid
End Function



