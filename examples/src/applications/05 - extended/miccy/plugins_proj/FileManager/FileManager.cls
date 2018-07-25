VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const CLASSID As String = "FileManager"
Implements Iwindow

Private Function Iwindow_provideCaption() As String
    Iwindow_provideCaption = "File Manager"
End Function
Private Function Iwindow_provideClosable() As Boolean
    Iwindow_provideClosable = True
End Function
Private Function Iwindow_provideWindowHandle() As Long
    Iwindow_provideWindowHandle = frmbuff.FileManager.hWnd
End Function

Private Sub Iwindow_ShowWindow()
    'show event.
End Sub
Private Sub Iwindow_HideWindow(Cancel As Boolean)
    'hide event.
End Sub
