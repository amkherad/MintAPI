VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Iwindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CLASSID As String = "Iwindow"

Public Sub ShowWindow(): throw Exps.NotImplementedException: End Sub
Public Sub HideWindow(Cancel As Boolean): throw Exps.NotImplementedException: End Sub

Public Function provideCaption() As String: throw Exps.NotImplementedException: End Function
Public Function provideWindowHandle() As Long: throw Exps.NotImplementedException: End Function
Public Function provideClosable() As Boolean: throw Exps.NotImplementedException: End Function
