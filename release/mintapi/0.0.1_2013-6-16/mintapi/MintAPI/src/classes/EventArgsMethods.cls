VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventArgsMethods"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function EventArgs(targetObject As Object) As EventArgs
    Set EventArgs = modMain.EventArgs(targetObject)
End Function
Public Function DisposingEventArgs(targetObject As Object, Optional Continue As Boolean = True) As DisposingEventArgs
    Set DisposingEventArgs = modMain.DisposingEventArgs(targetObject, Continue)
End Function
