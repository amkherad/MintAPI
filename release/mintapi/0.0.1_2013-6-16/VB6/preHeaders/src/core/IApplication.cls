VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Main(args) 'occures when starting application.
    throw NotImplementedException
End Sub
Public Sub Starting(ByVal e As EventArgs)
    throw NotImplementedException
End Sub
Public Sub Started(ByVal e As EventArgs)
    throw NotImplementedException
End Sub
Public Sub Terminating(ByVal e As DisposingEventArgs)
    throw NotImplementedException
End Sub
Public Sub Terminated(ByVal e As EventArgs)
    throw NotImplementedException
End Sub

Public Sub ApplicationInputMethodChanged(ByVal e As EventArgs)
    throw NotImplementedException
End Sub
Public Function ExecutionPath() As String 'including exefile name.
    throw NotImplementedException
End Function
Public Function ApplicationPath() As String 'excluding exefile name.
    throw NotImplementedException
End Function

Public Function VersionString() As String 'ex : 0.0.0.1054 ultimate
    throw NotImplementedException
End Function
Public Function VersionTags() As String 'ex : ultimate plugin01 monitorplugin
    throw NotImplementedException
End Function


Public Sub Error(ByVal e As IException)
    throw NotImplementedException
End Sub
