VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "root"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "root"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"

Implements INamespace

Const CLASSES_INSTANCE As String = "instance"

Private Function INamespace_namespacename() As String
    INamespace_namespacename = CLASSID
End Function


Public Function instance_() As Instance
    
End Function
Public Function application_() As Instance
    
End Function
Public Function library_() As Instance
    
End Function

Public Function host_() As Instance 'host executing process.
    
End Function
