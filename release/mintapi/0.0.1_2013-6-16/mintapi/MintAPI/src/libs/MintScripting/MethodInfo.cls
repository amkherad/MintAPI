VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MethodInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "MethodInfo"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"

'Implements IClass
'
'Private Function IClass_class_name() As String
'
'End Function

Public Property Get Name() As String
    
End Property
Public Property Get Arguments() As List
    
End Property

Public Property Get ReturnType() As VbVarType
    
End Property
