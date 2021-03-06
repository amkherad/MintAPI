VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DisposingEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "DisposingEventArgs class generally occures when some class or sth is disposing.\r\nThis supports cancel action."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "DisposingEventArgs"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IEventArgs
Implements IAliasable


Dim obj As Object
Dim canceled As Boolean

Friend Sub Initialize(targetObject As Object, Optional Continue As Boolean = True)
    Set obj = targetObject
    canceled = Not Continue
End Sub

Public Property Get Cancel() As Boolean
    Cancel = canceled
End Property
Public Property Let Cancel(Value As Boolean)
    canceled = Value
End Property

Public Sub CancelAction()
    canceled = True
End Sub

Public Property Get Sender() As Object: Set Sender = obj: End Property
Private Property Get IEventArgs_Sender() As Object: Set IEventArgs_Sender = obj: End Property

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
