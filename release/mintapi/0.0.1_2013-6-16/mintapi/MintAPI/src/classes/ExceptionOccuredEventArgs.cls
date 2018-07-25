VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExceptionOccuredEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Exception"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "ExceptionOccuredEventArgs"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IEventArgs
Implements IAliasable


Dim obj As Object
Dim exp As Exception

Dim inner_Data As EventArgs_InnerData

Friend Function GetInnerDataRefrence() As Long
    GetInnerDataRefrence = VarPtr(inner_Data)
End Function

Friend Sub Initialize(targetObject As Object, Optional ByVal Exception As Exception = Nothing)
    Set obj = targetObject
    Set exp = Exception
End Sub

Public Property Get Exception() As Exception
    Set Exception = exp
End Property

Public Property Get Sender() As Object: Set Sender = obj: End Property
Private Property Get IEventArgs_Sender() As Object: Set IEventArgs_Sender = obj: End Property

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
