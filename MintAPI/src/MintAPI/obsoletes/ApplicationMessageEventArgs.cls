VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApplicationMessageEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "ApplicationMessageEventArgs"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IEventArgs
Implements IAliasable


Dim Obj As Object
Dim objBList As New List
Dim retVList As New List

Friend Sub Initialize(TargetObject As Object, Args() As Variant)
    Set Obj = TargetObject
    
    
    
    Set retVList = objBList.Clone
End Sub

Public Property Get Sender() As Object: Set Sender = Obj: End Property
Private Property Get IEventArgs_Sender() As Object: Set IEventArgs_Sender = Obj: End Property


Public Property Get Count() As Long
    Count = objBList.Count
End Property

Public Property Get Item(Index As Long) As Variant
    If objBList.IsObject(Index) Then
        Set Item = objBList(Index)
    Else
            Item = objBList(Index)
    End If
End Property
Public Property Get Items() As List
    Set Items = retVList
End Property
'Public Property Let Item(Value As Variant)
'    objBList(Index) = Value
'End Property
'Public Property Set Item(Value As Object)
'    objBList(Index) = Value
'End Property

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
