VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Action"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE


'Ideas from Me , Qt , .net Framework

Option Base 0
Option Explicit
Const CLASSID As String = "Action"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

Implements ICloneable
Implements IAliasable


Dim sc As String

Friend Sub Initialize(Script As String)
    
End Sub

Public Sub fromScript(Script As String)
    
End Sub
Public Function toScript() As String
    
End Function

Public Property Get Name() As String
    
End Property


Public Property Get IsAutomated() As Boolean
    
End Property

Public Function Execute() As Object
    
End Function


Public Function Clone() As Action
    
End Function

Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property

