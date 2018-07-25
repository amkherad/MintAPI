VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "ControlEventHandler"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Dim obj As Object

Dim WithEvents f As VB.Form
Attribute f.VB_VarHelpID = -1
Dim WithEvents t As VB.Timer
Attribute t.VB_VarHelpID = -1

Friend Sub Initialize(targetObject As Object)
    Set obj = targetObject
End Sub

Private Sub f_Load()
    '
End Sub
Private Sub f_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
End Sub
Private Sub f_Terminate()
    '
End Sub
Private Sub t_Timer()
    '
End Sub


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
