VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAutoCodeGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="IAutoCodeGenerator" interface tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "IAutoCodeGenerator" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Interface Body Members>
'
'*********************************************
'*********************************************

Public Sub INop(): End Sub
Public Sub ISet(ByVal Var1 As IVariableDelegator, ByVal Var2 As IVariableDelegator): End Sub
Public Sub IRet(Optional ByVal Length As Long = 0): End Sub
Public Sub IAdd(ByVal Var1 As IVariableDelegator, ByVal Var2 As IVariableDelegator): End Sub
Public Sub ISub(ByVal Var1 As IVariableDelegator, ByVal Var2 As IVariableDelegator): End Sub
Public Sub IMul(ByVal Var1 As IVariableDelegator, ByVal Var2 As IVariableDelegator): End Sub
Public Sub IDiv(ByVal Var1 As IVariableDelegator, ByVal Var2 As IVariableDelegator): End Sub
Public Sub IPow(ByVal Var1 As IVariableDelegator, ByVal Var2 As IVariableDelegator): End Sub
Public Sub IFact(ByVal Var1 As IVariableDelegator): End Sub
Public Sub IPush(ByVal Var1 As IVariableDelegator): End Sub
Public Sub IPop(ByVal Var1 As IVariableDelegator): End Sub
Public Sub ISwap(ByVal Var1 As IVariableDelegator, ByVal Var2 As IVariableDelegator): End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

