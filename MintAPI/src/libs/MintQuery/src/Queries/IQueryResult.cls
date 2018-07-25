VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IQueryResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="IQueryResult" interface tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IResult

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "IQueryResult" '#DO NOT REMOVE IT
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

Public Sub Execute(): End Sub

Public Property Get Result() As Variant: End Property
Public Property Get ResultType() As VariantTypes: End Property

Public Function ToString() As String: End Function
Public Function Equals(ByRef Comparable As Variant) As Boolean: End Function
Public Function Compare(ByRef Comparable As Variant) As CompareResults: End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

Private Function IResult_Compare(Comparable As Variant) As MintAPI.CompareResults: IResult_Compare = Compare(Comparable): End Function
Private Function IResult_Equals(Comparable As Variant) As Boolean: IResult_Equals = Equals(Comparable): End Function
Private Sub IResult_Execute(): Call Execute: End Sub
Private Property Get IResult_Result() As Variant: Call EvaluateX(IResult_Result, Result): End Property
Private Property Get IResult_ResultType() As MintAPI.VariantTypes: IResult_ResultType = ResultType: End Property
Private Function IResult_ToString() As String: IResult_ToString = ToString: End Function
