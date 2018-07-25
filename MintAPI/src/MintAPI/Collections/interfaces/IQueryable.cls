VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IQueryable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="IQueryable" interface tags="">
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
Const CLASSID As String = "IQueryable" '#DO NOT REMOVE IT
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

Public Function GetEnumerator() As IEnumerator: End Function
Public Function NewEnum() As stdole.IUnknown: End Function
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

Public Function AsEnumerable() As IEnumerable: End Function

'Public Sub ForEach(ByVal Predicate As Method): End Sub
'Public Function Where(ByVal Predicate As Method) As IEnumerable: End Function
'Public Function Exists(ByVal Predicate As Method) As Boolean: End Function
'Public Function Count(ByVal Predicate As Method) As Int64: End Function
'Public Function Sort(ByVal Predicate As Method) As IEnumerable: End Function
'Public Function Search(ByVal Predicate As Method) As IEnumerable: End Function
'Public Function Find(ByVal Predicate As Method) As Variant: End Function
'Public Function FindAll(ByVal Predicate As Method) As IEnumerable: End Function
'Public Function TrueForAll(ByVal Predicate As Method) As IEnumerable: End Function
'Public Function Distinct(ByVal Predicate As Method) As IEnumerable: End Function
'Public Function Group(ByVal Predicate As Method) As IEnumerable: End Function
'Public Function Cast(ByVal Predicate As Method) As IEnumerable: End Function
'Public Function OfType(ByVal Predicate As Method) As IEnumerable: End Function
'Public Function Filter(ByVal Predicate As Method) As IEnumerable: End Function

Public Function ToList() As List: End Function
Public Function ToVector() As Vector: End Function
Public Function ToArray() As Variant: End Function
'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

