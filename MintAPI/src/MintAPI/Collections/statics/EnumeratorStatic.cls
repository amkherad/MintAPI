VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumeratorStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="EnumeratorStatic" static tags="">
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
Const CLASSID As String = "EnumeratorStatic" '#DO NOT REMOVE IT
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
'<section Creators>
'
Public Function CreateEnumerator(ByVal ObjEnumerator As IEnumerator) As IUnknown
Attribute CreateEnumerator.VB_UserMemId = 0
    Set CreateEnumerator = mint_enumerator.CreateEnumerator(ObjEnumerator)
End Function
Public Function LazyEnumerator(ByVal LazyCollection As LazyList) As LazyEnumerator
    
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

Public Function GetEnumerator(ByRef Enumerable As Variant) As IEnumerator
    If IsObject(Enumerable) Then
        Dim Obj As Object, Enmrbl As IEnumerable
        If Not TypeOf Obj Is IEnumerable Then throw Exps.InvalidArgumentException("Invalid Enumerable type.")
        Set Enmrbl = Obj
        Set GetEnumerator = Enmrbl.GetEnumerator
    ElseIf IsArray(Enumerable) Then
        Dim AEnumerator As New ArrayEnumerator
        Call AEnumerator.Constructor0(Enumerable)
        Set GetEnumerator = AEnumerator
    Else
        throw Exps.ArgumentException("Invalid Enumerable type.")
    End If
End Function
Public Function GetNewEnum(ByRef Enumerable As Variant) As IUnknown
    Set GetNewEnum = CreateEnumerator(GetEnumerator(Enumerable))
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
