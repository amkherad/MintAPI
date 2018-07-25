VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="Enumerable" static extensions tags="">
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
Const CLASSID As String = "Enumerable" '#DO NOT REMOVE IT
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

Public Function Count(ByVal This_Enumerable As IEnumerable) As Long
    
End Function
Public Function LongCount(ByVal This_Enumerable As IEnumerable) As Int64
    
End Function

Public Function IsAny(ByVal This_Enumerable As IEnumerable, Optional ByVal Predicate As Method) As Boolean
'Bool Predicate(Variant)
    If This_Enumerable Is Nothing Then throw Exps.ArgumentNullException
    If Predicate Is Nothing Then throw Exps.ArgumentNullException
    Dim V As Variant
    If Predicate Is Nothing Then
        For Each V In This_Enumerable
            If Not IsNullOrEmpty(V) Then
                IsAny = True
                Exit Function
            End If
        Next
    Else
        For Each V In This_Enumerable
            If CBool(Predicate(AL.Create(V))) Then
                IsAny = True
                Exit Function
            End If
        Next
    End If
    IsAny = False
End Function
Public Function IsAll(ByVal This_Enumerable As IEnumerable, ByVal Predicate As Method) As Boolean
'Bool Predicate(Variant)
    If This_Enumerable Is Nothing Then throw Exps.ArgumentNullException
    If Predicate Is Nothing Then throw Exps.ArgumentNullException
    Dim V As Variant
    IsAll = False
    For Each V In This_Enumerable
        If Not CBool(Predicate(AL.Create(V))) Then
            Exit Function
        Else
            IsAll = True
        End If
    Next
End Function
Public Function Where(ByVal This_Enumerable As IEnumerable, ByVal Predicate As Method) As EnumerableDelegate
'Bool Predicate(Variant)
    If This_Enumerable Is Nothing Then throw Exps.ArgumentNullException
    If Predicate Is Nothing Then throw Exps.ArgumentNullException
    Dim RetList As New List
    Dim V As Variant
    For Each V In This_Enumerable
        If Not CBool(Predicate(AL.Create(V))) Then
            Call RetList.Append(V)
        End If
    Next
    Set Where = New EnumerableDelegate
    Call Where.Constructor0(RetList)
End Function
Public Function SelectAs(ByVal This_Enumerable As IEnumerable, ByVal Predicate As Method) As EnumerableDelegate
'Variant Predicate(Variant)
    If This_Enumerable Is Nothing Then throw Exps.ArgumentNullException
    If Predicate Is Nothing Then throw Exps.ArgumentNullException
    Dim RetList As New List
    Dim V As Variant
    For Each V In This_Enumerable
        Call RetList.Append(Predicate(AL.Create(V)))
    Next
    Set SelectAs = New EnumerableDelegate
    Call SelectAs.Constructor0(RetList)
End Function
Public Function Skip(ByVal This_Enumerable As IEnumerable, ByVal Predicate As Method) As EnumerableDelegate
    
End Function
Public Function Take(ByVal This_Enumerable As IEnumerable, ByVal Predicate As Method) As EnumerableDelegate
    
End Function
Public Sub ForEach(ByVal This_Enumerable As IEnumerable, ByVal Action As Method)
'void Action(Variant)
    If This_Enumerable Is Nothing Then throw Exps.ArgumentNullException
    If Action Is Nothing Then throw Exps.ArgumentNullException
    
End Sub
Public Function Distinct(ByVal This_Enumerable As IEnumerable, Optional ByVal Predicate As Method) As EnumerableDelegate
    
End Function
'Public Function Sort(ByVal This_Enumerable As IEnumerable, ByVal Predicate As Method) As IEnumerable
'
'End Function
'Public Function Cast(ByVal This_Enumerable As IEnumerable, ByVal Predicate As Method) As IEnumerable
'
'End Function
Public Function OrderBy(ByVal This_Enumerable As IEnumerable, Optional ByVal Predicate As Method) As EnumerableDelegate
    
End Function
Public Function OrderByDescending(ByVal This_Enumerable As IEnumerable, Optional ByVal Predicate As Method) As EnumerableDelegate
    
End Function

Public Function ToList(ByVal This_Enumerable As IEnumerable) As List
    If This_Enumerable Is Nothing Then throw Exps.ArgumentNullException
    Set ToList = New List
    Dim V As Variant
    For Each V In This_Enumerable
        Call ToList.Append(V)
    Next
End Function
Public Function ToVector(ByVal This_Enumerable As IEnumerable) As Vector
    If This_Enumerable Is Nothing Then throw Exps.ArgumentNullException
    Set ToVector = ToList(This_Enumerable).ToVector
End Function
Public Function ToArray(ByVal This_Enumerable As IEnumerable) As Variant
    If This_Enumerable Is Nothing Then throw Exps.ArgumentNullException
    Call EvaluateX(ToArray, ToList(This_Enumerable).ToArray)
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
