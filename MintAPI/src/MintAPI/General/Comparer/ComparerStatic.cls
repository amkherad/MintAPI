VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComparerStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ComparerStatic" static tags="">
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
Const CLASSID As String = "ComparerStatic" '#DO NOT REMOVE IT
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
'<section Creators>
'
Public Property Get Default() As IComparer
Attribute Default.VB_UserMemId = 0
    Set Default = DefaultComparer
End Property
'Public Property Get BinaryComparer() As IComparer
'
'End Property

Public Function FromDelegator(ByVal ComparerDelegator As Method) As IComparer
    Dim CompDeleg As ComparerDelegator
    Set CompDeleg = New ComparerDelegator
    Call CompDeleg.Constructor0(ComparerDelegator)
    Set FromDelegator = CompDeleg
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

''<summary>Returns the maximum value in the array passed.</summary>
''<params>
''  <param name="vArray">An array to get the max value from it.</param>
''  <param name="Comparer">Comparer to compare values. Null to use default comparer.</param>
''</params>
Public Function MaxValue(ByRef vArray As Variant, Optional ByVal Comparer As IComparer) As Variant
    Dim Max As Variant
    Dim i As Long, VT As VariantTypeConstants
    If Not IsArray(vArray) Then throw Exps.Exception("MaxValue Requires An Array As Parameter.")
    If Not ArraySize(vArray) > 0 Then throw Exps.ArgumentNullException("Empty Array.")
    If Comparer Is Nothing Then Set Comparer = Default
    For i = 0 To UBound(vArray)
        If i = 0 Then
            Max = vArray(0)
            VT = VarType(vArray(0))
        Else
            If VarType(vArray(i)) <> VT Then throw Exps.ArgumentTypeMismatch
            If Comparer.Compare(vArray(i), Max) = crGreater Then Max = vArray(i)
        End If
    Next i
    MaxValue = Max
End Function
''<summary>Returns the minimum value in the array passed.</summary>
''<params>
''  <param name="vArray">An array to get the max value from it.</param>
''  <param name="Comparer">Comparer to compare values. Null to use default comparer.</param>
''</params>
Public Function MinValue(ByRef vArray As Variant, Optional ByVal Comparer As IComparer) As Variant
    Dim Min As Variant
    Dim i As Long, VT As VariantTypeConstants
    If Not IsArray(vArray) Then throw Exps.Exception("MinValue Requires An Array As Parameter.")
    If Not ArraySize(vArray) > 0 Then throw Exps.ArgumentNullException("Empty Array.")
    If Comparer Is Nothing Then Set Comparer = Default
    For i = 0 To UBound(vArray)
        If i = 0 Then
            Min = vArray(0)
            VT = VarType(vArray(0))
        Else
            If Not VarType(vArray(i)) = VT Then throw Exps.ArgumentTypeMismatch
            If Comparer.Compare(vArray(i), Min) = crLess Then Min = vArray(i)
        End If
    Next i
    MinValue = Min
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
