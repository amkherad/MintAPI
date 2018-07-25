VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Comparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="Comparer" tags="">
''  <summary>
''    Generic comparer to compare two variable with binary compare method.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IComparer

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Comparer" '#DO NOT REMOVE IT
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
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

Private Function IComparer_Compare(Param1 As Variant, Param2 As Variant) As CompareResults
    Dim VT1 As VariantTypes, VT2 As VariantTypes
    VT1 = VarType(Param1)
    VT2 = VarType(Param2)
    If VT1 <> VT2 Then
        IComparer_Compare = crNotEqual
        Exit Function
    End If
    
    Select Case VT1
        Case VT_ARRAY
            IComparer_Compare = ArrayCompare(Param1, Param2)
        Case VT_Record, VT_UDT
            IComparer_Compare = RecordCompare(Param1, Param2)
        Case VT_DISPATCH
            If TypeOf Param1 Is IComparable Then
                Dim c As IComparable
                Set c = Param1
                IComparer_Compare = c.CompareTo(Param2)
            ElseIf TypeOf Param1 Is IObject Then
                Dim O As IObject
                Set O = Param1
                IComparer_Compare = IIf(O.Equals(Param2), crEqual, crNotEqual)
            Else
                IComparer_Compare = IIf(Param1 Is Param2, crEqual, crNotEqual)
            End If
        Case Else
            Dim Result As Long
            If (Param1 = Param2) Then
                IComparer_Compare = crEqual
            ElseIf (Param1 > Param2) Then
                IComparer_Compare = crGreater
            ElseIf (Param1 < Param2) Then
                IComparer_Compare = crLess
            Else
                IComparer_Compare = crNotEqual
            End If
    End Select
End Function
Private Function IComparer_Equals(Param1 As Variant, Param2 As Variant) As Boolean
    IComparer_Equals = (IComparer_Compare(Param1, Param2) = crEqual)
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
