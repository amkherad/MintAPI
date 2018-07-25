VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MethodPrototypeStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MethodPrototypeStatic" static tags="">
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
Const CLASSID As String = "MethodPrototypeStatic" '#DO NOT REMOVE IT
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
Public Function FromMethod(ByVal Method As Method) As MethodPrototype
    If Method Is Nothing Then throw Exps.ArgumentNullException("Method")
    Set FromMethod = Method.Prototype.Clone
End Function
'public function FromMember(
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

Public Property Get Factory() As MethodPrototypeFactory
    Set Factory = New MethodPrototypeFactory
End Property

Public Property Get VoidMethod() As MethodPrototype
    Set VoidMethod = New MethodPrototype
End Property

Public Function IsUnspecified(ByVal Prototype As MethodPrototype) As Boolean
    If Prototype Is Nothing Then throw Exps.ArgumentNullException("Prototype")
    IsUnspecified = Prototype.IsUnspecified
End Function
Public Property Get Unspecified() As MethodPrototype
    Set Unspecified = New MethodPrototype
    Call Unspecified.Constructor1(True, False)
End Property
Friend Property Get UnspecifiedUnsafe() As MethodPrototype
    Set Unspecified = New MethodPrototype
    Call Unspecified.Constructor1(True, True)
End Property

Public Property Get AsEAXValue() As Long
    
End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
