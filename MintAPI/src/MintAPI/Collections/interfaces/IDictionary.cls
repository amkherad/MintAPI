VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="" interface tags="">
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
Const CLASSID As String = "MintAPIInterface" '#DO NOT REMOVE IT
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
Public Function NewEnum() As IUnknown: End Function
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Public Property Get Count() As Long: End Property
Public Sub CopyTo(ByRef Arr As Variant, ByVal Index As Long): End Sub
Public Property Get IsFixedSize() As Boolean: End Property
Public Property Get IsReadOnly() As Boolean: End Property
Public Property Get Item(ByRef Key As Variant) As Variant: End Property
Public Property Let Item(ByRef Key As Variant, ByRef Value As Variant): End Property
Public Property Set Item(ByRef Key As Variant, ByRef Value As Variant): End Property
Public Property Get Keys() As ICollection: End Property
Public Property Get Values() As ICollection: End Property
Public Sub Add(ByRef Key As Variant, ByRef Value As Variant): End Sub
Public Sub Clear(): End Sub
Public Function Contains(ByRef Key As Variant) As Boolean: End Function
Public Sub Remove(ByRef Key As Variant): End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

