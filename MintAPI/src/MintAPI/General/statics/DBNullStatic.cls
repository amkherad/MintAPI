VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBNullStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="DBNullStatic" static tags="">
''  <summary>
''    Provides some static methods for DBNull.
''  </summary>
''  <using>DBNull</using>
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
Const CLASSID As String = "DBNullStatic" '#DO NOT REMOVE IT
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

''<summary>
''  Represents the sole instance of the DBNull class.
''</summary>
Public Property Get Value() As DBNull
Attribute Value.VB_UserMemId = 0
    Set Value = DefaultDBNull
End Property

''<summary>
''  Represents the missing value in optional parameters.
''</summary>
'Public Property Get MissingValue() As Variant
'    Call memzero(MissingValue, VLEN_VARIANT)
'End Property

''<summary>
''  Determines that the given value is a DBNull or not.
''</summary>
''<params><param Name="Value">The value to be compared with DBNull.</param></params>
Public Function IsDBNull(Value) As Boolean
    If IsNull(Value) Then IsDBNull = True: Exit Function
    If IsEmpty(Value) Then IsDBNull = True: Exit Function
    
    If Not IsObject(Value) Then Exit Function
    Dim vObj As Object
    Set vObj = Value
    If TypeOf vObj Is MintAPI.DBNull Then _
        IsDBNull = True
End Function

''<summary>
''  Determines that the given value is a DBNull or null.
''</summary>
''<params><param Name="Value">The value to be compared with DBNull and null or even empty.</param></params>
Public Function IsDBNullOrNull(Value) As Boolean
    If IsNull(Value) Then IsDBNullOrNull = True: Exit Function
    If IsEmpty(Value) Then IsDBNullOrNull = True: Exit Function
    
    If Not IsObject(Value) Then Exit Function
    Dim vObj As Object
    Set vObj = Value
    If TypeOf vObj Is MintAPI.DBNull Then
        IsDBNullOrNull = True
    ElseIf vObj Is Nothing Then
        IsDBNullOrNull = True
    End If
End Function

''<summary>
''  Determines that the given value is a DBNull or null or a missing variant.
''</summary>
''<params><param Name="Value">The value to be compared with DBNull and null or even empty.</param></params>
'Public Function IsDBNullOrNullOrMissing(Value) As Boolean
'    If IsMissing(Value) Then IsDBNullOrNullOrMissing = True: Exit Function
'    If IsNull(Value) Then IsDBNullOrNullOrMissing = True: Exit Function
'    If IsEmpty(Value) Then IsDBNullOrNullOrMissing = True: Exit Function
'
'    If Not IsObject(Value) Then Exit Function
'    Dim vObj As Object
'    Set vObj = Value
'    If TypeOf vObj Is MintAPI.DBNull Then
'        IsDBNullOrNullOrMissing = True
'    ElseIf vObj Is Nothing Then
'        IsDBNullOrNullOrMissing = True
'    End If
'End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
