VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Member1" ,"Version"
Attribute VB_Ext_KEY = "Member2" ,"Guid"
'@PROJECT_LICENSE

''<class name="TypeInfo" abstract tags="immutable">
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
Const CLASSID As String = "TypeInfo" '#DO NOT REMOVE IT
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
Public Enum TypeKind
    tkEnum = 0
    tkRecord = 1
    tkModule = 2
    tkInterface = 3
    tkDispatch = 4
    tkCoClass = 5
    tkAlias = 6
    tkUnion = 7
    tkMax = 8
End Enum
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

Public Property Get TypeKind() As TypeKind: End Property

Public Property Get Name() As String: End Property
Public Property Get Description() As String: End Property
Public Property Get Guid() As Guid: End Property
Public Property Get Version() As Version: End Property
Public Property Get UniqueID() As Long: End Property
Public Property Get TypeLib() As TypeLibInfo: End Property

Public Function ToString() As String: End Function

''<summary>
''  Returns a pseudo-unique hashed 4 bytes number to identify this object instance.
''</summary>
''<retval type="int32">The object unique hash code.</retval>
Public Function GetHashCode() As Long: End Function

''<summary>
''  Determines if passed argument is equal to the object instance.
''</summary>
''<params>
''<param name="Value" type="variant">
''</params>
''<retval type="bool">A boolean indicating the equality of this instance versus the passed in value.</retval>
Public Function Equals(Value As Variant) As Boolean: End Function

''<summary>
''  Return a metaobject for the object to control the object. null if object does'nt have meta object.
''</summary>
''<retval type="MintAPI.MetaObject">A meta object for the object instance.</retval>
Public Function MetaObject() As MetaObject: End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

