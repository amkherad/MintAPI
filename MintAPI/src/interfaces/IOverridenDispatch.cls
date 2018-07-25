VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOverridenDispatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="IOverridenDispatch" interface tags="">
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
Const CLASSID As String = "IOverridenDispatch" '#DO NOT REMOVE IT
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

Public Function GetTypeInfo( _
    ByRef iTInfo As ITypeInfo, _
    ByRef LCID As Long) As ITypeInfo: End Function

Public Function GetIDsOfNames( _
    ByRef riid As StdGuid, _
    ByRef rgszNames() As String, _
    ByRef cNames As Long, _
    ByRef LCID As Long) As Long: End Function

Public Function GetTypeInfoCount() As Long: End Function

Public Function Invoke( _
    ByRef dispIdMember As Long, _
    ByRef riid As StdGuid, _
    ByRef LCID As Long, _
    ByRef wFlags As Long, _
    ByRef pDispParams As API_DISPPARAMS, _
    ByRef pVarResult As Variant, _
    ByRef pExcepInfo As API_EXCEPINFO) As Long: End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

