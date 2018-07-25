VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MetaObjectStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MetaObject" static tags="">
''  <summary>
''    Provides static methods for MetaObject.
''  </summary>
''  <using>MetaObject</using>
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
Const CLASSID As String = "MetaObject" '#DO NOT REMOVE IT
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
''<default/>
Public Function Create(ByRef MetaObject As MetaObject, ByVal Object As Object, ByVal NeedSignalEmitter As Boolean, ByVal VersionString As String) As MetaObject
Attribute Create.VB_UserMemId = 0
    Set MetaObject = New MetaObject
    Call MetaObject.Constructor0(Object, NeedSignalEmitter, VersionString)
    Set Create = MetaObject
End Function
Public Function Setup(ByRef MetaObject As MetaObject, ByVal Object As IObject, ByVal NeedSignalEmitter As Boolean, ByVal VersionString As String) As MetaObject
    If MetaObject Is Nothing Then
        Call Create(MetaObject, Object, NeedSignalEmitter, VersionString)
    End If
    Set Setup = MetaObject
End Function
Public Function SimpleMetaObject(ByRef MetaObject As MetaObject, ByVal Object As Object, VersionString As String) As MetaObject
    
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

Public Function MetaObjectOf(ByVal Object As IObject) As MetaObject
    Set MetaObjectOf = Object.MetaObject
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
