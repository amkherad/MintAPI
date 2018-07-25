VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypeInfoStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="TypeInfoStatic" static tags="">
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
Const CLASSID As String = "TypeInfoStatic" '#DO NOT REMOVE IT
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
Public Function TypeLibInfoFromFile(ByVal TypeLibPath As String) As TypeLibInfo
    Dim TypeLibPtr As ITypeLib
    Call API_LoadTypeLib(TypeLibPath, TypeLibPtr)
    Set TypeLibInfoFromFile = New TypeLibInfo
    Call TypeLibInfoFromFile.Constructor0(TypeLibPtr)
End Function
Public Function TypeLibInfoFromRegistry(ByVal Guid As Guid, ByVal Version As Version, ByVal LCID As Long) As TypeLibInfo
    If Guid Is Nothing Then throw Exps.ArgumentNullException
    If Version Is Nothing Then throw Exps.ArgumentNullException
    Dim TypeLibPtr As ITypeLib
    Call API_LoadRegTypeLib(Guid.ToAPIStdGuid, Version.Major, Version.Minor, LCID, TypeLibPtr)
    Set TypeLibInfoFromRegistry = New TypeLibInfo
    Call TypeLibInfoFromRegistry.Constructor0(TypeLibPtr)
End Function
Public Function InterfaceInfoFromObject(ByVal Object As Object) As InterfaceInfo
    If Object Is Nothing Then throw Exps.ArgumentNullException
    Set InterfaceInfoFromObject = New InterfaceInfo
    Call InterfaceInfoFromObject.Constructor0(Object)
End Function
Public Function ClassInfoFromObject(ByVal Object As Object) As CoClassInfo
    If Object Is Nothing Then throw Exps.ArgumentNullException
    Set ClassInfoFromObject = New CoClassInfo
    Call ClassInfoFromObject.Constructor0(Object)
End Function
Public Function RecordInfoFromRecord(ByRef Expression As Variant) As RecordInfo
    If Not Info.IsRecord(Expression) Then throw Exps.InvalidArgumentException
    
End Function
Public Function ParameterInfoFromRecord(ByRef Expression As Variant) As ParameterInfo
    If Not Info.IsRecord(Expression) Then throw Exps.InvalidArgumentException
    
End Function

'<default/>
Public Function FromVariable(ByRef Expression As Variant) As TypeInfo
Attribute FromVariable.VB_UserMemId = 0
    If IsMissing(Expression) Then throw Exps.ArgumentNullException
    If IsObject(Expression) Then
        Set FromVariable = ClassInfoFromObject(Expression)
    ElseIf Info.IsRecord(Expression) Then
        Set FromVariable = ClassInfoFromObject(Expression)
    Else
        
    End If
End Function

Friend Function GetTypeInfoOfITypeInfo(ByVal ITI As ITypeInfo2) As TypeInfo
    If ITI Is Nothing Then throw Exps.ArgumentNullException
    
    Select Case ITI.GetTypeKind
        Case tkAlias
            
        Case tkCoClass
            Dim TICoClass As New CoClassInfo
            Call TICoClass.Constructor1(ITI)
            Set GetTypeInfoOfITypeInfo = TICoClass
        Case tkDispatch
            
        Case tkEnum
            
        Case tkInterface
            Dim TIInterface As New InterfaceInfo
            Call TIInterface.Constructor1(ITI)
            Set GetTypeInfoOfITypeInfo = TIInterface
        Case tkModule
            
        Case tkRecord
            Dim TIRecord As New RecordInfo
            Call TIRecord.Constructor1(ITI)
            Set GetTypeInfoOfITypeInfo = TIRecord
        Case tkUnion
            
        'Case tkMax
            
        Case Else
            throw Exps.InvalidOperationException
    End Select
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



'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
