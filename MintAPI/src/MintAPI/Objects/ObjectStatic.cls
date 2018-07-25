VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ObjectStatic" static tags="">
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
Const CLASSID As String = "ObjectStatic" '#DO NOT REMOVE IT
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
'<section Variables/Events>
'

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Generic Methods Helpers>
'
Public Function Equals(ByVal Object As Object, ByRef ValueWhat As Variant) As Boolean
    If TypeOf Object Is IObject Then
        Dim IObj As IObject
        Set IObj = Object
        Equals = IObj.Equals(ValueWhat)
    End If
End Function
Public Function ReferenceEquals(ByVal Object1 As Object, ByVal Object2 As Object) As Boolean
    ReferenceEquals = (ObjPtr(Object1) = ObjPtr(Object2))
End Function
Public Function GetHashCode(ByVal Object As IObject) As Long
    GetHashCode = ObjPtr(CUnk(Object))
End Function
Public Function ToString(ByVal Object As IObject) As String
    ToString = Object.ToString
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


'!DO NOT REFERENCE TO TypeInfo BASED CLASSES.
Public Function IsEqual(ByVal Object As Object, ByVal Guid As Guid) As Boolean
    
End Function
''<summary>Initializes the object using IObject.Initialize method.</summary>
''<remarks>
''  This is very usefull when trying to initialize a loaded object that not initialized.
''  <br>
''  This function used when instanciate a class using MintAPI internal type provider.
''</remarks>
Public Sub CallBasicConstructor(ByVal Object As IObject, ByVal Arguments As ArgumentList)
    Call Object.Construct(Nothing)
End Sub
Public Function Alias(ByVal Object As IObject) As String
    Alias = Object.MetaObject.Alias
End Function
Public Function Notify(ByVal Object As IObject, ByVal Name As String, ByVal Arguments As ArgumentList) As Variant
    
End Function

Public Function MetaObject() As MetaObjectStatic
    Set MetaObject = mint_static_mngr.MetaObject
End Function


Public Function QueryInformation(ByVal Obj As IObject, ByVal QueryName As String, ParamArray Args() As Variant) As ArgumentList
'    If Obj Is Nothing Then throw Exps.ArgumentNullException
'    Dim OQR As New ObjectQR
'    Call OQR.Constructor0(QueryName, AL.FromParamArray(8))
'    If Obj.Stub(OQR) Then
'        Set QueryInformation = OQR.Responses
'    Else
'        throw Exps.QueryInformationFailedException
'    End If
End Function

Public Function InterfaceInfoFromObject(ByVal Object As Object) As InterfaceInfo
    Set InterfaceInfoFromObject = TypeInfo.InterfaceInfoFromObject(Object)
End Function
Public Function ClassInfoFromObject(ByVal Object As Object) As CoClassInfo
    Set ClassInfoFromObject = TypeInfo.ClassInfoFromObject(Object)
End Function

Public Function CreateInstance(ByVal AliasString As String, Optional ByVal ServerName As String = vbNullString) As Object
    Set CreateInstance = VBA.CreateObject(AliasString, ServerName)
End Function
Public Function CreateInstanceFromFile(ByVal Path As String, ByVal ClassName As String) As Object
    Set CreateInstanceFromFile = VBA.GetObject(Path, ClassName)
End Function
Public Function CreateInstanceFromGUID(ByVal Guid As Guid) As Object
    'Set CreateInstanceFromGUID = VBA.GetObject(FilePath, ClassName)
End Function

Public Function Invoke(ByVal Object As Object, ByVal MemberID As Long) As Variant
    
End Function

Public Function Provide(ByVal ProviderEngine As String, ByVal Alias As String, ByRef outReturn As Object) As Boolean
    
End Function
Public Function ProvideImplicit(ByVal ProviderEngine As String, ByVal Alias As String) As Object
Attribute ProvideImplicit.VB_UserMemId = 0
    
End Function
Public Function ProvideStaticObject(ByVal ProviderEngine As String, ByVal Alias As String, ByRef outReturn As Object) As Boolean
    
End Function

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Tools>
'


'Public Function CreateObjectBuffer(ByVal Name As String, ParamArray Args() As Variant) As ObjectBuffer
'    Dim Arg() As Variant
'    Arg = Args
'    Dim OB As New ObjectBuffer
'    Call OB.InitializeW(Name, Arg)
'    Set CreateObjectBuffer = OB
'End Function
'Public Function CreateObjectBufferC(ByVal Name As String, Args() As Variant) As ObjectBuffer
'    Dim OB As New ObjectBuffer
'    Call OB.InitializeW(Name, Args)
'    Set CreateObjectBufferC = OB
'End Function
'Public Function CreateHandleObject(ByVal Handle As Long) As Object
'    Set CreateHandleObject = CreateObjectBuffer(OBJECTBUFFER_HANDLE, Handle)
'End Function
'Public Function CreateHDCObject(ByVal hdc As Long) As Object
'    Set CreateHDCObject = CreateObjectBuffer(OBJECTBUFFER_DISPLAY_CONTEXT, hdc)
'End Function
'Public Function CreateRenderTargetObject(Target As Variant) As Object
'    Set CreateRenderTargetObject = CreateObjectBuffer(OBJECTBUFFER_RENDERTARGET, Target)
'End Function
'Public Function Inherit(Optional ByVal InheritFrom As Object = Nothing) As Object
'    Set Inherit = CreateObjectBuffer(OBJECTBUFFER_INHERIT, InheritFrom)
'End Function



Public Sub Override(ByVal Destination As Method, ByVal Source As Method)
    If Destination Is Nothing Then throw Exps.ArgumentNullException
    If Source Is Nothing Then throw Exps.ArgumentNullException
    
    If Not Destination.IsMember Then throw Exps.InvalidOperationException
    If Not Source.IsMember Then throw Exps.InvalidOperationException
    
    
End Sub
Public Function AsDynamic(ByVal Obj As Object) As DynamicObject
    Set AsDynamic = DynamicObject.Create(Obj)
End Function


Public Function HiddenContext(ByVal Context As Variant) As Object
    Dim HC As New HiddenContext
    Call HC.Constructor0(Context)
    Set HiddenContext = HC
End Function
Public Function IsHiddenContext(ByVal ObjContext As Object) As Boolean
    If ObjContext Is Nothing Then Exit Function
    IsHiddenContext = (TypeOf ObjContext Is HiddenContext)
End Function
Public Function GetHiddenContext(ByVal ObjContext As Object) As Variant
    If ObjContext Is Nothing Then throw Exps.ArgumentNullException("ObjContext")
    If Not TypeOf ObjContext Is HiddenContext Then throw Exps.InvalidArgumentException("ObjContext")
    Dim HC As HiddenContext
    Set HC = ObjContext
    Call EvaluateX(GetHiddenContext, HC.Context)
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
