VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InterfaceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Member1" ,"Guid"
Attribute VB_Ext_KEY = "Member2" ,"Version"
'@PROJECT_LICENSE

''<class name="InterfaceInfo" tags="">
''  <summary>
''    Some class description.
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

Implements IObject

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "InterfaceInfo" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"
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
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Dim myTI As ITypeInfo
Dim myTI2 As ITypeInfo2
Dim myTLI As ITypeLib
Dim myTA As API_TYPEATTR, myTA_Inited As Boolean
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Constructors>
'
''<summary>
''  Basic constructor to use by MintAPI class provider,
''  When creating new instance of this class.
''</summary>
''<params>
''  <param name="Arguments">
''      A list of construction arguments.
''  </param>
''</params>
''<cargs order="0" count="1">
''  <carg type="long" name="ID"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal Object As Object)
    
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal TypeInfo As ITypeInfo)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal InterfaceInfo As InterfaceInfo)
    
    
    This_Constructed = True
End Sub
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

Public Property Get Alias() As String
    Alias = TypeLib.Alias & "::" & Name
End Property

Public Property Get Name() As String
THROW_IfNull
    Dim pName As String
    Call myTI.GetDocumentation(MEMBERID_NIL, pName, vbNullString, vbNullPtr, vbNullString)
    Name = pName
End Property
Public Property Get Description() As String
THROW_IfNull
    Dim pDesc As String
    Call myTI.GetDocumentation(MEMBERID_NIL, vbNullString, pDesc, vbNullPtr, vbNullString)
    Description = pDesc
End Property
Public Property Get HelpContext() As Long
THROW_IfNull
    Call myTI.GetDocumentation(MEMBERID_NIL, vbNullString, vbNullString, HelpContext, vbNullString)
End Property
Public Property Get HelpFile() As String
THROW_IfNull
    Dim pHF As String
    Call myTI.GetDocumentation(MEMBERID_NIL, vbNullString, vbNullString, vbNullPtr, pHF)
    HelpFile = pHF
End Property

Public Property Get Culture() As Culture
INIT_TypeAttr
    Set Culture = MintAPI.Culture.FromLCID(myTA.LCID)
End Property
Public Property Get Guid() As Guid
INIT_TypeAttr
    Set Guid = MintAPI.Guid.FromAPIGuid(myTA.Guid)
End Property
Public Property Get Version() As Version
INIT_TypeAttr
    Set Version = MintAPI.Version.Create(myTA.wMajorVerNum, myTA.wMinorVerNum)
End Property

Public Property Get TypeLib() As TypeLibInfo
THROW_IfNull
    
End Property

Public Property Get TypeKind() As TypeKind
THROW_IfNull
    TypeKind = tkInterface
End Property

Public Property Get UniqueID() As Long
    UniqueID = mString.HashString(TypeLib.Guid.ToString & Guid.ToString, Defaults.RangedLength)
End Property

Public Property Get Members() As Vector
THROW_IfNull
    
End Property
Public Property Get GetMember(ByVal Name As String) As MemberInfo
    
End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Basic Members>
'
Public Function Equals(Value As Variant) As Boolean
    Equals = Objects.Equals(Me, Value)
End Function
Public Function ToString() As String
    ToString = CLASSALIAS
End Function
Public Function GetHashCode() As Long
    GetHashCode = Objects.GetHashCode(Me)
End Function
Public Function MetaObject() As MetaObject
    If myMetaObject Is Nothing Then
        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    End If
    Set MetaObject = myMetaObject
End Function

Public Function Clone() As InterfaceInfo
    Set Clone = New InterfaceInfo
    Call Clone.ConstructorCopy(Me)
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
'Private Sub Class_Initialize()
'
'End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
'    Call IUnknown_Release(myTI)
'    Call memzero(myTI, VLEN_OBJECT)
End Sub

Private Sub THROW_IfNull()
'    If Not This_Constructed Then throw Exps.InvalidStatusException
'    If p_TinfI = -1 Then
'        If p_Object Is Nothing Then throw Exps.InvalidStatusException
'        If p_Disp Is Nothing Then throw Exps.InvalidStatusException
'    End If
End Sub

Private Function CastToIDispatch(ByVal p_Obj As Object) As IVBDispatch
'    Dim pPtr As Long
'    Call IUnknown_QueryInterface(p_Obj, MintAPI.Guid.IDispatch, pPtr, Nothing, True, True)
'    Call memcpy(CastToIDispatch, pPtr, VLEN_PTR)
End Function

Private Sub InitTypeInfo()
'    If Not myTI_Inited Then
'        If myTI Is Nothing Then
'            Dim NumOfTypeInfoImpl As Long
'            NumOfTypeInfoImpl = p_Disp.GetTypeInfoCount
'            If NumOfTypeInfoImpl = 0 Then throw Exps.InvalidStatusException("The Interface Does Not Provide A ITypeInfo Interface.")
'            Call memcpy(myTI, ByVal p_Disp.GetTypeInfo(0, LOCALE_USER_DEFAULT), VLEN_PTR)
'            Call IUnknown_AddRef(myTI)
'        End If
'    End If
End Sub

Private Sub InitTypeInfo2()
'    If Not myTI2_Inited Then
'        If myTI Is Nothing Then
'            Call memcpy(myTI, ByVal p_Disp.GetTypeInfo(0, LOCALE_USER_DEFAULT), VLEN_PTR)
'            Call IUnknown_AddRef(myTI)
'        End If
'    End If
End Sub

Private Sub InitTypeLibInfo()
    
End Sub

Private Sub INIT_TypeAttr()
'    If (Not myTA_Inited) Or (p_TA_RefreshMode) Then
'        On Error GoTo CatchError
'        Dim pPtr As Long
'        Call InitTypeInfo
'        Call myTI.GetTypeAttr(pPtr)
'        Call memcpy(myTA, ByVal pPtr, Len(myTA))
'        Call myTI.ReleaseTypeAttr(pPtr)
'        myTA_Inited = True
'        Exit Sub
'CatchError:
'        throw Exps.SystemCallFailureException(Err.Description)
'    End If
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo): Call Constructor(CInfo): End Sub
Private Function IObject_Equals(Value As Variant) As Boolean: IObject_Equals = Equals(Value): End Function
Private Function IObject_GetHashCode() As Long: IObject_GetHashCode = GetHashCode: End Function
Private Function IObject_MetaObject() As MetaObject: Set IObject_MetaObject = MetaObject: End Function
Private Function IObject_ToString() As String: IObject_ToString = ToString: End Function
Private Function IObject_Stub(ByVal ProxyInfo As ObjectProxyInfo) As Boolean
    
    IObject_Stub = False
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
