VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoClassInfo"
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

''<class name="CoClassInfo" tags="">
''  <summary>
''    Provides information about COM classes that have implemented the IDispatch.
''  </summary>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
Implements ICloneable
Implements TypeInfo
Implements InterfaceInfo

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "CoClassInfo" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"
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
Friend Sub Constructor0(ByVal Obj As Object)
    Dim Disp As IVBDispatch
    
    Call memcpy(Disp, Obj, VLEN_PTR)
    
    Dim TypeInfoPtr As Long
    Call memcpy(TypeInfoPtr, Disp.GetTypeInfo(0, LOCALE_USER_DEFAULT), VLEN_PTR)
    If TypeInfoPtr = vbNullPtr Then throw Exps.InvalidOperationException
    Call IUnknown.AddRefPtr(TypeInfoPtr)
    Call memcpy(myTI, TypeInfoPtr, VLEN_PTR)
    
    Call memzero(Disp, VLEN_PTR)
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal TypeInfo As ITypeInfo)
    Dim TypeInfoPtr As Long
    Call memcpy(TypeInfoPtr, TypeInfo, VLEN_PTR)
    Call IUnknown.AddRefPtr(TypeInfoPtr)
    Call memcpy(myTI, TypeInfoPtr, VLEN_PTR)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As CoClassInfo)
    
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
THROW_IfNull
    Alias = TypeLib.Alias & "::" & Name & "/" & Version.ToString
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
    Dim SG As API_StdGuid
    Call memcpy(SG, ByVal myTA.Guid.Data1, 16)
    Set Guid = MintAPI.Guid.FromAPIGuid(SG)
End Property
Public Property Get Version() As Version
INIT_TypeAttr
    Set Version = MintAPI.Version.Create(myTA.wMajorVerNum, myTA.wMinorVerNum)
End Property

Public Property Get TypeLib() As TypeLibInfo
THROW_IfNull
    If myTLI Is Nothing Then
        Call myTI.GetContainingTypeLib(myTLI, MEMBERID_NIL)
    End If
    Set TypeLib = New TypeLibInfo
    Call TypeLib.Constructor0(myTLI)
End Property

Public Property Get TypeKind() As TypeKind
THROW_IfNull
    TypeKind = tkCoClass
End Property

Public Property Get UniqueID() As Long
    UniqueID = mString.HashString(TypeLib.Guid.ToString & Guid.ToString, Defaults.RangedLength)
End Property

Public Property Get DefaultInterface() As InterfaceInfo
    
End Property

Public Property Get Members() As Vector
THROW_IfNull
    Set Members = EnumMembers(tkMax)
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

Public Function Clone() As CoClassInfo
    Set Clone = New CoClassInfo
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
'    ...
'End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    If myTA_Inited Then _
        If Not myTI Is Nothing Then _
            Call myTI.ReleaseTypeAttr(myTA)
    Set myTI = Nothing
End Sub

Private Sub INIT_TypeAttr(Optional ByVal Reload As Boolean = False)
    If myTI Is Nothing Then throw Exps.InvalidOperationException
    If myTA_Inited Then
        If (Not Reload) Then Exit Sub
        Call myTI.ReleaseTypeAttr(myTA)
    End If
    
    Call myTI.GetTypeAttr(myTA)
    myTA_Inited = True
End Sub

Private Sub THROW_IfNull()
    If myTI Is Nothing Then throw Exps.InvalidOperationException
End Sub

Private Function EnumMembers(ByVal TypeKind As TypeKind) As Vector
INIT_TypeAttr
    Dim i As Long, TotalTypes As Long, tk As Long
    Dim RetList As New List, Funcs As Long
    
    Funcs = myTA.cFuncs
    For i = 0 To Funcs - 1
        
    Next
    
    Set EnumMembers = RetList.ToVector
End Function
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


'=============================================
'=============================================
'=============================================
'<section TypeInfo Implementations>
'
Private Property Get TypeInfo_Name() As String: TypeInfo_Name = Name: End Property
Private Property Get TypeInfo_Description() As String: TypeInfo_Description = Description: End Property
Private Property Get TypeInfo_Guid() As Guid: Set TypeInfo_Guid = Guid: End Property
Private Property Get TypeInfo_TypeKind() As TypeKind: TypeInfo_TypeKind = TypeKind: End Property
Private Property Get TypeInfo_Version() As Version: Set TypeInfo_Version = Version: End Property
Private Property Get TypeInfo_TypeLib() As TypeLibInfo: Set TypeInfo_TypeLib = TypeLib: End Property
Private Property Get TypeInfo_UniqueID() As Long: TypeInfo_UniqueID = UniqueID: End Property
Private Function TypeInfo_Equals(Value As Variant) As Boolean: TypeInfo_Equals = Equals(Value): End Function
Private Function TypeInfo_ToString() As String: TypeInfo_ToString = ToString: End Function
Private Function TypeInfo_GetHashCode() As Long: TypeInfo_GetHashCode = GetHashCode: End Function
Private Function TypeInfo_MetaObject() As MetaObject: Set TypeInfo_MetaObject = MetaObject: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------


'=============================================
'=============================================
'=============================================
'<section InterfaceInfo Implementations>
'
Private Function InterfaceInfo_Clone() As InterfaceInfo: Set InterfaceInfo_Clone = Clone: End Function
Private Function InterfaceInfo_Equals(Value As Variant) As Boolean: InterfaceInfo_Equals = Equals(Value): End Function
Private Function InterfaceInfo_GetHashCode() As Long: InterfaceInfo_GetHashCode = GetHashCode: End Function
Private Function InterfaceInfo_MetaObject() As MetaObject: Set InterfaceInfo_MetaObject = MetaObject: End Function
Private Function InterfaceInfo_ToString() As String: InterfaceInfo_ToString = ToString: End Function
Private Property Get InterfaceInfo_Alias() As String: InterfaceInfo_Alias = Alias: End Property
Private Property Get InterfaceInfo_Culture() As Culture: Set InterfaceInfo_Culture = Culture: End Property
Private Property Get InterfaceInfo_Description() As String: InterfaceInfo_Description = Description: End Property
Private Property Get InterfaceInfo_Guid() As Guid: Set InterfaceInfo_Guid = Guid: End Property
Private Property Get InterfaceInfo_HelpContext() As Long: InterfaceInfo_HelpContext = HelpContext: End Property
Private Property Get InterfaceInfo_HelpFile() As String: InterfaceInfo_HelpFile = HelpFile: End Property
Private Property Get InterfaceInfo_GetMember(ByVal Name As String) As MemberInfo: Set InterfaceInfo_GetMember = GetMember(Name): End Property
Private Property Get InterfaceInfo_Members() As Vector: Set InterfaceInfo_Members = Members: End Property
Private Property Get InterfaceInfo_Name() As String: InterfaceInfo_Name = Name: End Property
Private Property Get InterfaceInfo_TypeKind() As TypeKind: InterfaceInfo_TypeKind = TypeKind: End Property
Private Property Get InterfaceInfo_TypeLib() As TypeLibInfo: Set InterfaceInfo_TypeLib = TypeLib: End Property
Private Property Get InterfaceInfo_UniqueID() As Long: InterfaceInfo_UniqueID = UniqueID: End Property
Private Property Get InterfaceInfo_Version() As Version: Set InterfaceInfo_Version = Version: End Property
'
'</section>
'---------------------------------------------
'---------------------------------------------
