VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypeLibInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Member1" ,"Version"
Attribute VB_Ext_KEY = "Member2" ,"Guid"
Attribute VB_Ext_KEY = "Member3" ,"Culture"
'@PROJECT_LICENSE

''<class name="TypeLibInfo" tags="immutable">
''  <summary>
''    Provides information about a type library.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>TypeInfo, InterfaceInfo</using>
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
Const CLASSID As String = "TypeLibInfo" '#DO NOT REMOVE IT
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
Public Enum TypeLibFlags
    tlfRestricted = 1
    tlfControl = 2
    tlfHidden = 4
    tlfHasDiskImage = 8
End Enum
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

Dim myTLI As ITypeLib
Dim myTLibAttr As API_TLIBATTR, myTLibAttr_Inited As Boolean
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
Friend Sub Constructor0(ByVal TLibPtr As ITypeLib)
    Call memcpy(myTLI, TLibPtr, VLEN_PTR)
    Call IUnknown.AddRef(myTLI)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal TypeLibInfo As TypeLibInfo)
    
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
    
End Property

Public Property Get Name() As String
THROW_IfNull
    Dim pName As String
    Call myTLI.GetDocumentation(MEMBERID_NIL, pName, vbNullString, vbNullString, vbNullString)
    Name = pName
End Property

Public Property Get Description() As String
THROW_IfNull
    Dim pDesc As String
    Call myTLI.GetDocumentation(MEMBERID_NIL, vbNullString, pDesc, vbNullString, vbNullString)
    Description = pDesc
End Property

Public Property Get HelpFile() As String
THROW_IfNull
    Dim pHF As String
    Call myTLI.GetDocumentation(MEMBERID_NIL, vbNullString, vbNullString, vbNullString, pHF)
    HelpFile = pHF
End Property

Public Property Get Flags() As TypeLibFlags
INIT_TypeLibAttr
    Flags = myTLibAttr.wLibFlags
End Property

Public Property Get IsOnDisk() As Boolean
    IsOnDisk = ((Flags And tlfHasDiskImage) = tlfHasDiskImage)
End Property


Public Property Get Culture() As Culture
INIT_TypeLibAttr
    Set Culture = MintAPI.Culture.FromLCID(myTLibAttr.LCID)
End Property

Public Property Get Guid() As Guid: THROW_IfNull
INIT_TypeLibAttr
    Set Guid = MintAPI.Guid.FromAPIGuid(myTLibAttr.Guid)
End Property

Public Property Get Version() As Version
INIT_TypeLibAttr
    Set Version = MintAPI.Version.Create(myTLibAttr.wMajorVerNum, myTLibAttr.wMinorVerNum)
End Property

Public Function IsSameTypeLib(ByVal TypeLibInfo As TypeLibInfo) As Boolean
THROW_IfNull
    If TypeLibInfo Is Nothing Then throw Exps.ArgumentNullException
    IsSameTypeLib = Guid.Equals(TypeLibInfo.Guid)
End Function

Public Property Get Members() As Vector
    Set Members = EnumMembers(tkMax)
End Property
Public Property Get Interfaces() As Vector
    Set Interfaces = EnumMembers(tkInterface)
End Property
Public Property Get CoClasses() As Vector
    Set CoClasses = EnumMembers(tkCoClass)
End Property
Public Property Get Enumerables() As Vector
    Set Enumerables = EnumMembers(tkEnum)
End Property
Public Property Get Records() As Vector
    Set Records = EnumMembers(tkRecord)
End Property
Public Property Get Dispatches() As Vector
    Set Dispatches = EnumMembers(tkDispatch)
End Property

Public Function GetTypeInfo(ByRef Index As Variant) As TypeInfo

End Function

Public Sub Register()

End Sub
Public Sub UnRegister()

End Sub

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

Public Function Clone() As TypeLibInfo
    Set Clone = New TypeLibInfo
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
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    If myTLibAttr_Inited Then _
        If Not myTLI Is Nothing Then _
            Call myTLI.ReleaseTLibAttr(myTLibAttr)
End Sub

Private Sub INIT_TypeLibAttr(Optional ByVal Reload As Boolean = False)
    If myTLI Is Nothing Then throw Exps.InvalidOperationException
    If myTLibAttr_Inited And (Not Reload) Then Exit Sub
    Call myTLI.GetLibAttr(myTLibAttr)
    myTLibAttr_Inited = True
End Sub

Private Sub THROW_IfNull()
    If myTLI Is Nothing Then throw Exps.InvalidOperationException
End Sub

Private Function EnumMembers(ByVal TypeKind As TypeKind) As Vector
    Dim i As Long, TotalTypes As Long
    Dim RetList As New List
    
    Dim TpInfo As ITypeInfo, cTpInfo As TypeInfo
    
    TotalTypes = myTLI.GetTypeInfoCount
    
    While i < TotalTypes
        Call myTLI.GetTypeInfo(i, TpInfo)
        
        Set cTpInfo = TypeInfo.GetTypeInfoOfITypeInfo(TpInfo)
        If cTpInfo.TypeKind = TypeKind Then _
            Call RetList.Append(cTpInfo)
    Wend
    
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
