VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="DynamicObject" tags="">
''  <summary>
''    Represents a dynamic object.
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

'##################@@@@@@@@@@@@@@@@@@ Function replacement - becarefull!

Option Base 0
Option Explicit

Implements IObject
Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "DynamicObject" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"

Const EXPS_GetEnumeratorMissing         As String = "GetEnumerator is missing on dynamic object's members."
Const EXPS_IEnumeratorExpected          As String = "An IEnumerator value expected."
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

Dim p_Name As String
Dim p_ThisGuid As Guid
Dim p_Guid As Guid
Dim p_Guids As New List '<Guid>

Dim p_Delegate As Object 'Maybe Null
Dim p_DelegateGuid As Guid 'Maybe Null

Dim p_Members As New Dictionary
Dim p_OldVFTable As New Dictionary

Dim p_MyTypeInfo As DynamicClassInfo
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
Friend Sub Constructor0(ByVal Name As String, ByVal Obj As Object)
    p_Name = Name
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As DynamicObject)

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

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
THROW_IfNull
    Dim MemberInfo As MemberInfo, Method As Method
    
    If Not p_Members.TryGetValue(FUNCNAME_IEnumerable_GetEnumerator, MemberInfo) Then _
        throw Exps.InvalidOperationException(EXPS_GetEnumeratorMissing)
    
    Set Method = MemberInfo.Method
    
    Dim Result As Variant, ObjResult As Object
    Set Result = Method()
    
    If IsEmptyVariable(Result) Then throw Exps.InvalidOperationException(EXPS_IEnumeratorExpected)
    If Not IsObject(Result) Then throw Exps.InvalidOperationException(EXPS_IEnumeratorExpected)
    Set ObjResult = Result
    If Not TypeOf ObjResult Is IEnumerator Then throw Exps.InvalidOperationException(EXPS_IEnumeratorExpected)
    
    Dim Enumerator As IEnumerator
    Set Enumerator = Result
    
    Set NewEnum = CreateEnumerator(Enumerator)
End Function

Public Property Get Name() As String
THROW_IfNull
    Name = p_Name
End Property

Public Property Get Guid() As Guid
THROW_IfNull
    Set Guid = p_Guid
End Property

Public Property Get Members() As Vector
THROW_IfNull
    
End Property

Public Function GetMember(ByVal Name As String) As MemberInfo
THROW_IfNull
    
End Function

Public Sub AddMember(ByVal Name As String, ByVal Method As Method, ByVal MemberType As MemberType)
Attribute AddMember.VB_MemberFlags = "40"
THROW_IfNull
    
End Sub
Public Sub RemoveMember(ByVal Name As String)
THROW_IfNull
    
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
    If TypeOf p_Delegate Is IObject Then
        Dim IO As IObject
        Set IO = p_Delegate
        Equals = IO.Equals(Value)
    ElseIf Not p_Delegate Is Nothing Then
        Equals = Objects.ReferenceEquals(p_Delegate, Value)
    Else
        Equals = Objects.Equals(Me, Value)
    End If
End Function
Public Function ToString() As String
    If TypeOf p_Delegate Is IObject Then
        Dim IO As IObject
        Set IO = p_Delegate
        ToString = IO.ToString()
    ElseIf Not p_Delegate Is Nothing Then
        ToString = "DynamicObject<" & Objects.ClassInfoFromObject(p_Delegate).Alias & ">"
    Else
        ToString = CLASSALIAS
    End If
End Function
Public Function GetHashCode() As Long
    If TypeOf p_Delegate Is IObject Then
        Dim IO As IObject
        Set IO = p_Delegate
        GetHashCode = IO.GetHashCode
    ElseIf Not p_Delegate Is Nothing Then
        GetHashCode = Objects.GetHashCode(p_Delegate)
    Else
        GetHashCode = Objects.GetHashCode(Me)
    End If
End Function
Public Function MetaObject() As MetaObject
    
End Function

Public Function Clone() As DynamicObject
    Set Clone = New DynamicObject
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
Private Sub Class_Initialize()
    Dim VTbl As Long, VTblMember As Long
    Call memcpy(VTbl, ByVal ObjPtr(Me), VLEN_PTR)
    
    Set p_Guid = MintAPI.Guid.Create
    
    '======================
    Call memcpy(VTblMember, _
        ByVal (VTbl + (FUNC_ORDER_QueryInterface * VLEN_PTR)), VLEN_PTR)
    Call p_OldVFTable.Add(FUNCNAME_IUNKNOWN_QueryInterface, VTblMember)
    '----
    Call memcpy(ByVal (VTbl + (FUNC_ORDER_QueryInterface * VLEN_PTR)), _
        AddressOf mint_internals.mint_internals_DynamicObject_QueryInterface, VLEN_PTR)
    '======================
    
    '======================
    Call memcpy(VTblMember, _
        ByVal (VTbl + (FUNC_ORDER_GetTypeInfoCount * VLEN_PTR)), VLEN_PTR)
    Call p_OldVFTable.Add(FUNCNAME_IDISPATCH_GetTypeInfoCount, VTblMember)
    '----
    Call memcpy(ByVal (VTbl + (FUNC_ORDER_GetTypeInfoCount * VLEN_PTR)), _
        AddressOf mint_internals.mint_internals_DynamicObject_GetTypeInfoCount, VLEN_PTR)
    '======================
    
    '======================
    Call memcpy(VTblMember, _
        ByVal (VTbl + (FUNC_ORDER_GetTypeInfo * VLEN_PTR)), VLEN_PTR)
    Call p_OldVFTable.Add(FUNCNAME_IDISPATCH_GetTypeInfo, VTblMember)
    '----
    Call memcpy(ByVal (VTbl + (FUNC_ORDER_GetTypeInfo * VLEN_PTR)), _
        AddressOf mint_internals.mint_internals_DynamicObject_GetTypeInfo, VLEN_PTR)
    '======================
    
    '======================
    Call memcpy(VTblMember, _
        ByVal (VTbl + (FUNC_ORDER_GetIDsOfNames * VLEN_PTR)), VLEN_PTR)
    Call p_OldVFTable.Add(FUNCNAME_IDISPATCH_GetIDsOfNames, VTblMember)
    '----
    Call memcpy(ByVal (VTbl + (FUNC_ORDER_GetIDsOfNames * VLEN_PTR)), _
        AddressOf mint_internals.mint_internals_DynamicObject_GetIDsOfNames, VLEN_PTR)
    '======================
    
    '======================
    Call memcpy(VTblMember, _
        ByVal (VTbl + (FUNC_ORDER_Invoke * VLEN_PTR)), VLEN_PTR)
    Call p_OldVFTable.Add(FUNCNAME_IDISPATCH_Invoke, VTblMember)
    '----
    Call memcpy(ByVal (VTbl + (FUNC_ORDER_Invoke * VLEN_PTR)), _
        AddressOf mint_internals.mint_internals_DynamicObject_Invoke, VLEN_PTR)
    '======================
End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Private Property Get ThisGuid() As Guid
    If p_ThisGuid Is Nothing Then _
        Set p_ThisGuid = TypeInfo.ClassInfoFromObject(Me).Guid
    Set ThisGuid = p_ThisGuid
End Property

Private Sub THROW_IfNull()
    
End Sub

Friend Function DelegateGuid() As Guid
    Set DelegateGuid = p_DelegateGuid
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Callbacks>
'
Public Function QueryInterface( _
    ByRef riid As StdGuid, _
    ByRef pvObj As Long) As Long
    '-----------------------------
    Dim OK As Long
    On Error Resume Next
    
    Dim MeGuid As Guid
    Set MeGuid = ThisGuid
    If Not MeGuid Is Nothing Then _
        If API_IsEqualGUID(riid, MeGuid.ToStdGuid) Then _
            OK = vbTrue: GoTo OKChecked
    
    If Not p_DelegateGuid Is Nothing Then _
        If API_IsEqualGUID(riid, p_DelegateGuid.ToStdGuid) Then _
            OK = vbTrue: GoTo OKChecked
    
    If API_IsEqualGUID(riid, MintAPI.Guid.IUnknown.ToStdGuid) Then
        OK = vbTrue
    ElseIf API_IsEqualGUID(riid, p_Guid.ToStdGuid) Then
        OK = vbTrue
    Else
        Dim G As Guid
        For Each G In p_Guids
            If Not G Is Nothing Then
                If API_IsEqualGUID(riid, G.ToStdGuid) Then
                    OK = vbTrue
                    Exit For
                End If
            End If
        Next
    End If
    
OKChecked:
    If OK Then
        Dim ObjPtr As Long
        Call memcpy(ObjPtr, Me, VLEN_PTR)
        pvObj = ObjPtr
        Call IUnknown.AddRefPtr(ObjPtr)
    Else
        QueryInterface = E_NOINTERFACE
    End If
End Function
Public Function GetTypeInfo( _
    ByRef iTInfo As ITypeInfo, _
    ByRef LCID As Long) As ITypeInfo
Attribute GetTypeInfo.VB_MemberFlags = "40"
    '-----------------------------
    
    MsgBox "GetTypeInfo"
    
End Function
Public Function GetIDsOfNames( _
    ByRef riid As StdGuid, _
    ByRef rgszNames() As String, _
    ByRef cNames As Long, _
    ByRef LCID As Long) As Long
Attribute GetIDsOfNames.VB_MemberFlags = "40"
    '-----------------------------
    Dim OK As Long
    
    MsgBox "GetIDsOfNames"
    
End Function
Public Function GetTypeInfoCount() As Long
Attribute GetTypeInfoCount.VB_MemberFlags = "40"
    
    MsgBox "GetTypeInfoCount"
    
End Function
Public Function Invoke( _
    ByRef dispIdMember As Long, _
    ByRef riid As StdGuid, _
    ByRef LCID As Long, _
    ByRef wFlags As Long, _
    ByRef pDispParams As API_DISPPARAMS, _
    ByRef pVarResult As Variant, _
    ByRef pExcepInfo As API_EXCEPINFO) As Long
    '-----------------------------
    
    MsgBox "Invoke"
    
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
