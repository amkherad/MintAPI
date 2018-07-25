VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Method"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Member1" ,"ParametersScheme"
'@PROJECT_LICENSE

''<class name="Method" tags="proxy_immutable">
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
Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Method" '#DO NOT REMOVE IT
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
Public Enum CallingConvensions
    ccStdCall
    ccAPICall = ccStdCall
    ccCdecl
    ccSysCall
    ccFastCall
    ccThisCall 'Set ECX = Class ptr.
    ccPascal
    'ccMint
    ccDotNetStandard
    ccExplicitThis
    ccHasThis
    ccVarArgs
    ccDotNetAny
    ccCustom = 99
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

Private p_Name As String
Private p_Description As String
Private p_TargetFunction As Long
Private p_Prototype As MethodPrototype
Private p_IsMember As Boolean
Private p_MemberOf As Object '!MUST BE! strong reference
Private p_IsOnTheFly As Boolean
Private p_OnTheFlyLength As Long
Private p_IsExecutable As Boolean 'Cached copy!!
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
Friend Sub Constructor0(ByVal Name As String, ByVal TargetFunction As Long, _
                        ByVal Prototype As MethodPrototype, ByVal IsMember As Boolean, _
                        Optional ByVal MemberOf As Object)
    p_Name = Name
    p_TargetFunction = TargetFunction
    Set p_Prototype = Prototype
    p_IsMember = IsMember
    Set p_MemberOf = MemberOf
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As Method)
    p_Name = That.Name
    p_TargetFunction = That.Reference
    Set p_Prototype = That.Prototype
    p_IsMember = That.IsMember
    Set p_MemberOf = That.MemberOf
    p_IsOnTheFly = That.IsOnTheFly
    p_OnTheFlyLength = That.OnTheFlyLength
    
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

''<summary>Represents the method's physical address.</summary>
''<retval>Returns method's physical reference as a 4-bytes pointer.</retval>
Public Property Get Reference() As Long
    Reference = p_TargetFunction
End Property
''<summary>Represents the method's name determined in class constructor.</summary>
''<retval>Returns method's name, this can be null.</retval>
Public Property Get Name() As String
    Name = p_Name
End Property

Public Property Get Description() As String
    If mString.IsEmptyOrWhiteSpace(p_Description) Then
        Description = Name
    Else
        Description = p_Description
    End If
End Property
Public Property Let Description(ByVal Value As String)
    p_Description = Value
End Property

Public Property Get IsMember() As Boolean
    IsMember = p_IsMember
End Property
Friend Property Get MemberOf() As Object
    Set MemberOf = p_MemberOf
End Property

''<summary>Indicates that the method can be executed or not.</summary>
''<retval>Returns a boolean that Indicates the method can be executed or not.</retval>
Public Property Get IsExecutable() As Boolean
    IsExecutable = (p_TargetFunction <> 0)
    If IsExecutable Then _
        IsExecutable = (API_IsBadCodePtr(p_TargetFunction) = NO_VALUE)
    If IsExecutable Then _
        IsExecutable = Memory.CanAccessTo(p_TargetFunction, 0, mapExecute)
    p_IsExecutable = IsExecutable
End Property

Public Property Get Prototype() As MethodPrototype
    Set Prototype = p_Prototype
End Property

Public Function Invoke(Optional ByVal Args As ArgumentList) As Variant
Attribute Invoke.VB_UserMemId = 0
THROW_NotExecutable
    Dim Attrs As MINT_INVOKERS_ATTRIBUTES, Obj As Object
    If p_IsMember Then
        Set Obj = p_MemberOf
        Attrs = MIA_PUSHOBJECT
'    ElseIf IsInIDE Then
'        throw Exps.InvalidOperationException("Unsafe call does not allowed in design mode.")
    End If
    
    Call MoveVariantX(Invoke, _
        mint_invokers_invoke(Obj, Me, p_Prototype, Args, Attrs))
End Function

Public Function BeginInvoke(Optional ByVal Args As ArgumentList, Optional ByVal Callback As AsyncCallback, Optional ByVal Object As Object) As IAsyncResult
Async
THROW_NotExecutable
    
End Function
Public Function EndInvoke(ByVal AR As IAsyncResult) As IAsyncResult
Async
THROW_NotExecutable
    
End Function

Public Function ProfileInvoke(Optional outRetVal, Optional ByVal Args As ArgumentList) As ProfilerInfo
THROW_NotExecutable
    Set ProfileInvoke = New ProfilerInfo
    Call ProfileInvoke.StartProfile
    
    If IsMissing(outRetVal) Then
        Call Invoke(Args)
    ElseIf DBNull.IsDBNull(outRetVal) Then
        Call Invoke(Args)
    Else
        Call MoveVariant(outRetVal, Invoke(Args))
    End If
    
    Call ProfileInvoke.EndProfile
End Function

Friend Property Get IsOnTheFly() As Boolean
THROW_NotExecutable
    IsOnTheFly = p_IsOnTheFly
End Property
Friend Property Let IsOnTheFly(ByVal Value As Boolean)
THROW_NotExecutable
    p_IsOnTheFly = Value
End Property

Friend Property Get OnTheFlyLength() As Long
THROW_NotExecutable
    OnTheFlyLength = p_OnTheFlyLength
End Property
Friend Property Let OnTheFlyLength(ByVal Value As Long)
THROW_NotExecutable
    p_OnTheFlyLength = Value
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
Public Function Clone() As Method
    Set Clone = New Method
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
End Sub

Private Sub THROW_NotExecutable()
    If p_IsExecutable Then Exit Sub
    If Not IsExecutable Then _
        throw Exps.InvalidCallException("Target method is not executable.")
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
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo)
    Call Constructor(CInfo)
End Sub
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function
Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function
Private Function IObject_MetaObject() As MetaObject
    Set IObject_MetaObject = MetaObject
End Function
Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
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
'<section IObject Implementation>
'
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
