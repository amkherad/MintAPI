VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MethodPrototype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MethodPrototype" tags="">
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
Implements Freezable
Implements ICloneable
'Implements IDisposable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "MethodPrototype" '#DO NOT REMOVE IT
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

Private p_ParametersScheme As ParametersScheme
Private p_CallingConvension As CallingConvensions
Private p_ReturnValueType As VariantTypes
Private p_IsFreezed As Boolean
Private p_IsUnspecified As Boolean
Private p_AllowUnsafeCall As Boolean
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
Friend Sub Constructor0(ByVal Factory As MethodPrototypeFactory)
    Set p_ParametersScheme = Factory.ParametersScheme
    p_CallingConvension = Factory.CallingConvension
    p_ReturnValueType = Factory.ReturnValueType
    p_IsUnspecified = False
    p_AllowUnsafeCall = Factory.AllowUnsafeCall
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal IsUnspecified As Boolean, ByVal AllowUnsafeCall As Boolean)
    Set p_ParametersScheme = New ParametersScheme
    p_CallingConvension = ccStdCall
    p_ReturnValueType = VT_VOID
    p_IsUnspecified = IsUnspecified
    p_AllowUnsafeCall = AllowUnsafeCall
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As MethodPrototype)
    Set p_ParametersScheme = That.ParametersScheme
    p_CallingConvension = That.CallingConvension
    p_ReturnValueType = That.ReturnValueType
    p_IsUnspecified = That.IsUnspecified
    p_AllowUnsafeCall = That.AllowUnsafeCall
    
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

Public Property Get ParametersScheme() As ParametersScheme
    Set ParametersScheme = p_ParametersScheme
End Property

Public Property Get CallingConvension() As CallingConvensions
    CallingConvension = p_CallingConvension
End Property

Public Property Get ReturnValueType() As VariantTypes
    ReturnValueType = p_ReturnValueType
End Property

Public Property Get AllowUnsafeCall() As Boolean
    AllowUnsafeCall = p_AllowUnsafeCall
End Property

Friend Property Get IsUnspecified() As Boolean
    IsUnspecified = p_IsUnspecified
End Property
Friend Property Let IsUnspecified(ByVal Value As Boolean)
    p_IsUnspecified = Value
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

Public Function Clone() As MethodPrototype
    Set Clone = New MethodPrototype
    Call Clone.ConstructorCopy(Me)
End Function

'Public Sub Dispose()
'
'End Sub
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
'<section Freezable Implementations>
'
Private Function Freezable_Freeze() As Boolean
    If p_IsFreezed Then Exit Function
    p_IsFreezed = True
    Freezable_Freeze = True
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
'<section IDisposable Implementations>
'
'Private Sub IDisposable_Dispose(): Call Dispose: End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
