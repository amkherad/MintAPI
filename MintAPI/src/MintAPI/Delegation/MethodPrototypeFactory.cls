VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MethodPrototypeFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MethodPrototypeFactory" static subclass tags="factory">
''  <summary>
''    Some class description.
''  </summary>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'Implements IObject
Implements ICloneable
'Implements IDisposable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "MethodPrototypeFactory" '#DO NOT REMOVE IT
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
'Dim This_Constructed As Boolean
'Dim myMetaObject As MetaObject

Private p_ParametersScheme As New ParametersScheme
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
'Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
'
'    This_Constructed = True
'End Sub
'Friend Sub Constructor0(ByVal ID As Long)
'
'    This_Constructed = True
'End Sub
'Friend Sub ConstructorCopy(ByVal That As {My Object Type})
'
'    This_Constructed = True
'End Sub
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

Public Function NewArg(ByVal Name As String, ByVal ParameterType As VariantTypes, Optional ByVal ParameterAttributes As ParameterAttributes) As MethodPrototypeFactory
    Dim Param As New ParameterInfo
    Dim ParamType As VariantTypes, ParamAttrs As ParameterAttributes
    
    If (ParameterAttributes And paRetVal) = paRetVal Then throw Exps.InvalidOperationException
    
    ParamType = ParameterType
    ParamAttrs = ParameterAttributes
    
    If (ParamType And VT_BYREF) = VT_BYREF Then
        ParamAttrs = ParamAttrs Or paByRef
        ParamType = ParamType Xor VT_BYREF
    End If
    
    Dim Position As Long
    Position = p_ParametersScheme.OnStackSize
    
    Call Param.Constructor0(Name, ParamType, ParamAttrs, vbNullPtr, Position)
    Call p_ParametersScheme.Append(Param)
    Set NewArg = Me
End Function
Public Function SetRetVal(ByVal VarType As VariantTypes) As MethodPrototypeFactory
    p_ReturnValueType = VarType
    
    Set SetRetVal = Me
End Function
Public Function SetCallingConvension(ByVal CC As CallingConvensions) As MethodPrototypeFactory
    p_CallingConvension = CC
    
    Set SetCallingConvension = Me
End Function

Public Function AllowUnsafeUsage(Optional ByVal AbsoluteValue As Boolean = True) As MethodPrototypeFactory
    p_AllowUnsafeCall = AbsoluteValue
    
    Set AllowUnsafeUsage = Me
End Function


'Friend Property Get IsUnspecified() As Boolean
'    IsUnspecified = p_IsUnspecified
'End Property
Friend Property Get ParametersScheme() As ParametersScheme
    Set ParametersScheme = p_ParametersScheme
End Property
Friend Property Get CallingConvension() As CallingConvensions
    CallingConvension = p_CallingConvension
End Property
Friend Property Get ReturnValueType() As VariantTypes
    ReturnValueType = p_ReturnValueType
End Property
Friend Property Get AllowUnsafeCall() As Boolean
    AllowUnsafeCall = p_AllowUnsafeCall
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
'Public Function Equals(Value As Variant) As Boolean
'    Equals = Objects.Equals(Me, Value)
'End Function
'Public Function ToString() As String
'    ToString = CLASSALIAS
'End Function
'Public Function GetHashCode() As Long
'    GetHashCode = Objects.GetHashCode(Me)
'End Function
'Public Function MetaObject() As MetaObject
'    If myMetaObject Is Nothing Then
'        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
'    End If
'    Set MetaObject = myMetaObject
'End Function

Public Function Clone() As MethodPrototype
Attribute Clone.VB_UserMemId = 0
    Set Clone = New MethodPrototype
    Call Clone.Constructor0(Me)
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
Private Sub Class_Initialize()
    p_CallingConvension = ccStdCall
End Sub
'Private Sub Class_Terminate()
'    If Not myMetaObject Is Nothing Then
'        Signal.Emit Me, SIGNAL_TERMINATING
'        Call myMetaObject.Dispose
'        Set myMetaObject = Nothing
'    End If
'End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'

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
