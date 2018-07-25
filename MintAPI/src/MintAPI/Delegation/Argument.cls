VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Argument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="Argument" tags="">
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
Const CLASSID As String = "Argument" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"

Const PROPERTY_NAME As String = "argname"
Const PROPERTY_VALUE As String = "argvalue"
Const PROPERTY_ISBYREF As String = "argbr"
Const PROPERTY_TYPE As String = "argtp"
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
Dim p_Value As Variant
Dim p_CloneArray As Boolean
Dim p_IsByVal As Boolean
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
Friend Sub Constructor0(ByVal Name As String, ByRef Value As Variant) 'ByRef
    p_Name = Name
    Call memcpy(p_Value, Value, VLEN_VARIANT)
    If VBA.IsObject(p_Value) Then Call IUnknown.AddRef(p_Value)
    p_IsByVal = False
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal Name As String, ByRef Value As Variant) 'ByVal
    p_Name = Name
    Call Evaluate(p_Value, Value)
    p_IsByVal = True
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal Argument As Argument)
    
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

Public Property Get Name() As String
    Name = p_Name
End Property
Public Property Let Name(ByVal Value As String)
    p_Name = Value
End Property

Public Property Get Value() As Variant
    Call EvaluateX(Value, p_Value)
End Property
Public Property Let Value(ByRef Value As Variant)
    Call SetValue(Value)
End Property
Public Property Set Value(ByVal Value As Object)
    Call SetValue(Value)
End Property
Public Sub SetValue(ByRef Value As Variant)
    Call Evaluate(p_Value, Value)
End Sub

Public Property Get ValueInd() As Variant
    If IsRecord Then
        Call EvaluateX(ValueInd, p_Value)
    Else
        Call API_VariantCopyInd(ValueInd, p_Value)
    End If
End Property

Public Property Get AbsoluteSize() As Long
    Dim VT As VariantTypes
    VT = ArgumentType
    If (VT And VT_BYREF) = VT_BYREF Then
        AbsoluteSize = VLEN_PTR
    ElseIf (VT And VT_BSTR) = VT_BSTR Then
        AbsoluteSize = VLEN_PTR
    ElseIf (VT And VT_DISPATCH) = VT_DISPATCH Then
        AbsoluteSize = VLEN_PTR
    ElseIf (VT And VT_ARRAY) = VT_ARRAY Then
        AbsoluteSize = VLEN_PTR
    ElseIf (VT And VT_UDT) = VT_UDT Then
        AbsoluteSize = VLEN_PTR
    Else
        AbsoluteSize = Info.GetArgumentLength(VT)
    End If
End Property

Public Property Get Reference() As Long
    Reference = (VarPtr(p_Value) + VARIANT_DATA_OFFSET)
End Property

Public Property Get AbsoluteReference() As Long
    AbsoluteReference = (VarPtr(p_Value) + VARIANT_DATA_OFFSET)
    If IsByRef Then Call memcpy(AbsoluteReference, ByVal AbsoluteReference, VLEN_PTR)
End Property


Public Property Get ArgumentType() As VariantTypes
    Call memcpy(ArgumentType, p_Value, VLEN_VARTYPE)
End Property
Public Property Get IsByRef() As Boolean
    IsByRef = (ArgumentType And VT_BYREF) = VT_BYREF
End Property
Public Property Get IsByVal() As Boolean
    IsByVal = (ArgumentType And VT_BYREF) <> VT_BYREF
End Property
Public Property Get IsAccessible() As Boolean
    IsAccessible = True
End Property
Public Property Get IsObject() As Boolean
    IsObject = (ArgumentType And VT_DISPATCH) = VT_DISPATCH
End Property
Public Property Get IsArray() As Boolean
    IsArray = (ArgumentType And VT_ARRAY) = VT_ARRAY
End Property
Public Property Get IsString() As Boolean
    IsString = (ArgumentType And VT_BSTR) = VT_BSTR
End Property
Public Property Get IsRecord() As Boolean
    IsRecord = ((ArgumentType And VT_UDT) = VT_UDT)
End Property

Public Property Get IsAlive() As Boolean
    IsAlive = True
End Property

Public Sub CopyTo(ByRef What As Variant)
    Call API_VariantCopy(What, p_Value)
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
    On Error GoTo CatchErr
    ToString = ByteArray(p_Value, False).ToString
CatchErr:
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

Public Function Clone() As Argument
    Set Clone = New Argument
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
    p_IsByVal = True
End Sub
Private Sub Class_Terminate()
    If p_IsByVal Then
        Call API_VariantClear(p_Value)
    Else
        If VBA.IsObject(p_Value) Then Call IUnknown.Release(p_Value)
        Call memzero(p_Value, VLEN_VARIANT)
    End If
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
'<section ICloneable Implementations>
'
Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
