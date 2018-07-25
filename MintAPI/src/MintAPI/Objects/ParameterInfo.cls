VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParameterInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Member1" ,"Module"
'@PROJECT_LICENSE

''<class name="ParameterInfo" tags="immutable">
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
Const CLASSID As String = "ParameterInfo" '#DO NOT REMOVE IT
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
Public Enum ParameterAttributes
    paNone = 0
    'paByVal = 1
    paByRef = 1
    paRetVal = 2
    paLCID = 4
    paOptional = 16
    paParams = 32
    paHasDefault = 4096
    paHasFieldMarshal = 8192
    paReserved3 = 16384
    paReserved4 = 32768
    paReservedMask = 61440
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

Dim p_Name As String
Dim p_Type As VariantTypes
Dim p_Attributes As ParameterAttributes
Dim p_DefaultValue As Variant
Dim p_Position As Long
Dim p_Description As String
Dim p_Module As Module
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
Friend Sub Constructor0(ByVal Name As String, ByVal ParamType As VariantTypes, _
    ByVal Attributes As ParameterAttributes, ByRef DefaultValue As Variant, _
    ByVal Position As Long)
    p_Name = Name
    p_Type = ParamType
    p_Attributes = Attributes
    p_DefaultValue = DefaultValue
    p_Position = Position
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal TypeInfo As ITypeInfo, ByRef TAttr As API_TYPEATTR)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As ParameterInfo)
    p_Name = That.Name
    p_Type = That.ParameterType
    p_Attributes = That.Attributes
    p_DefaultValue = That.DefaultValue
    p_Position = That.Position
    Set p_Module = That.Module
    p_Description = That.Description
    
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
THROW_IfNull
    Name = p_Name
End Property

Public Property Get ParameterType() As VariantTypes
THROW_IfNull
    ParameterType = p_Type
End Property

Public Property Get Attributes() As ParameterAttributes
THROW_IfNull
    Attributes = p_Attributes
End Property

Public Property Get IsByRef() As Boolean
THROW_IfNull
    IsByRef = ((p_Attributes And paByRef) = paByRef)
End Property
Public Property Get IsByVal() As Boolean
    IsByVal = Not IsByRef
End Property

Public Property Get IsOptional() As Boolean
THROW_IfNull
    IsOptional = ((p_Attributes And paOptional) = paOptional)
End Property

Public Property Get IsRetVal() As Boolean
THROW_IfNull
    IsRetVal = ((p_Attributes And paRetVal) = paRetVal)
End Property

Public Property Get IsParams() As Boolean
THROW_IfNull
    IsParams = ((p_Attributes And paParams) = paParams)
End Property

Public Property Get Position() As Long
THROW_IfNull
    Position = p_Position
End Property

Public Property Get Size() As Long
THROW_IfNull
    Size = Info.GetArgumentLength(p_Type)
End Property

Public Property Get OnStackSize() As Long
THROW_IfNull
    OnStackSize = Info.GetArgumentLength(p_Type)
    If OnStackSize < VLEN_PTR Then OnStackSize = VLEN_PTR
End Property


Public Property Get DefaultValue() As Variant
THROW_IfNull
    Call EvaluateX(DefaultValue, p_DefaultValue)
End Property

Public Property Get Module() As Module
THROW_IfNull
    If p_Module Is Nothing Then
        Set Module = Instance.Module
    Else
        Set Module = p_Module
    End If
End Property

Public Property Get Description() As String
THROW_IfNull
    Description = p_Description
End Property

Public Function IsMatch(ByRef Val As Variant) As Boolean
    IsMatch = (p_Attributes = VariantType(Val))
End Function

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

Public Function Clone() As ParameterInfo
    Set Clone = New ParameterInfo
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
End Sub

Private Sub THROW_IfNull()
    If Not This_Constructed Then _
        throw Exps.ClassNotInitializedException
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
