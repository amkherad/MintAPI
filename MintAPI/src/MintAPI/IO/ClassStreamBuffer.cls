VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassStreamBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Member1" ,"Buffer"
'@PROJECT_LICENSE

''<class name="ClassStreamBuffer" tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Provides stream data for class stream engine.
''      </constructor>
''  </constructors>
''  <using>IClassStream</using>
''  <idea>Me</idea>
''  <sample>
''      out Console, Endl 'Endl passes a ClassStreamBuffer that containes new line.
''  </sample>
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
Const CLASSID As String = "ClassStreamBuffer" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"

Const PROPERTY_NAME As String = "Name"
Const PROPERTY_BUFFERTYPE As String = "BufferType"
Const PROPERTY_BUFFER As String = "Buffer"
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Public Enum ClassStreamBufferType
    'Basic data's
    csbtNull = &H0 'NOTE: totally null.
    csbtData = &H0
    
    csbtNewLine = &H1
    csbtFlush = &H2
    csbtClear = &H3
    
    csbtTypeMask = &HFF
    '-----------------
    'Complicated data's
    csbtSeek = &H100 Or csbtData
    '-----------------
    'Data types
    csbtDirectionUnspecified = 0
    csbtOutput = &H10000 Or csbtData
    csbtInput = &H20000 Or csbtData
    csbtDirectionMask = &H30000
    
    csbtBinary = &H100 Or csbtData
    'csbtUnknown = &H300 Or csbtData
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
Dim p_BType As ClassStreamBufferType
Dim p_Flags As Long

Dim CtoB As Boolean 'Convert to Binary.
Dim p_Bin As ByteArray
Dim p_BRPtr As Long 'ByRef Ptr
Dim p_VType As Long
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
Friend Sub Constructor0(ByVal VBufferType As ClassStreamBufferType, Optional ByRef Value As Variant, Optional ByVal Length As Long = -1)
    If (VBufferType And csbtInput) = csbtInput Then
        If (VBufferType And csbtOutput) = csbtOutput Then _
            throw Exps.InvalidOperationException
        If IsMissing(Value) Then _
            throw Exps.OptionalArgumentNotPassed
        If IsEmpty(Value) Or IsNull(Value) Then _
            throw Exps.ArgumentNullException
        p_VType = VariantType(Value)
        If (p_VType And VT_BYREF) = VT_BYREF Then
            p_BRPtr = VarPtr(Value)
            
            'Exclude the VT_BYREF value from p_VType to get the absolute var type.
            p_VType = p_VType Xor VT_BYREF
        Else
            p_BRPtr = VarPtr(Value)
        End If
        CtoB = True
    ElseIf (VBufferType And csbtTypeMask) = csbtData Then
        If (VBufferType And csbtBinary) = csbtBinary Then
            Set p_Bin = ByteArray(Value, True, Length)
        Else
            Set p_Bin = ByteArray(Value, False, Length)
        End If
    Else
        Set p_Bin = Nothing
    End If
    p_BType = VBufferType
        
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByRef Value As Variant, Optional ByVal Length As Long = -1)
    Set p_Bin = ByteArray(Value, False, Length)
    p_BType = csbtData Or csbtOutput
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal Class As ClassStreamBuffer)
    Me.BufferType = Class.BufferType
    Me.Name = Class.Name
    Call Me.SetBuffer(Class.Buffer)
    Me.Flags = Class.Flags
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

Public Property Get Flags() As Long
    Flags = p_Flags
End Property
Public Property Let Flags(ByVal Value As Long)
    p_Flags = Value
End Property

Public Property Get IsInput() As Boolean
    IsInput = (p_BType And csbtInput) = csbtInput
End Property
Public Property Get IsOutput() As Boolean
    IsOutput = (p_BType And csbtOutput) = csbtOutput
End Property

Public Property Get InputVarType() As VariantTypes
    InputVarType = p_VType
End Property
Public Property Get InputVarLen() As Long
    InputVarLen = Info.GetArgumentLength(p_VType)
End Property

Public Property Get BufferType() As ClassStreamBufferType
    BufferType = p_BType
End Property
Public Property Let BufferType(ByVal Value As ClassStreamBufferType)
    Select Case Value
        'Case csbtInput, csbtOutput, csbtBinary, csbtBinary Or csbtInput, csbtBinary Or csbtOutput
        Case csbtNull, csbtNewLine, csbtFlush, csbtClear, csbtSeek
            p_BRPtr = vbNullPtr
            p_VType = vbNullPtr
            Set p_Bin = Nothing
        Case Else
            throw Exps.InvalidOperationException
    End Select
    p_BType = Value
End Property

Public Property Get ConvertToBinary() As Boolean
    ConvertToBinary = CtoB
End Property
Public Property Let ConvertToBinary(ByVal Value As Boolean)
    If IsInput Then throw Exps.InvalidOperationException
    CtoB = Value
End Property

Public Sub SetBuffer(ByRef Value As Variant, Optional ByVal Length As Long = -1)
    Set p_Bin = ByteArray(Value, CtoB, Length)
End Sub
Public Sub MoveToInput(Optional ByVal DestroySource As Boolean = True)
    If Not IsInput Then throw Exps.InvalidOperationException
    If p_VType = vbNullPtr Then throw Exps.InvalidOperationException
    If p_BRPtr = vbNullPtr Then throw Exps.InvalidOperationException
    If p_Bin Is Nothing Then throw Exps.InvalidOperationException
    If p_Bin.Length <= 0 Then throw Exps.InvalidOperationException
    Dim ALen As Long, TBAPtr As Long, DataLen As Long
    TBAPtr = p_Bin.DataHandle
    DataLen = p_Bin.Length
    If TBAPtr = vbNullPtr Then throw Exps.InvalidOperationException
    Select Case p_VType
        Case VT_BSTR, VT_LPSTR, VT_LPWSTR
            Call memcpy(ByVal p_BRPtr, ByVal API_SysAllocStringLen(TBAPtr, DataLen), VLEN_STRPTR)
        Case VT_I4, VT_DISPATCH, VT_ARRAY, VT_UI4, VT_INT, VT_UINT, VT_INT_PTR, VT_UINT_PTR, VT_HRESULT, VT_UserDefined
            If DataLen < 4 Then throw Exps.InvalidOperationException
            Call memcpy(ByVal p_BRPtr, ByVal TBAPtr, VLEN_PTR)
        Case VT_I2, VT_BOOL, VT_UI2
            If DataLen < 2 Then throw Exps.InvalidOperationException
            Call memcpy(ByVal p_BRPtr, ByVal TBAPtr, 2)
        Case VT_I8, VT_DATE, VT_CY, VT_CF
            If DataLen < 8 Then throw Exps.InvalidOperationException
            Call memcpy(ByVal p_BRPtr, ByVal TBAPtr, VLEN_DOUBLE)
        Case VT_I1, VT_UI1
            Call memcpy(ByVal p_BRPtr, ByVal TBAPtr, 1)
        Case VLEN_VARIANT ', VT_
            If DataLen < 16 Then throw Exps.InvalidOperationException
            Call memcpy(ByVal MemPtr(p_BRPtr + VARIANT_DATA_OFFSET), ByVal TBAPtr, VLEN_VARIANT)
        Case VT_DECIMAL ', VT_
            If DataLen < 16 Then throw Exps.InvalidOperationException
            Call memcpy(ByVal MemPtr(p_BRPtr + VARIANT_DATA_OFFSET), ByVal TBAPtr, VLEN_DECIMAL)
        Case Else
            throw Exps.InvalidOperationException
    End Select
    
    If DestroySource Then _
        Call p_Bin.Dispose
End Sub

Public Property Get Buffer() As ByteArray
    Select Case p_BType
        Case csbtNull
            Set Buffer = ByteArray
        Case csbtNewLine
            Set Buffer = ByteArray(vbNewLine)
        Case csbtFlush
            Set Buffer = ByteArray
        Case csbtClear
            Set Buffer = ByteArray(Chr$(18))
        Case Else
            If (p_BType And csbtData) = csbtData Then
                Set Buffer = p_Bin
            Else
                Set Buffer = ByteArray
            End If
    End Select
End Property

Public Property Get ReadLength() As Long
    Select Case p_VType
        Case VT_BSTR, VT_BSTR, VT_BSTR
            ReadLength = -1
        Case Else
            If (p_VType And VT_ARRAY) = VT_ARRAY Then
                ReadLength = -1
            Else
                ReadLength = Info.GetArgumentLength(p_VType)
            End If
    End Select
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
    Dim BufferTypeStr As String
    Select Case p_BType
        Case csbtNull
            BufferTypeStr = "Null"
        Case csbtNewLine
            BufferTypeStr = "New Line"
        Case csbtFlush
            BufferTypeStr = "Flush"
        Case csbtClear
            BufferTypeStr = "Clear"
        Case csbtOutput
            BufferTypeStr = "Output"
        Case csbtInput
            BufferTypeStr = "Input"
        Case csbtBinary
            BufferTypeStr = "<Binary>" & " Length:" & CStr(Buffer.Length)
        Case csbtBinary Or csbtInput
            BufferTypeStr = "<Binary Input>" & " Length:" & CStr(Buffer.Length)
        Case csbtBinary Or csbtOutput
            BufferTypeStr = "<Binary Output>" & " Length:" & CStr(Buffer.Length)
        Case Else
            If p_BType And csbtInput = csbtInput Then
                BufferTypeStr = "<Unknown Input>" & " Length:" & CStr(Buffer.Length)
            ElseIf p_BType And csbtOutput = csbtOutput Then
                BufferTypeStr = "<Unknown Output>" & " Length:" & CStr(Buffer.Length)
            Else
                BufferTypeStr = "<Unknown>" & " Length:" & CStr(Buffer.Length)
            End If
    End Select
    
    ToString = CLASSALIASSTRING & IIf(p_Name = "", "", "(" & p_Name & ")") & " Value Type: " & BufferTypeStr
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

Public Function Clone() As Object
    Dim CSB As New ClassStreamBuffer
    Call CSB.ConstructorCopy(CSB)
    Set Clone = CSB
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
Private Sub Class_InitProperties()
    BufferType = csbtNull
    CtoB = True
End Sub

Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub


Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Name = .ReadProperty(PROPERTY_NAME, "")
        BufferType = .ReadProperty(PROPERTY_BUFFERTYPE, ClassStreamBufferType.csbtNull)
        Call SetBuffer(.ReadProperty(PROPERTY_BUFFER, Nothing))
    End With
End Sub
Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROPERTY_NAME, Name)
        Call .WriteProperty(PROPERTY_BUFFERTYPE, BufferType)
        Call .WriteProperty(PROPERTY_BUFFER, Buffer)
    End With
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
'<section ICloneable Implementations>
'
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
