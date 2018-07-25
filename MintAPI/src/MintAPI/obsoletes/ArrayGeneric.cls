VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayGeneric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ArrayGeneric" tags="">
''  <summary>
''    Generic array representation class.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using></using>
''  <idea>Me , .net Framework , Qt</idea>
''  <sample>
''      Dim Dims(3) As Long
''      Dims(0) = 10
''      Dims(1) = 2
''      Dims(2) = 5
''      stdout Arrays.Create(atByteArray, Dims).Length 'Results in an array with 100 items.
''  </sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
Implements ICloneable
'Implements IEnumerable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ArrayGeneric" '#DO NOT REMOVE IT
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
Public Enum ArrayTypes
    atLongArray = vbLong Or vbArray
    atIntegerArray = vbInteger Or vbArray
    atByteArray = vbByte Or vbArray
    atSingleArray = vbSingle Or vbArray
    atDoubleArray = vbDouble Or vbArray
    atStringArray = vbString Or vbArray
    atObjectArray = VBObject Or vbArray
    atBooleanArray = vbBoolean Or vbArray
    atDateArray = vbDate Or vbArray
    atCurrencyArray = vbCurrency Or vbArray
    atVariantArray = vbVariant Or vbArray
    atUserDefinedTypeArray = VariantTypes.VT_RECORD Or vbArray
    atUDT = atUserDefinedTypeArray
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

Dim refVariant As Variant
Dim refSafeArray As SafeArray
Dim refSAPtr As Long
Dim ln As Long
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
Friend Sub ConstructorArg(ByVal Arguments As ArgumentList)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal ArrayType As ArrayTypes, ByVal Length As Long)
    refSAPtr = API_SafeArrayCreateVector((ArrayType Xor VT_ARRAY), 0, Length)
    VariantType(refVariant) = ArrayType Or vbArray
    VariantDataPtr(refVariant) = refSAPtr
    Call API_CopyMemory(ByVal VarPtr(refSafeArray), ByVal VarPtr(refSAPtr), 4)
    ln = Length
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal ArrayType As ArrayTypes, Dimensions() As Long, Optional ByVal LowerBounds As Long = 0)
    Dim Dims As Long, i As Long, l_Bound As Long
    Dims = ArraySize(Dimensions)
    If Dims <= 0 Then throw Exps.NegativeArgumentException("Dementions must be positive non zero number.")

    ln = 1
    l_Bound = LBound(Dimensions)
    For i = l_Bound To UBound(Dimensions)
        ln = ln * Dimensions(i)
    Next
    If ln = 0 Then _
        throw Exps.InvalidArgumentValueException(Mtr("One (or more) of the dimentions declared as zero."))

    Dim DimenHndl As Long, Length As Long
    Length = Dims * SIZEOF_SAFEARRAYBOUND
    DimenHndl = API_CoTaskMemAlloc(Length)
    For i = 0 To Dims - 1
        MemLong(DimenHndl + (i * SIZEOF_SAFEARRAYBOUND)) = Dimensions(l_Bound + i)
        MemLong(DimenHndl + (i * SIZEOF_SAFEARRAYBOUND) + VLEN_LONG) = LowerBounds
    Next
    
    VariantType(refVariant) = (ArrayType Or vbArray)
    refSAPtr = API_SafeArrayCreateN((ArrayType Xor VT_ARRAY), Dims, ByVal DimenHndl)
    VariantDataPtr(refVariant) = refSAPtr
    Call memcpy(ByVal VarPtr(refSafeArray), ByVal refSAPtr, VLEN_PTR)
    
    Call API_CoTaskMemFree(DimenHndl)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal AG As ArrayGeneric)
    
    
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

Public Property Get Handle() As Long
    Handle = refSAPtr
End Property
Public Property Get DataHandle() As Long
    DataHandle = refSafeArray.pvData
End Property

Public Sub Dispose()
    
End Sub
Public Sub Clear()
    
End Sub
Public Property Get Length() As Long
    Length = ln
End Property
Public Property Get LongLength() As Int64
    LongLength.LowerValue = ln
    LongLength.HigherValue = 0
End Property
Public Function GetLength(ByVal Dimension As Long) As Long
    Dim Dims As Long
    Dims = API_SafeArrayGetDim(refSAPtr)
    
    If Dimension > Dims Then _
        throw Exps.OutOfRangeException
    
    Dim DimsTablePtr As Long, retVal As Long
    DimsTablePtr = refSAPtr + SIZEOF_SAFEARRAY
    
    If Dimension <= 0 Then
        retVal = 1
        Dim i As Long, IsOnceExecuted As Boolean
        For i = 0 To Dims - 1
            retVal = retVal * MemLong(DimsTablePtr + (SIZEOF_SAFEARRAYBOUND * i))
            IsOnceExecuted = True
        Next i
        If Not IsOnceExecuted Then Exit Function
    Else
        DimsTablePtr = DimsTablePtr + ((Dims - Dimension) * SIZEOF_SAFEARRAYBOUND)
        retVal = MemLong(DimsTablePtr)
    End If
    GetLength = retVal
End Function
Public Property Get Rank() As Long
    
End Property
Public Property Get IsReadOnly() As Boolean
    
End Property

''<summary>
''  Bounds a secendary array to the source array<br>
''  When one of the arrays change other one take effect too.<br>
''  Aslo when one of the arrays resize other one resize too.
''</summary>
''<sample>
''  <pre>
''      Dim b() As Byte
''      Call Arrays.DefineVector(atByteArray, 1000).BoundArray(b)
''      b(999) = 15
''      Call Debugger.Show(b, dbgsModal)
''  </pre>
''</sample>
Public Function BoundArray(ByRef arr As Variant) As Long
    
End Function

Public Function CopyMemory(ByVal Source As Long, Optional ByVal SourceStartAt As Long = 0, Optional ByVal SourceLength As Long = -1, Optional ByVal DestinationStartAt As Long = 0, Optional ByVal DestinationLength As Long = -1, Optional FillWhit As Variant) As Boolean
    
End Function

Public Function DataMemory() As Memory
    
End Function

Public Function IsObjectArray() As Boolean
    
End Function
Public Function IsArrayArray() As Boolean
    
End Function
Public Property Get ArrayType() As ArrayTypes
    
End Property

Public Function IsObject(ByVal Index As Long, Optional Dimension As Long = 0) As Boolean
    
End Function
Public Function IsArray(ByVal Index As Long, Optional Dimension As Long = 0) As Boolean
    
End Function

Public Property Get LinearItem(ByVal Index As Long)
    
End Property
Public Property Let LinearItem(ByVal Index As Long, Value As Variant)
    
End Property
Public Property Set LinearItem(ByVal Index As Long, Value As Object)
    
End Property

Public Property Get Item(ByVal Index As Long, Optional Dimension As Long = 0)
    
End Property
Public Property Let Item(ByVal Index As Long, Optional Dimension As Long = 0, Value As Variant)
    
End Property
Public Property Set Item(ByVal Index As Long, Optional Dimension As Long = 0, Value As Object)
    
End Property
Public Property Get ItemPointer(ByVal Index As Long, Optional Dimension As Long = 0) As Memory
    
End Property
Public Property Let ItemPointer(ByVal Index As Long, Optional Dimension As Long = 0, Value As Memory)
    
End Property
Public Property Set ItemPointer(ByVal Index As Long, Optional Dimension As Long = 0, Value As Memory)
    
End Property

Public Function IndexOf() As Long
    
End Function
Public Function LastIndexOf() As Long
    
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
Public Property Get Alias() As String
    Alias = CLASSALIASSTRING
End Property
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
    Set MetaObject = InitMetaObject
End Function

Public Function Clone() As Object
    
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
    
End Sub
Private Sub Class_Terminate()
    Signal.Emit Me, SIGNAL_TERMINATING
    If Not myMetaObject Is Nothing Then
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
Private Sub IObject_Initialize(ByVal Arguments As ArgumentList)
    Call ConstructorArg(Arguments)
End Sub
Private Property Get IObject_Alias() As String
    IObject_Alias = CLASSALIASSTRING
End Property
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
Private Function IObject_Notify(ByVal Name As String, ByVal Arguments As ArgumentList) As Variant
    Select Case LCase$(Name)
        'Case ""
            
        Case Else
            throw Exps.UnknownNotificationException, Me
    End Select
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


'=============================================
'=============================================
'=============================================
'<section MetaObject constructor>
'
Private Function InitMetaObject() As MetaObject
    If Not myMetaObject Is Nothing Then
        Set InitMetaObject = myMetaObject
    End If
    Call Objects.MetaObject.Setup(MetaObject, Me, True, False, CLASSVERSIONSTRING)
    '...
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
