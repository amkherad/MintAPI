VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BigNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''@PROJECT_LICENSE
'
'
''Ideas from Me , [Out College Programming Confrence] , [VBCorLib integeral section]
'
'Option Base 0
'Option Explicit
'Const CLASSID As String = "BigNumber"
'Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
'Const CLASSVERSIONSTRING As String = "0.1"
'Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING
'
'Implements ICloneable
''Implements IAliasable
'
'
'Const EXTRA_BUFFER As Long = 5
'Const NEGATIVE_SIGN As Long = 45
'Const POSITIVE_SIGN As Long = 43
'Const POINT_SIGN As Long = 46
'Const SCIENTIFIC_Ec As Long = 69
'Const SCIENTIFIC_E As Long = 101
''Const SCIENTIFIC_Fc As Long = 70
''Const SCIENTIFIC_F As Long = 102
'
''Numbers save from 0(NULL) to 9(...)
'
'Private Enum BIGNUM_SIGN
'    sNegative = -1
'    sZero = 0
'    sPositive = 1
'End Enum
'
''Save number as scientific.
'Private Type BIGNUM_STRUCT
'    Sign As BIGNUM_SIGN
'
'    Float As ByteArray
'    Integer As ByteArray
'    base_N As Long
'    SE As Long 'the Only Limited Value.
'
'    Floatable As Boolean
'    show_pos_sign As Boolean
'End Type
'
'Private n As BIGNUM_STRUCT
'
'Friend Function GetStruct() As BIGNUM_STRUCT
'    GetStruct = n
'End Function
'
'Friend Sub InitializeW(bnStruct As BIGNUM_STRUCT)
'    n = bnStruct
'    'If n.B Is Nothing Then Set n.B = New ByteArray
'End Sub
'Friend Sub Initialize(Optional InitialValue)
'    If Not IsMissing(InitialValue) Then Call SetLocalValue(InitialValue)
'    'If n.B Is Nothing Then Set n.B = New ByteArray
'End Sub
'Private Sub Class_Initialize()
'    n.Floatable = True
'    n.show_pos_sign = True
'End Sub
'
'Friend Sub internal_SetBytes(B() As Byte)
'    'Call n.B.setBuffer(B())
'End Sub
'Friend Sub internal_reserveBytesAtEnd(Length As Long)
'    Dim appended_bytes() As Byte
'    ReDim appended_bytes(Length - 1)
'    'Call n.B.Append(appended_bytes)
'End Sub
'Friend Sub internal_appendBytes(B() As Byte)
'    'Call n.B.Append(B())
'End Sub
'Friend Function GetByteArray() As ByteArray
'    'Set GetByteArray = n.B
'End Function
'Friend Function internal_validate(B As ByteArray, Optional ValidateNonScientific As Boolean = False) As Boolean
''    If B.Length = 0 Then
''        internal_validate = False
''        Exit Function
''    End If
''    Dim i As Long, byteVal As Byte
''    For i = 0 To B.Length - 1
''        byteVal = B(i)
''        'If byteVal <= 9 Then Exit For
''        If ((byteVal >= 48) And (byteVal <= 57)) Then Exit For
''        If byteVal = POINT_SIGN Then Exit For
''        If byteVal = SCIENTIFIC_E Then Exit For
''        If byteVal = SCIENTIFIC_Ec Then Exit For
'''        If byteVal = SCIENTIFIC_F Then Exit For
'''        If byteVal = SCIENTIFIC_Fc Then Exit For
''        If byteVal = NEGATIVE_SIGN Then Exit For
''        If byteVal = POSITIVE_SIGN Then Exit For
''        '---------------------------------------
''        internal_validate = False
''        Exit Function
''    Next
''    internal_validate = True
'End Function
'Private Sub internal_clean()
''    Dim IntP As ByteArray
''    Set IntP = IntegeralValue
''    'Remove integeral value length to just 1.
''    If IntP.Length > 1 Then
''        Call n.Integer.Append(IntP.SubArray(0, IntP.Length - 1))
''    End If
''    'Cleanup the zeros at last of float part and first of integeral part.
''    Dim i As Long
''    For i = 0 To n.Float.Length - 1
''        If n.Float(i) <> 0 Then
''            Set n.Float = n.Float.SubArray(i)
''            Exit For
''        End If
''    Next
''    For i = IntP.Length - 1 To 0 Step -1
''        If IntP(i) <> 0 Then
''            Set n.Integer = n.Integer.SubArray(0, i)
''            Exit For
''        End If
''    Next
'End Sub
'
''=====================================================
''=====================================================
'
'Public Function Clone() As BigNumber
'    Dim dup As New BigNumber
'    Dim dupN As BIGNUM_STRUCT
'    dupN = n
'    'Set dupN.B = n.B.Clone
'    If IsZero Then
'        Call dup.Initialize
'    Else
'        Call dup.InitializeW(n)
'    End If
'    Set Clone = dup
'End Function
'
'Public Sub Clear()
'    'Call n.B.setBuffer(0)
'    'n.dblDotIndex = 0
'End Sub
'
'Public Property Get Sign() As NumberSign
'    Sign = n.Sign
'End Property
'Public Property Let Sign(Value As NumberSign)
'    n.Sign = Value
'    If Value = nsZero Then
'        Call Clear
'    Else
'        If IsZero Then _
'            Call SetLocalValue(1)
'    End If
'End Property
'
'Public Property Get IsPositive() As Boolean
'    IsPositive = (Sign = nsPositive)
'End Property
'Public Property Get IsNegative() As Boolean
'    IsNegative = (Sign = nsNegative)
'End Property
'Public Property Get IsZero() As Boolean
'    IsZero = ((n.Sign = nsZero) Or ((n.Integer.Length = 0) And (n.Float.Length = 0)))
'End Property
'Public Property Let IsZero(Value As Boolean)
'    Sign = IIf(Value, nsZero, nsPositive)
'End Property
'
'Public Property Get E() As Long 'BigNumber
''    If n.SE Is Nothing Then
''        Set E = New BigNumber
''        E.IsZero = True
''    Else
''        Set E = BigNumber(n.SE)
''    End If
'    E = n.SE
'End Property
'Public Property Let E(Value As Long)
'    n.SE = Value
'End Property
'
'Public Property Get Base() As Long
'    Base = n.base_N
'End Property
'Public Property Let Base(Value As Long)
'    n.base_N = Value
'End Property
'
'Public Property Get CalculatedAfterPoint() As ByteArray
''    If n.Float Is Nothing Then
''        Set CalculatedAfterPoint = ByteArray("0")
''    Else
''        Set CalculatedAfterPoint = n.Float.Clone
''    End If
'End Property
'Public Property Let CalculatedAfterPoint(Value As ByteArray)
'
'End Property
'Public Property Set CalculatedAfterPoint(Value As ByteArray)
'    CalculatedAfterPoint = Value
'End Property
'Public Property Get CalculatedIntegeralValue() As ByteArray
'    If n.Integer Is Nothing Then
'        Set CalculatedIntegeralValue = ByteArray("0")
'    Else
'        Set CalculatedIntegeralValue = n.Integer.Clone
'    End If
'End Property
'Public Property Let CalculatedIntegeralValue(Value As ByteArray)
'
'End Property
'Public Property Set CalculatedIntegeralValue(Value As ByteArray)
'    IntegeralValue = Value
'End Property
'
'Public Property Get AfterPoint() As ByteArray
'    If n.Float Is Nothing Then
'        Set AfterPoint = ByteArray("0")
'    Else
'        Set AfterPoint = n.Float.Clone
'    End If
'End Property
'Public Property Let AfterPoint(Value As ByteArray)
'
'End Property
'Public Property Set AfterPoint(Value As ByteArray)
'    AfterPoint = Value
'End Property
'Public Property Get IntegeralValue() As ByteArray
'    If n.Integer Is Nothing Then
'        Set IntegeralValue = ByteArray("0")
'    Else
'        Set IntegeralValue = n.Integer.Clone
'    End If
'End Property
'Public Property Let IntegeralValue(Value As ByteArray)
'
'End Property
'Public Property Set IntegeralValue(Value As ByteArray)
'    IntegeralValue = Value
'End Property
'
'Public Property Get DigitOfAfterPoint(Index As Long) As Integer
'
'End Property
'Public Property Let DigitOfAfterPoint(Index As Long, Value As Integer)
'
'End Property
'Public Property Get DigitOfIntegeralValue(Index As Long) As Integer
'
'End Property
'Public Property Let DigitOfIntegeralValue(Index As Long, Value As Integer)
'
'End Property
'
'
'Public Function BigNumberOf(bgNum) As BigNumber
'
'End Function
'
'Public Sub SetLocalValue(Value As Variant)
'    Dim VT As VbVarType
'    VT = VarType(Value)
'    If VT = VBObject Then
'        If Value Is Nothing Then Exit Sub
'        If TypeOf Value Is BigNumber Then
'            Call fromBigNumber(Value)
'        End If
'    ElseIf (VT And vbArray) = vbArray Then
'        If (VT And vbByte) = vbByte Then
'            n = FromByteArray(ByteArray(Value)).GetStruct
'        Else
'            n = FromByteArray(ByteArray(ArrayToByteArray(Value))).GetStruct
'        End If
'    ElseIf VT = vbBoolean Then
'        If Value Then
'            Call SetLocalValue(CLng(1))
'        Else
'            Call SetLocalValue(CLng(-1))
'        End If
'    ElseIf VT = vbString Then
'        n = FromString(CStr(Value)).GetStruct
'    Else
'        n = FromString(CStr(Value)).GetStruct
'    End If
'End Sub
'
'Public Sub fromBigNumber(ByVal bgNum As BigNumber)
'    n = bgNum.GetStruct
'    If n.Integer Is Nothing Then
'        Set n.Integer = New ByteArray
'    Else
'        Set n.Integer = n.Integer.Clone
'    End If
''    If n.SE Is Nothing Then
''        Set n.SE = New ByteArray
''    Else
''        Set n.SE = n.SE.Clone
''    End If
'    If n.Float Is Nothing Then
'        Set n.Float = New ByteArray
'    Else
'        Set n.Float = n.Float.Clone
'    End If
'End Sub
'
'Public Function ToDouble() As Double
'
'End Function
'
'Public Function toScientific(NumberOfDigitsAfterDecimal As Long) As String
'
'End Function
'Public Function toScientificBA() As ByteArray
'
'End Function
'Public Function ToString() As String
'    Dim RetVal As String
'    If n.Sign = sNegative Then
'        RetVal = "-"
'    ElseIf n.Sign = sPositive Then
'        If n.show_pos_sign Then
'            RetVal = "+"
'        End If
'    End If
'    If IsZero Then
'        RetVal = RetVal & "0"
'    Else
'        Dim i As Long, numVal As String
''        For i = 0 To n.B.Length - 1
''            numVal = numVal & (n.B(i) + 48)
''        Next
'        RetVal = RetVal * numVal
'    End If
'    ToString = RetVal
'End Function
'Public Function ToByteArray() As ByteArray
'    Set ToByteArray = ByteArray(ToString)
'End Function
'
'Public Function FromString(strBigNumber As String) As BigNumber
'    Set FromString = FromByteArray(ByteArray(strBigNumber))
'End Function
'Public Function FromByteArray(baBigNumber As ByteArray) As BigNumber
'    Dim baNum As ByteArray
'    Set baNum = baBigNumber
'
'    If baNum.IsEmpty Then throw Exps.InvalidArgumentValueException("Empty byte array.")
'    Dim keepBackup As BigNumber
'    Set keepBackup = Me.Clone
'
'    If Not internal_validate(baNum) Then throw Exps.InvalidArgumentValueException
'
'    Dim btSign As Byte
'    btSign = baNum.Last
'
'    On Error GoTo bn_error
'    Call Clear
'
'    If btSign = NEGATIVE_SIGN Then
'        Sign = sNegative
'        Set baNum = baNum.SubArray(0, baNum.Length - 1)
'    ElseIf btSign = POSITIVE_SIGN Then
'        Sign = sPositive
'        Set baNum = baNum.SubArray(0, baNum.Length - 1)
'    Else
'        If baNum.Length = 0 Then
'            Sign = sZero
'        ElseIf baNum.Length = 1 Then
'            If baNum(0) = "0" Then
'                Sign = sZero
'            Else
'                Sign = sPositive
'            End If
'        Else
'            Sign = sPositive
'        End If
'    End If
'
'    Dim fPointIndex As Long, eIndex As Long, i As Long, byteVal As Byte
'    fPointIndex = baNum.LastIndexOf(ByteArray("."))
'    IntegeralValue = baNum.SubArray(fPointIndex + 1)
'    eIndex = -1
'    For i = 0 To baNum.Length - 1
'        byteVal = baNum(i)
'        If (byteVal = SCIENTIFIC_E) Or (byteVal = SCIENTIFIC_Ec) Then 'Or (byteVal = SCIENTIFIC_F) Or (byteVal = SCIENTIFIC_Fc)
'            eIndex = i
'            Exit For
'        End If
'    Next
'    If fPointIndex = -1 Then
'        AfterPoint = ByteArray
'    Else
'        Dim afPointSIndex As Long
'        afPointSIndex = eIndex
'        byteVal = baNum.SubArray(afPointSIndex + 1, 1)
'        If byteVal = NEGATIVE_SIGN Or byteVal = POSITIVE_SIGN Then
'            afPointSIndex = afPointSIndex + 1
'        End If
'        AfterPoint = baNum.SubArray(afPointSIndex + 1, fPointIndex - afPointSIndex)
'    End If
'    If eIndex = -1 Then
'        n.SE = 0
'    Else
'        n.SE = CLng(baNum.SubArray(0, eIndex).ToString)
'    End If
'
'    Call internal_clean
'    Exit Function
'bn_error:
'    Call fromBigNumber(keepBackup)
'End Function
'
'Public Function GetByteAt(Index As Long) As Byte
'
'End Function
'Public Sub SetByteAt(Index As Long, Value As Byte)
'
'End Sub
'
'Public Property Get DotIndex() As Long
'    'DotIndex = n.dblDotIndex
'End Property
'Public Property Let DotIndex(Value As Long)
'    If Value < 0 Then throw Exps.NegativeArgumentException("Floating Point Index Can't Be Negative.")
'    'If Value > n.B.Length Then throw Exps.InvalidArgumentValueException("Floating Point Index Can't Be Greater Than Number Length.")
'    'n.dblDotIndex = Value
'End Property
'
'Public Property Get Length() As Long
''    Dim ln As Long
''    ln = n.B.Length
''    If n.dblDotIndex > 0 Then ln = ln + 1
''    If n.Sign = sNegative Then ln = ln + 1
''    Length = ln
'End Property
'Public Function BytesLength(Optional IncludePlus As Boolean = False) As Long
''    BytesLength = n.B.Length
''    If n.dblDotIndex > 0 Then _
''        BytesLength = BytesLength + 1
''
''    If Not IsZero Then
''        If IsPositive Then
''            If IncludePlus Then _
''                BytesLength = BytesLength + 1
''        Else
''            BytesLength = BytesLength + 1
''        End If
''    End If
'End Function
'Public Function CalculateLengthIn(RecordSize As Long, Optional IncludePlus As Boolean = False) As Long
''    Dim btsLength As Long
''    btsLength = BytesLength(IncludePlus)
''    CalculateLengthIn = (btsLength \ RecordSize) + IIf((btsLength Mod RecordSize) > 0, 1, 0)
'End Function
'
'
'Private Function NewBigNumber(n As BIGNUM_STRUCT) As BigNumber
'
'End Function
'
'Public Function NoSign() As BigNumber
'    Set NoSign = NewBigNumber(n)
'    If Not NoSign.IsZero Then _
'        NoSign.Sign = nsPositive
'End Function
'Public Function bnNot() As BigNumber 'internal
'    Dim i As Long, B As ByteArray
'    Set bnNot = Clone
'    Set B = bnNot.GetByteArray
'    If IsZero Then
'        Exit Function
'    ElseIf IsNegative Then
'        'Call bnNot.internal_reserveBytesAtEnd(1)
'        For i = 0 To B.Length - 1
'            B(i) = 9 - B(i)
'        Next
'        B.Last = (&H7F And B.Last) '0xxx xxxx
'        bnNot.Sign = nsPositive
'    ElseIf IsPositive Then
'        For i = 0 To B.Length - 1
'            B(i) = 9 - B(i)
'        Next
'        B.Last = (&H8F Or B.Last) '1xxx xxxx
'        bnNot.Sign = nsNegative
'    End If
'End Function
'Public Function bnAnd(bgNum2) As BigNumber ' bgNum2 As BigNumber
'
'End Function
'Public Function bnOr(bgNum2) As BigNumber  ' bgNum2 As BigNumber
'
'End Function
'Public Function bnXor(bgNum2) As BigNumber  ' bgNum2 As BigNumber
'
'End Function
'
'Public Sub Increament()
'
'End Sub
'Public Sub Decreament()
'
'End Sub
'
'Public Function bnAdd(bgNum2) As BigNumber  ' bgNum2 As BigNumber
'
'End Function
'Public Function bnSub(bgNum2) As BigNumber  ' bgNum2 As BigNumber
'
'End Function
'Public Function bnMul(bgNum2) As BigNumber  ' bgNum2 As BigNumber
'
'End Function
'Public Function bnDiv(bgNum2) As BigNumber  ' bgNum2 As BigNumber
'
'End Function
'Public Function bnMod(bgNum2) As BigNumber  ' bgNum2 As BigNumber
'
'End Function
'Public Function bnPow(bgNum2) As BigNumber  ' bgNum2 As BigNumber
'
'End Function
'Public Function bnLog(bgBaseNum) As BigNumber  ' bgNum2 As BigNumber
'
'End Function
'
'
'Public Function bnShiftLeft(TimesToShift As Long) As BigNumber
'
'End Function
'Public Function bnShiftRight(TimesToShift As Long) As BigNumber
'
'End Function
'Public Function bnRotateLeft(TimesToRotate As Long) As BigNumber
'
'End Function
'Public Function bnRotateRight(TimesToRotate As Long) As BigNumber
'
'End Function
'
'Public Function bnCompareIntegeralValue(bgNum2) As CompareValue
'    Dim b2 As BigNumber
'    Set b2 = BigNumberOf(bgNum2)
'
'
'
'End Function
'
'Public Function bnCompare(bgNum2) As CompareValue
'    Dim b2 As BigNumber
'    Set b2 = BigNumberOf(bgNum2)
'    Dim IntV As ByteArray, Float As ByteArray, E As Long 'MUST CHANGE!
'    Set IntV = b2.IntegeralValue
'    Set Float = b2.AfterPoint
'    E = b2.E
'
'    Dim b1Len As Long, b2Len As Long
'    b1Len = n.Integer.Length + n.SE
'    b2Len = b2.IntegeralValue.Length + b2.E
'
'    If b1Len = b2Len Then
'        Dim IntCompare As CompareValue
'        IntCompare = bnCompareIntegeralValue(b2)
'
'        If IntCompare = cvEqual Then
'
'        Else
'            bnCompare = IntCompare
'        End If
'    Else
'        If b1Len > b2Len Then
'            bnCompare = cvGreater
'        Else
'            bnCompare = cvLess
'        End If
'    End If
'End Function
'
'
'Private Function ICloneable_Clone() As Object
'    Set ICloneable_Clone = Clone
'End Function
'
'Private Property Get IAliasable_Alias() As String
'    IAliasable_Alias = CLASSALIASSTRING
'End Property
