VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Generics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Generics"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Public Enum equalToFlags
    AllMustEqual = 1
    SomeEqual = 2
    SomeNotEqual = 4
    AllMustNotEqual = 8

    SomeCompare = SomeEqual Or SomeNotEqual
    AllCompare = AllMustEqual Or AllMustNotEqual
    NotValue = 32
    AllValues = AllMustEqual Or SomeEqual Or SomeNotEqual Or AllMustNotEqual
End Enum


Private Declare Function API_ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Enum ShellWindowStyles
    swsHide = VbAppWinStyle.vbHide
    swsMaximizedFocus = VbAppWinStyle.vbMaximizedFocus
    swsMinimizedFocus = VbAppWinStyle.vbMinimizedFocus
    swsMinimizedNoFocus = VbAppWinStyle.vbMinimizedNoFocus
    swsNormalFocus = VbAppWinStyle.vbNormalFocus
    swsNormalNoFocus = VbAppWinStyle.vbNormalNoFocus
End Enum

Friend Sub Initialize()
    
End Sub


Public Function RotateLeft(Value As Long, ByVal times As Long) As Long
    RotateLeft = bitOperations.RotateLeft(Value, times)
End Function
Public Function RotateLeftI(Value As Integer, ByVal times As Long) As Integer
    RotateLeftI = bitOperations.RotateLeftI(Value, times)
End Function
Public Function RotateRight(Value As Long, ByVal times As Long) As Long
    RotateRight = bitOperations.RotateRight(Value, times)
End Function
Public Function RotateRightI(Value As Integer, ByVal times As Long) As Integer
    RotateRightI = bitOperations.RotateRightI(Value, times)
End Function
Public Function ShiftLeft(Value As Long, ByVal times As Long) As Long
    ShiftLeft = bitOperations.ShiftLeft(Value, times)
End Function
Public Function ShiftRight(Value As Long, ByVal times As Long) As Long
    ShiftRight = bitOperations.ShiftRight(Value, times)
End Function
Public Function Power2(Value As Long) As Long
    Power2 = bitOperations.Power2(Value)
End Function


Public Function bnShiftLeft(bgNum As BigNumber, TimesToShift As Long) As BigNumber
    Set bnShiftLeft = bgNum.bnShiftLeft(TimesToShift)
End Function
Public Function bnShiftRight(bgNum As BigNumber, TimesToShift As Long) As BigNumber
    Set bnShiftRight = bgNum.bnShiftRight(TimesToShift)
End Function
Public Function bnRotateLeft(bgNum As BigNumber, TimesToRotate As Long) As BigNumber
    Set bnRotateLeft = bgNum.bnRotateLeft(TimesToRotate)
End Function
Public Function bnRotateRight(bgNum As BigNumber, TimesToRotate As Long) As BigNumber
    Set bnRotateRight = bgNum.bnRotateRight(TimesToRotate)
End Function

Public Function bnAdd(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnAdd = bgNum1.bnAdd(bgNum2)
End Function
Public Function bnSub(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnSub = bgNum1.bnSub(bgNum2)
End Function
Public Function bnMul(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnMul = bgNum1.bnMul(bgNum2)
End Function
Public Function bnDiv(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnDiv = bgNum1.bnDiv(bgNum2)
End Function
Public Function bnMod(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnMod = bgNum1.bnMod(bgNum2)
End Function
Public Function bnPow(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnPow = bgNum1.bnPow(bgNum2)
End Function
Public Function bnLog(bgNum As BigNumber, bgBaseNum As BigNumber) As BigNumber
    Set bnLog = bgNum.bnLog(bgBaseNum)
End Function



Public Function MaxValue(vArray As Variant) As Variant
    '##BLOCK_DESCRIPTION Returns the maximum value in the array passed.
    '##PARAMETER_DESCRIPTION vArray An array - use Array() function.
    Dim max As Variant
    Dim i As Long, vt As VariantTypeConstants
    If Not IsArray(vArray) Then throw Exception("MaxValue Requires An Array As Parameter.")
    If Not ArraySize(vArray) > 0 Then throw ArgumentNullException("Empty Array.")
    For i = 0 To UBound(vArray)
        If i = 0 Then
            max = vArray(0)
            vt = VarType(vArray(0))
        Else
            If vArray(i) > max Then max = vArray(i)
            If Not VarType(vArray(i)) = vt Then throw Exception("Array Items Must Be Of The Same Type.")
        End If
    Next i
    MaxValue = max
End Function
Public Function MinValue(vArray As Variant) As Variant
    '##BLOCK_DESCRIPTION Returns the minimum value in the array passed.
    '##PARAMETER_DESCRIPTION vArray An array - use Array() function.
    Dim Min As Variant
    Dim i As Long, vt As VariantTypeConstants
    If Not IsArray(vArray) Then throw Exception("MinValue Requires An Array As Parameter.")
    If Not ArraySize(vArray) > 0 Then throw ArgumentNullException("Empty Array.")
    For i = 0 To UBound(vArray)
        If i = 0 Then
            Min = vArray(0)
            vt = VarType(vArray(0))
        Else
            If Not VarType(vArray(i)) = vt Then throw Exception("Array Items Must Be Of The Same Type.")
            If vArray(i) < Min Then Min = vArray(i)
        End If
    Next i
    MinValue = Min
End Function
Public Function Compare(str1 As String, str2 As String) As Boolean
    Compare = (str1 Like str2)
End Function
Public Function Equals(Object1 As Object, Object2 As Object) As Boolean
    Equals = (Object1 Is Object2)
End Function
Public Function TypesEquals(Object1 As Object, Object2 As Object) As Boolean
    TypesEquals = (TypeName(Object1) = TypeName(Object2))
End Function

Public Function R(ByVal ARGBColor As Long) As Byte
    ARGBColor = ARGBColor And &HFF
    R = ARGBColor 'And &HFF
End Function
Public Function G(ByVal ARGBColor As Long) As Byte
    ARGBColor = ARGBColor \ &H100
    G = ARGBColor And &HFF
End Function
Public Function B(ByVal ARGBColor As Long) As Byte
    ARGBColor = ARGBColor \ &H10000
    B = ARGBColor And &HFF
End Function
Public Function A(ByVal ARGBColor As Long) As Byte
    ARGBColor = ARGBColor \ &H1000000
    A = ARGBColor And &HFF
End Function
Public Function C(ByVal CMYKColor As Long) As Byte
    CMYKColor = CMYKColor And &HFF
    C = CMYKColor 'And &HFF
End Function
Public Function M(ByVal CMYKColor As Long) As Byte
    CMYKColor = CMYKColor \ &H100
    M = CMYKColor And &HFF
End Function
Public Function Y(ByVal CMYKColor As Long) As Byte
    CMYKColor = CMYKColor \ &H10000
    Y = CMYKColor And &HFF
End Function
Public Function K(ByVal CMYKColor As Long) As Byte
    CMYKColor = CMYKColor \ &H1000000
    K = CMYKColor And &HFF
End Function


Public Function LoadPlugin(Path As String, Optional EntryClassName) As IPlugin

End Function
Public Sub UnloadPlugin(IPlugin As IPlugin)

End Sub


Public Function GlobalFilters(Optional includes, Optional excludes) As GlobalFilters
    Dim retVal As GlobalFilters
    If IsMissing(includes) Then
        Dim p(1) As String
        p(0) = "*": p(1) = "*.*"
        includes = p
    End If
    If VarType(includes) = (vbArray Or vbString) Then
        retVal.IncludeTemplates = includes
    Else
        ReDim retVal.IncludeTemplates(0)
        retVal.IncludeTemplates(0) = CStr(includes)
    End If
    If Not IsMissing(excludes) Then
        If VarType(excludes) = (vbArray Or vbString) Then
            retVal.ExcludeTemplates = excludes
        Else
            ReDim retVal.ExcludeTemplates(0)
            retVal.ExcludeTemplates(0) = CStr(excludes)
        End If
    End If
    GlobalFilters = retVal
End Function
Public Function FileFilters(Optional includes, Optional excludes) As GlobalFilters
    Dim retVal As GlobalFilters
    If IsMissing(includes) Then
        Dim p(1) As String
        p(0) = "*": p(1) = "*.*"
        includes = p
    End If
    If VarType(includes) = (vbArray Or vbString) Then
        retVal.IncludeTemplates = includes
    Else
        ReDim retVal.IncludeTemplates(0)
        retVal.IncludeTemplates(0) = CStr(includes)
    End If
    If Not IsMissing(excludes) Then
        If VarType(excludes) = (vbArray Or vbString) Then
            retVal.ExcludeTemplates = excludes
        Else
            ReDim retVal.ExcludeTemplates(0)
            retVal.ExcludeTemplates(0) = CStr(excludes)
        End If
    End If
    FileFilters = retVal
End Function


Private Function IncludeFilter_single(fs() As String, Expression As String) As Boolean
    Dim i As Long
    For i = 0 To ArraySize(fs) - 1
        If (Expression) Like CStr(fs(i)) Then
            IncludeFilter_single = True
            Exit Function
        End If
    Next
    IncludeFilter_single = False
End Function
Public Function IsFilterIncluded(Filter As GlobalFilters, Expression As String) As Boolean
    IsFilterIncluded = ((IncludeFilter_single(Filter.IncludeTemplates, Expression)) And (Not IncludeFilter_single(Filter.ExcludeTemplates, Expression)))
End Function

Public Sub InitializeRandomGenerator(): Call Randomize: End Sub
Public Function GetNextRandom(Optional UppderValue As Double = 1, Optional LowerValue As Double = 0, Optional Initialize_NewInstance_Of_RandomGenerator) As Double
    If Not IsMissing(Initialize_NewInstance_Of_RandomGenerator) Then _
        Call InitializeRandomGenerator
    GetNextRandom = LowerValue + ((UppderValue - LowerValue) * Rnd)
End Function

Public Function equalTo(Flags As equalToFlags, EnumerateValue, ParamArray Args() As Variant) As Boolean
    Dim iArgs() As Variant
    iArgs = Args
    equalTo = baseMethods.equalToArr(Flags, EnumerateValue, iArgs)
End Function

Public Function ByteArrayToString(B() As Byte, Optional Length As Long = -1) As String
    ByteArrayToString = baseMethods.ByteArrayToString(B, Length)
End Function
Public Function ByteArrayToLong(B() As Byte) As Long
    ByteArrayToLong = baseMethods.ByteArrayToLong(B)
End Function
Public Function ByteArrayToInteger(B() As Byte) As Integer
    ByteArrayToInteger = baseMethods.ByteArrayToInteger(B)
End Function
Public Function ByteArrayToDouble(B() As Byte) As Double
    ByteArrayToDouble = baseMethods.ByteArrayToDouble(B)
End Function
Public Function ByteArrayToSingle(B() As Byte) As Single
    ByteArrayToSingle = baseMethods.ByteArrayToSingle(B)
End Function
Public Function ByteArrayToDate(B() As Byte) As Date
    ByteArrayToDate = baseMethods.ByteArrayToDate(B)
End Function
Public Function ByteArrayToCurrency(B() As Byte) As Currency
    ByteArrayToCurrency = baseMethods.ByteArrayToCurrency(B)
End Function
Public Function ByteArrayToBoolean(B() As Byte) As Boolean
    ByteArrayToBoolean = baseMethods.ByteArrayToBoolean(B)
End Function

Public Function StringToByteArray(Value As String, Optional Length As Long = -1) As Byte()
    StringToByteArray = baseMethods.StringToByteArray(Value, Length)
End Function
Public Function LongToByteArray(Value As Long) As Byte()
    LongToByteArray = baseMethods.LongToByteArray(Value)
End Function
Public Function IntegerToByteArray(Value As Integer) As Byte()
    IntegerToByteArray = baseMethods.IntegerToByteArray(Value)
End Function
Public Function DoubleToByteArray(Value As Double) As Byte()
    DoubleToByteArray = baseMethods.DoubleToByteArray(Value)
End Function
Public Function DateToByteArray(Value As Date) As Byte()
    DateToByteArray = baseMethods.DateToByteArray(Value)
End Function
Public Function CurrencyToByteArray(Value As Currency) As Byte()
    CurrencyToByteArray = baseMethods.CurrencyToByteArray(Value)
End Function
Public Function BooleanToByteArray(Value As Boolean) As Byte()
    BooleanToByteArray = baseMethods.BooleanToByteArray(Value)
End Function


Public Function MemoryToByteArray(ByVal targetAddress As Long, SourceSize As Long, Optional times As Long = 1) As Byte()
    MemoryToByteArray = baseMethods.MemoryToByteArray(targetAddress, SourceSize, times)
End Function
Public Sub CopyMemoryToByteArray(ByVal targetAddress As Long, SourceSize As Long, targetByteArray() As Byte, Optional times As Long = 1)
    Call baseMethods.CopyMemoryToByteArray(targetAddress, SourceSize, targetByteArray, times)
End Sub
Public Sub ByteArrayToMemory(ByVal targetAddress As Long, targetByteArray() As Byte, BytesToCopy As Long, Optional FillWithNull As Boolean = False)
    Call baseMethods.ByteArrayToMemory(targetAddress, targetByteArray, BytesToCopy, FillWithNull)
End Sub


Public Function RepeatString(str As String, Count As Long, Optional Splitter As String = "") As String
    Dim i As Long
    For i = 1 To Count
        RepeatString = RepeatString & str & Splitter
    Next
End Function



'Public Sub ByteArrayToParamArrayArgs(targetByteArray() As Byte, target)
'    Call baseMethods.ByteArrayToParamArrayArgs(targetByteArray, target)
'End Sub
Public Function ArrayCompare(a1, a2, Optional LengthToCompare As Long = -1) As Boolean
    ArrayCompare = baseMethods.ArrayCompare(a1, a2, LengthToCompare)
End Function
Public Function BinaryCompare(a1, a2, Optional LengthToCompare As Long = -1) As Boolean
    BinaryCompare = baseMethods.BinaryCompare(a1, a2, LengthToCompare)
End Function
Public Function ArrayToByteArray(targetArray) As Byte()
    ArrayToByteArray = baseMethods.ArrayToByteArray(targetArray)
End Function
Public Function GetSubByteArray(targetByteArray() As Byte, ByVal StartIndex As Long, Optional ByVal Length As Long = -1) As Byte()
    GetSubByteArray = baseMethods.GetSubByteArray(targetByteArray, StartIndex, Length)
End Function
Public Function GetByteArraySomeLength(targetArray() As Byte, Length As Long) As Byte()
    GetByteArraySomeLength = baseMethods.GetByteArraySomeLength(targetArray, Length)
End Function
Public Function IsByteArrayNumeric(bArr() As Byte) As Boolean
    IsByteArrayNumeric = baseMethods.IsByteArrayNumeric(bArr)
End Function
Public Function IsByteArrayAlphabetic(bArr() As Byte) As Boolean
    IsByteArrayAlphabetic = baseMethods.IsByteArrayAlphabetic(bArr)
End Function
Public Function IsByteArrayLikeAnother(fA() As Byte, sA() As Byte) As Boolean
    IsByteArrayLikeAnother = baseMethods.IsByteArrayLikeAnother(fA, sA)
End Function
Public Function ByteArrayToUpper(bArr() As Byte) As Byte()
    ByteArrayToUpper = baseMethods.ByteArrayToUpper(bArr)
End Function
Public Function ByteArrayToLower(bArr() As Byte) As Byte()
    ByteArrayToLower = baseMethods.ByteArrayToLower(bArr)
End Function
Public Function FillArray(targetArray, whattofill, OutputSize As Long, Optional FillToLeft = True)
    FillArray = baseMethods.FillArray(targetArray, whattofill, OutputSize, FillToLeft)
End Function
Public Sub InsertArrayIndex(targetArray, Index, Item)
    Call baseMethods.InsertArrayIndex(targetArray, Index, Item)
End Sub
Public Sub InsertArrayIndexArray(targetArray, Index, insertArray, Optional Length As Long = -1)
    Call baseMethods.InsertArrayIndexArray(targetArray, Index, insertArray, Length)
End Sub
Public Function AppendToArray(sourceArray, Item)
    Call baseMethods.AppendToArray(sourceArray, Item)
End Function
Public Function AppendArrayToArray(sourceArray, targetArray)
    Call baseMethods.AppendArrayToArray(sourceArray, targetArray)
End Function




Public Function Log10(ByVal Number As Variant) As Double
    Dim Num As Double
    Const Lg10 As Double = 2.30258509299405
    Num = CDbl(Number)
    Log10 = Log(Num) / Lg10
End Function
Public Function LogX(ByVal Number As Variant, ByVal X As Variant) As Double
    Dim Num As Double
    Dim nBase As Double
    Num = CDbl(Number)
    nBase = CDbl(X)
    LogX = Log(Num) / Log(nBase)
End Function
Public Function Ceil(ByVal Number As Variant) As Double
    Dim Num As Double
    Num = CDbl(Number)
    Ceil = Int(Num + 1)
End Function
Public Function Floor(ByVal Number As Variant) As Double
    Dim Num As Double
    Num = CDbl(Number)
    Floor = Int(Num)
End Function
Public Function AbsCeil(ByVal Number As Variant, Optional ByVal KeepSign As Boolean = False) As Double
    Dim Num As Double, Sign As Integer
    Num = CDbl(Number)
    Sign = Sgn(Num)
    AbsCeil = IIf(KeepSign, Sign, 1) * (Int(Abs(Num) + 1))
End Function
Public Function AbsFloor(ByVal Number As Variant, Optional ByVal KeepSign As Boolean = False) As Double
    Dim Num As Double, Sign As Integer
    Num = CDbl(Number)
    Sign = Sgn(Num)
    AbsFloor = IIf(KeepSign, Sign, 1) * (Int(Abs(Num)))
End Function
' arc sine
' error if value is outside the range [-1,1]
Public Function ArcSin(Value As Double) As Double
    If Abs(Value) <> 1 Then
        ArcSin = Atn(Value / Sqr(1 - Value * Value))
    Else
        ArcSin = 1.5707963267949 * Sgn(Value)
    End If
End Function
' arc cosine
' error if NUMBER is outside the range [-1,1]
Public Function ArcCos(ByVal Number As Double) As Double
    If Abs(Number) <> 1 Then
        ArcCos = 1.5707963267949 - Atn(Number / Sqr(1 - Number * Number))
    ElseIf Number = -1 Then
        ArcCos = 3.14159265358979
    Else
        throw OutOfRangeException
    End If
    'elseif number=1 --> Acos=0 (implicit)
End Function
' arc cotangent
' error if NUMBER is zero
Public Function ArcCot(Value As Double) As Double
    ArcCot = Atn(1 / Value)
End Function
' arc secant
' error if value is inside the range [-1,1]
Public Function ArcSec(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ASec = ACos(1 / value)
    If Abs(Value) <> 1 Then
        ArcSec = 1.5707963267949 - Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ArcSec = 3.14159265358979 * Sgn(Value)
    End If
End Function
' arc cosecant
' error if value is inside the range [-1,1]
Public Function ArcCsc(Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ACsc = ASin(1 / value)
    If Abs(Value) <> 1 Then
        ArcCsc = Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ArcCsc = 1.5707963267949 * Sgn(Value)
    End If
End Function
' hyperbolic sine
Public Function SinH(Value As Double) As Double
    Dim Temp As Double
    Temp = exp(Value)
    SinH = (Temp - 1 / Temp) / 2
End Function
' hyperbolic cosine
Public Function CosH(Value As Double) As Double
    Dim Temp As Double
    Temp = exp(Value)
    CosH = (Temp + 1 / Temp) / 2
End Function
' hyperbolic tangent
Public Function TanH(Value As Double) As Double
    Dim Temp As Double
    Temp = exp(Value)
    TanH = (Temp - 1 / Temp) / (Temp + 1 / Temp)
End Function
' hyperbolic cotangent
Public Function CotH(Value As Double) As Double
    Dim Temp As Double
    Temp = exp(Value)
    CotH = (Temp + 1 / Temp) / (Temp - 1 / Temp)
End Function
' hyperbolic secant
Public Function SecH(Value As Double) As Double
    Dim Temp As Double
    Temp = exp(Value)
    SecH = 2 / (Temp + 1 / Temp)
End Function
' hyperbolic cosecant
Public Function CscH(Value As Double) As Double
    Dim Temp As Double
    Temp = exp(Value)
    CscH = 2 / (Temp - 1 / Temp)
End Function
' hyperbolic arc sine
Public Function ArcSinH(Value As Double) As Double
    ArcSinH = Log(Value + Sqr(Value * Value + 1))
End Function
' hyperbolic arc cosine
' error if NUMBER is inside the range [-1,1]
Public Function ArcCosH(Value As Double) As Double
    ArcCosH = Log(Value + Sqr(Value * Value - 1))
End Function
' hyperbolic arc tangent
' error if value is zero
Public Function ArcTanH(Value As Double) As Double
    ArcTanH = Log((1 / Value + 1) / (1 / Value - 1)) / 2
End Function
' hyperbolic arc cotangent
Public Function ArcCotH(Value As Double) As Double
    ArcCotH = Log((Value + 1) / (Value - 1)) / 2
End Function
' hyperbolic arc secant
' error if value is outside the range [-1,1]
Public Function ArcSecH(Value As Double) As Double
    ArcSecH = Log((Sqr(1 - Value * Value) + 1) / Value)
End Function
' hyperbolic arc cosecant
Public Function ArcCscH(Value As Double) As Double
    ArcCscH = Log((Sgn(Value) * Sqr(1 + Value * Value) + 1) / Value)
End Function




Public Function CLongArray(ArrayValue As Variant) As Long()
    '##BLOCK_DESCRIPTION Converts an array to a long array
    Dim vRet() As Long, i As Long, l As Long, u As Long
    If Not IsArray(ArrayValue) Then _
        throw InvalidArgumentTypeException("Parameter is not an array")
    If IsEmptyArray(ArrayValue) Then _
        throw Exception("Array is empty")
    l = LBound(ArrayValue)
    u = UBound(ArrayValue)
    ReDim vRet(l To u)
    On Error GoTo Failure
    For i = l To u
        vRet(i) = CLng(ArrayValue(i))
    Next i
    On Error GoTo 0
    CLongArray = vRet
    Exit Function
Failure:
End Function
Public Function CIntegerArray(ArrayValue As Variant) As Integer()
    '##BLOCK_DESCRIPTION Converts an array to a integer array
    Dim vRet() As Integer, i As Long, l As Long, u As Long
    If Not IsArray(ArrayValue) Then _
        throw InvalidArgumentTypeException("Parameter is not an array")
    If IsEmptyArray(ArrayValue) Then _
        throw Exception("Array is empty")
    l = LBound(ArrayValue)
    u = UBound(ArrayValue)
    ReDim vRet(l To u)
    On Error GoTo Failure
    For i = l To u
        vRet(i) = CInt(ArrayValue(i))
    Next i
    On Error GoTo 0
    CIntegerArray = vRet
    Exit Function
Failure:
End Function
Public Function CByteArray(ArrayValue As Variant) As Byte()
    '##BLOCK_DESCRIPTION Converts an array to a byte array
    Dim vRet() As Byte, i As Long, l As Long, u As Long
    If Not IsArray(ArrayValue) Then _
        throw InvalidArgumentTypeException("Parameter is not an array")
    If IsEmptyArray(ArrayValue) Then _
        throw Exception("Array is empty")
    l = LBound(ArrayValue)
    u = UBound(ArrayValue)
    ReDim vRet(l To u)
    On Error GoTo Failure
    For i = l To u
        vRet(i) = CByte(ArrayValue(i))
    Next i
    On Error GoTo 0
    CByteArray = vRet
    Exit Function
Failure:
End Function
Public Function CDecimalArray(ArrayValue As Variant) As Variant()
    '##BLOCK_DESCRIPTION Converts an array to a decimal array
    Dim vRet() As Variant, i As Long, l As Long, u As Long
    If Not IsArray(ArrayValue) Then _
        throw InvalidArgumentTypeException("Parameter is not an array")
    If IsEmptyArray(ArrayValue) Then _
        throw Exception("Array is empty")
    l = LBound(ArrayValue)
    u = UBound(ArrayValue)
    ReDim vRet(l To u)
    On Error GoTo Failure
    For i = l To u
        vRet(i) = CDec(ArrayValue(i))
    Next i
    On Error GoTo 0
    CDecimalArray = vRet
    Exit Function
Failure:
End Function
Public Function CCurrencyArray(ArrayValue As Variant) As Currency()
    '##BLOCK_DESCRIPTION Converts an array to a currency array
    Dim vRet() As Currency, i As Long, l As Long, u As Long
    If Not IsArray(ArrayValue) Then _
        throw InvalidArgumentTypeException("Parameter is not an array")
    If IsEmptyArray(ArrayValue) Then _
        throw Exception("Array is empty")
    l = LBound(ArrayValue)
    u = UBound(ArrayValue)
    ReDim vRet(l To u)
    On Error GoTo Failure
    For i = l To u
        vRet(i) = CCur(ArrayValue(i))
    Next i
    On Error GoTo 0
    CCurrencyArray = vRet
    Exit Function
Failure:
End Function
Public Function CSingleArray(ArrayValue As Variant) As Single()
    '##BLOCK_DESCRIPTION Converts an array to a single array
    Dim vRet() As Single, i As Long, l As Long, u As Long
    If Not IsArray(ArrayValue) Then _
        throw InvalidArgumentTypeException("Parameter is not an array")
    If IsEmptyArray(ArrayValue) Then _
        throw Exception("Array is empty")
    l = LBound(ArrayValue)
    u = UBound(ArrayValue)
    ReDim vRet(l To u)
    On Error GoTo Failure
    For i = l To u
        vRet(i) = CSng(ArrayValue(i))
    Next i
    On Error GoTo 0
    CSingleArray = vRet
    Exit Function
Failure:
End Function
Public Function CDoubleArray(ArrayValue As Variant) As Double()
    '##BLOCK_DESCRIPTION Converts an array to a double array
    Dim vRet() As Double, i As Long, l As Long, u As Long
    If Not IsArray(ArrayValue) Then _
        throw InvalidArgumentTypeException("Parameter is not an array")
    If IsEmptyArray(ArrayValue) Then _
        throw Exception("Array is empty")
    l = LBound(ArrayValue)
    u = UBound(ArrayValue)
    ReDim vRet(l To u)
    On Error GoTo Failure
    For i = l To u
        vRet(i) = CDbl(ArrayValue(i))
    Next i
    On Error GoTo 0
    CDoubleArray = vRet
    Exit Function
Failure:
End Function
Public Function CDateArray(ArrayValue As Variant) As Date()
    '##BLOCK_DESCRIPTION Converts an array to a date array
    Dim vRet() As Date, i As Long, l As Long, u As Long
    If Not IsArray(ArrayValue) Then _
        throw InvalidArgumentTypeException("Parameter is not an array")
    If IsEmptyArray(ArrayValue) Then _
        throw Exception("Array is empty")
    l = LBound(ArrayValue)
    u = UBound(ArrayValue)
    ReDim vRet(l To u)
    On Error GoTo Failure
    For i = l To u
        vRet(i) = CDate(ArrayValue(i))
    Next i
    On Error GoTo 0
    CDateArray = vRet
    Exit Function
Failure:
End Function
Public Function CBooleanArray(ArrayValue As Variant) As Boolean()
    '##BLOCK_DESCRIPTION Converts an array to a boolean array
    Dim vRet() As Boolean, i As Long, l As Long, u As Long
    If Not IsArray(ArrayValue) Then _
        throw InvalidArgumentTypeException("Parameter is not an array")
    If IsEmptyArray(ArrayValue) Then _
        throw Exception("Array is empty")
    l = LBound(ArrayValue)
    u = UBound(ArrayValue)
    ReDim vRet(l To u)
    On Error GoTo Failure
    For i = l To u
        vRet(i) = CBool(ArrayValue(i))
    Next i
    On Error GoTo 0
    CBooleanArray = vRet
    Exit Function
Failure:
End Function
Public Function CStringArray(ArrayValue As Variant) As String()
    '##BLOCK_DESCRIPTION Converts an array to a string array
    Dim vRet() As String, i As Long, l As Long, u As Long
    If Not IsArray(ArrayValue) Then _
        throw InvalidArgumentTypeException("Parameter is not an array")
    If IsEmptyArray(ArrayValue) Then _
        throw Exception("Array is empty")
    l = LBound(ArrayValue)
    u = UBound(ArrayValue)
    ReDim vRet(l To u)
    On Error GoTo Failure
    For i = l To u
        vRet(i) = CStr(ArrayValue(i))
    Next i
    On Error GoTo 0
    CStringArray = vRet
    Exit Function
Failure:
End Function




Public Function ShellExecute(ByVal hWnd As Long, ByVal Operation As String, ByVal FileName As String, ByVal Parameters As String, ByVal Directory, Optional ByVal ShowCommand As ShellWindowStyles = ShellWindowStyles.swsNormalFocus) As Long
    ShellExecute = API_ShellExecute(hWnd, Operation, FileName, Parameters, Directory, ShowCommand)
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
