VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"BigNumber"
'@PROJECT_LICENSE

''<class name="mMath" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "mMath" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

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
'<section Creators>
'
'Public Function BigNumber(Optional InitialValue) As BigNumber
'
'End Function
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


Public Function RotateLeft(ByVal Value As Long, ByVal Times As Long) As Long
    RotateLeft = bitOperations.RotateLeft(Value, Times)
End Function
Public Function RotateLeftI(ByVal Value As Integer, ByVal Times As Long) As Integer
    RotateLeftI = bitOperations.RotateLeftI(Value, Times)
End Function
Public Function RotateRight(ByVal Value As Long, ByVal Times As Long) As Long
    RotateRight = bitOperations.RotateRight(Value, Times)
End Function
Public Function RotateRightI(ByVal Value As Integer, ByVal Times As Long) As Integer
    RotateRightI = bitOperations.RotateRightI(Value, Times)
End Function
Public Function ShiftLeft(ByVal Value As Long, ByVal Times As Long) As Long
    ShiftLeft = mHelper.ShiftLeft(Value, Times)
End Function
Public Function ShiftRight(ByVal Value As Long, ByVal Times As Long) As Long
    ShiftRight = mHelper.ShiftRight(Value, Times)
End Function
'Public Function Power2(ByVal Value As Long) As Long
'    Power2 = bitOperations.Power2(Value)
'End Function

''<summary>Returns the ceiling of the value.</summary>
Public Function Ceiling(ByVal Value As Double) As Double
    If Value < 0# Then
        Ceiling = Fix(Value)
    Else
        Ceiling = -Int(-Value)
    End If
End Function
''<summary>Returns the floor of the value.</summary>
Public Function Floor(ByVal Value As Double) As Double
    Floor = Int(Value)
End Function

Public Function GetFloatingPoint(ByVal Value As Double) As Long
    GetFloatingPoint = Value - Int(Value)
End Function

''<summary>Represents the constant value of PI (3.141592).</summary>
Public Property Get PI() As Double
    PI = C_PI
End Property
''<summary>Represents the constant value of Euler's number (2.718281).</summary>
Public Property Get E() As Double
    E = baseConstants.E_Number
End Property

''<summary>Converts a degree into radian.</summary>
Public Function DegreeToRadian(ByVal Degree As Double) As Double
    DegreeToRadian = ((Degree * C_PI) / 180#)
End Function
''<summary>Converts a radian into degree.</summary>
Public Function RadianToDegree(ByVal Radian As Double) As Double
    RadianToDegree = ((Radian * 180#) / C_PI)
End Function
''<summary>Converts a radian into grad.</summary>
Public Function RadianToGrad(ByVal Radian As Double) As Double
    RadianToGrad = ((Radian * 120#) / C_PI)
End Function
''<summary>Converts a degree into grad.</summary>
Public Function DegreeToGrad(ByVal Radian As Double) As Double
    DegreeToGrad = ((Radian * 120#) / C_PI)
End Function
''<summary>Converts a grad into degree.</summary>
Public Function GradToDegree(ByVal Grad As Double) As Double
    GradToDegree = ((Grad * 180#) / 120#)
End Function
''<summary>Converts a grad into radian.</summary>
Public Function GradToRadian(ByVal Grad As Double) As Double
    GradToRadian = ((Grad * C_PI) / 120#)
End Function



''<summary>Select the maximum value between two given argument.</summary>
''<remarks>Both arguments must be numeric or string.</remarks>
Public Function Max(ByRef Param1 As Variant, ByRef Param2 As Variant) As Variant
    On Error GoTo errCatch
    If Param1 < Param2 Then
        Max = Param2
    Else
        Max = Param1
    End If
    Exit Function
errCatch:
    throw Exps.InvalidArgumentException("Values could not be compared.")
End Function
''<summary>Select the minimum value between two given argument.</summary>
''<remarks>Both arguments must be numeric or string.</remarks>
Public Function Min(ByRef Param1 As Variant, ByRef Param2 As Variant) As Variant
    On Error GoTo errCatch
    If Param1 > Param2 Then
        Min = Param2
    Else
        Min = Param1
    End If
    Exit Function
errCatch:
    throw Exps.InvalidArgumentException("Values could not be compared.")
End Function
''<summary>Returns the Log10 of the given value.</summary>
Public Function Log10(ByVal Value As Variant) As Double
    Dim Num As Double
    Const Lg10 As Double = 2.30258509299405
    Num = CDbl(Value)
    Log10 = Log(Num) / Lg10
End Function
''<summary>Returns the logarithm of the given value using base value.</summary>
Public Function LogX(ByVal Value As Variant, ByVal Base As Variant) As Double
    Dim Num As Double, nBase As Double
    Num = CDbl(Value)
    nBase = CDbl(Base)
    LogX = Log(Num) / Log(nBase)
End Function
''<summary>Returns the natural logarithm of the given value.</summary>
Public Function ln(ByVal Value As Variant) As Double
    Dim Num As Double
    Num = CDbl(Value)
    ln = Log(Num) / Log(baseConstants.E_Number)
End Function
''<summary>Returns the unsigned ceiling value of the given number.</summary>
Public Function AbsCeil(ByVal Value As Variant, Optional ByVal KeepSign As Boolean = False) As Double
    Dim Num As Double, Sign As Integer
    Num = CDbl(Value)
    Sign = Sgn(Num)
    AbsCeil = IIf(KeepSign, Sign, 1&) * (Int(Abs(Num) + 1&))
End Function
''<summary>Returns the unsigned floor value of the given number.</summary>
Public Function AbsFloor(ByVal Value As Variant, Optional ByVal KeepSign As Boolean = False) As Double
    Dim Num As Double, Sign As Integer
    Num = CDbl(Value)
    Sign = Sgn(Num)
    AbsFloor = IIf(KeepSign, Sign, 1&) * (Int(Abs(Num)))
End Function

''<summary>Indicates the hypotenuse of the given edges.</summary>
Public Function Hypotenuse(ByVal X As Double, ByVal Y As Double, Optional Z) As Double 'vatar
    Dim ZBuff As Double
    If Not IsMissing(Z) Then ZBuff = CDbl(Z) ^ 2&
    Hypotenuse = Math.Sqr((X ^ 2&) + (Y ^ 2&) + (ZBuff))
End Function

''<summary>Indicates the value of the ArcSin of the given radian.</summary>
''<remarks>An Error occures if value is outside the range [-1,1].</remarks>
Public Function ArcSin(ByVal Value As Double) As Double
    If Abs(Value) <> 1 Then
        ArcSin = Atn(Value / Sqr(1 - Value * Value))
    Else
        ArcSin = 1.5707963267949 * Sgn(Value)
    End If
End Function
''<summary>Indicates the value of the ArcCos of the given radian.</summary>
''<remarks>An Error occures if value is outside the range [-1,1].</remarks>
Public Function ArcCos(ByVal Value As Double) As Double
    If Abs(Value) <> 1 Then
        ArcCos = 1.5707963267949 - Atn(Value / Sqr(1# - Value * Value))
    ElseIf Value = -1 Then
        ArcCos = 3.14159265358979
    Else
        throw Exps.InvalidOperationException
    End If
    'elseif number=1 --> Acos=0 (implicit)
End Function
''<summary>Indicates the value of the ArcCot of the given radian.</summary>
''<remarks>An Error occures error if value is zero.</remarks>
Public Function ArcCot(ByVal Value As Double) As Double
    ArcCot = Atn(1# / Value)
End Function
''<summary>Indicates the value of the ArcSec of the given radian.</summary>
''<remarks>An Error occures error if value is inside the range [-1,1].</remarks>
Public Function ArcSec(ByVal Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ASec = ACos(1 / value)
    If Abs(Value) <> 1& Then
        ArcSec = 1.5707963267949 - Atn((1# / Value) / Sqr(1# - (1# / (Value * Value))))
    Else
        ArcSec = 3.14159265358979 * Sgn(Value)
    End If
End Function
''<summary>Indicates the value of the ArcCsc of the given radian.</summary>
''<remarks>An Error occures error if value is inside the range [-1,1].</remarks>
Public Function ArcCsc(ByVal Value As Double) As Double
    ' NOTE: the following lines can be replaced by a single call
    '            ACsc = ASin(1 / value)
    If Abs(Value) <> 1 Then
        ArcCsc = Atn((1 / Value) / Sqr(1 - 1 / (Value * Value)))
    Else
        ArcCsc = 1.5707963267949 * Sgn(Value)
    End If
End Function
''<summary>Indicates the value of the SinH(hyperbolic sine) of the given radian.</summary>
Public Function SinH(ByVal Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    SinH = (Temp - 1 / Temp) / 2
End Function
''<summary>Indicates the value of the CosH(hyperbolic cosine) of the given radian.</summary>
Public Function CosH(ByVal Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    CosH = (Temp + 1 / Temp) / 2
End Function
''<summary>Indicates the value of the TanH(hyperbolic tangent) of the given radian.</summary>
Public Function TanH(ByVal Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    TanH = (Temp - 1 / Temp) / (Temp + 1 / Temp)
End Function
''<summary>Indicates the value of the CotH(hyperbolic cotangent) of the given radian.</summary>
Public Function CotH(ByVal Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    CotH = (Temp + 1 / Temp) / (Temp - 1 / Temp)
End Function
''<summary>Indicates the value of the SecH(hyperbolic secant) of the given radian.</summary>
Public Function SecH(ByVal Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    SecH = 2 / (Temp + 1 / Temp)
End Function
''<summary>Indicates the value of the CscH(hyperbolic cosecant) of the given radian.</summary>
Public Function CscH(ByVal Value As Double) As Double
    Dim Temp As Double
    Temp = Exp(Value)
    CscH = 2 / (Temp - 1 / Temp)
End Function
''<summary>Indicates the value of the ArcSinH(hyperbolic arc sine) of the given radian.</summary>
Public Function ArcSinH(ByVal Value As Double) As Double
    ArcSinH = Log(Value + Sqr(Value * Value + 1))
End Function
''<summary>Indicates the value of the ArcCosH(hyperbolic arc cosine) of the given radian.</summary>
''<remarks>An Error occures error if value is inside the range [-1,1].</remarks>
Public Function ArcCosH(ByVal Value As Double) As Double
    ArcCosH = Log(Value + Sqr(Value * Value - 1))
End Function
''<summary>Indicates the value of the ArcTanH(hyperbolic arc tangent) of the given radian.</summary>
''<remarks>An Error occures error if value is zero.</remarks>
Public Function ArcTanH(ByVal Value As Double) As Double
    ArcTanH = Log((1 / Value + 1) / (1 / Value - 1)) / 2
End Function
''<summary>Indicates the value of the ArcCotH(hyperbolic arc cotangent) of the given radian.</summary>
Public Function ArcCotH(ByVal Value As Double) As Double
    ArcCotH = Log((Value + 1) / (Value - 1)) / 2
End Function
''<summary>Indicates the value of the ArcSecH(hyperbolic arc secant) of the given radian.</summary>
''<remarks>An Error occures error if value is outside the range [-1,1].</remarks>
Public Function ArcSecH(ByVal Value As Double) As Double
    ArcSecH = Log((Sqr(1 - Value * Value) + 1) / Value)
End Function
''<summary>Indicates the value of the ArcCscH(hyperbolic arc cosecant) of the given radian.</summary>
Public Function ArcCscH(ByVal Value As Double) As Double
    ArcCscH = Log((Sgn(Value) * Sqr(1 + Value * Value) + 1) / Value)
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
