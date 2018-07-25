VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="Convert" static tags="">
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
Const CLASSID As String = "Convert" '#DO NOT REMOVE IT
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
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

Public Function CInt64(ByRef Expression As Variant) As Int64
    If IsMissing(Expression) Then Exps.ArgumentNullException
    If IsNull(Expression) Then throw Exps.ArgumentNullException

    Dim VT As VariantTypes
    VT = VariantType(Expression)
    If (VT And VT_BYREF) = VT_BYREF Then VT = VT Xor VT_BYREF
    Select Case VT
        Case VT_CY
            Dim CY As Currency
            CY = (CCur(Expression) * 0.0001@)
            Call memcpy(CInt64, CY, VLEN_CURRENCY)
        Case VT_BSTR, VT_LPSTR, VT_LPWSTR, VT_BSTR_BLOB
            On Error GoTo Str_ErrCatch
            Dim SCur As Currency
            SCur = (CCur(Expression) * 0.0001)
            Call memcpy(CInt64, SCur, VLEN_CURRENCY)
            GoTo Continue_Ex
Str_ErrCatch: throw Exps.InvalidCastException
Continue_Ex:
        Case VT_I1, VT_I2, VT_I4, VT_I8, VT_UI1, VT_UI2, VT_UI4, VT_UI8, VT_UINT, VT_UINT_PTR, VT_INT, VT_INT_PTR
            CInt64.LowerValue = CLng(Expression)
        Case VT_DATE
            Call memcpy(CInt64, Expression, VLEN_DATE)
        Case VT_DECIMAL
            Dim CYDec As Currency
            CYDec = (CCur(Expression) * 0.0001@)
            Call memcpy(CInt64, CYDec, VLEN_CURRENCY)
        Case VT_R4
            Dim SSng As Single: SSng = CSng(Expression)
            CInt64.LowerValue = Fix(SSng)
            CInt64.HigherValue = mMath.GetFloatingPoint(SSng)
        Case VT_R8
            Dim SDbl As Double: SDbl = CDbl(Expression)
            CInt64.LowerValue = Fix(SDbl)
            CInt64.HigherValue = mMath.GetFloatingPoint(SDbl)
        Case VT_SAFEARRAY
            'If API_SafeArrayGetDim(GetArrayPointer(Expression, True)) > 1 Then throw Exps.MultiDimentionException
            If ArraySize(Expression) > 8 Then throw Exps.IndexOutOfRangeException

        Case VT_BOOL
            If Expression Then CInt64.LowerValue = -1: CInt64.HigherValue = -1
        Case VT_DISPATCH
            'If Not TypeOf Expression Is BigNumber Then throw Exps.InvalidArgumentTypeException
            'Dim BN As BigNumber
            'Set BN = Expression

    End Select
End Function
Public Function CInt64ToCur(IntValue As Int64) As Currency
    Call memcpy(ByVal VarPtr(CInt64ToCur), ByVal VarPtr(IntValue), VLEN_CURRENCY)
    CInt64ToCur = (CInt64ToCur * 10000@)
End Function
Public Function CInt64ToDec(IntValue As Int64) As Variant
    Dim Val As Currency
    Call memcpy(ByVal VarPtr(Val), ByVal VarPtr(IntValue), VLEN_CURRENCY)
    CInt64ToDec = CDec(Val * 10000@)
End Function
Public Function CInt64ToString(Value As Int64) As String
    
End Function

Public Function ToString(Expression, Optional ByVal IsTry As Boolean, Optional ByVal DefaultValue As Byte) As String
    On Error GoTo CatchErr
    ToString = CStr(Expression): Exit Function
CatchErr:
    If IsTry Then
        ToString = DefaultValue
    Else
        throw Exps.InvalidCastException
    End If
End Function

Public Function ToByte(Expression, Optional ByVal IsTry As Boolean, Optional ByVal DefaultValue As Byte) As Byte
    On Error GoTo CatchErr
    ToByte = CByte(ToLong(Expression)): Exit Function
CatchErr:
    If IsTry Then
        ToByte = DefaultValue
    Else
        throw Exps.InvalidCastException
    End If
End Function
Public Function ToInteger(Expression, Optional ByVal IsTry As Boolean, Optional ByVal DefaultValue As Integer) As Integer
    On Error GoTo CatchErr
    ToInteger = CInt(ToLong(Expression)): Exit Function
CatchErr:
    If IsTry Then
        ToInteger = DefaultValue
    Else
        throw Exps.InvalidCastException
    End If
End Function
Public Function ToLong(Expression, Optional ByVal IsTry As Boolean, Optional ByVal DefaultValue As Long) As Long
    On Error GoTo CatchErr
    ToLong = CLng(Expression): Exit Function
CatchErr:
    If IsTry Then
        ToLong = DefaultValue
    Else
        throw Exps.InvalidCastException
    End If
End Function
Public Function ToInt64(Expression, Optional ByVal IsTry As Boolean, Optional ByVal DefaultValue As Currency) As Int64
    On Error GoTo CatchErr
    ToInt64 = CInt64(Expression): Exit Function
CatchErr:
    If IsTry Then
        ToInt64 = CInt64(DefaultValue)
    Else
        throw Exps.InvalidCastException
    End If
End Function
Public Function ToSingle(Expression, Optional ByVal IsTry As Boolean, Optional ByVal DefaultValue As Single) As Single
    On Error GoTo CatchErr
    ToSingle = CSng(Expression): Exit Function
CatchErr:
    If IsTry Then
        ToSingle = DefaultValue
    Else
        throw Exps.InvalidCastException
    End If
End Function
Public Function ToDouble(Expression, Optional ByVal IsTry As Boolean, Optional ByVal DefaultValue As Double) As Double
    On Error GoTo CatchErr
    ToDouble = CDbl(Expression): Exit Function
CatchErr:
    If IsTry Then
        ToDouble = DefaultValue
    Else
        throw Exps.InvalidCastException
    End If
End Function
Public Function ToCurrency(Expression, Optional ByVal IsTry As Boolean, Optional ByVal DefaultValue As Currency) As Currency
    On Error GoTo CatchErr
    ToCurrency = CCur(Expression): Exit Function
CatchErr:
    If IsTry Then
        ToCurrency = DefaultValue
    Else
        throw Exps.InvalidCastException
    End If
End Function
Public Function ToDate(Expression, Optional ByVal IsTry As Boolean, Optional ByVal DefaultValue As Date) As Date
    On Error GoTo CatchErr
    ToDate = CDate(Expression): Exit Function
CatchErr:
    If IsTry Then
        ToDate = DefaultValue
    Else
        throw Exps.InvalidCastException
    End If
End Function
Public Function ToDecimal(Expression, Optional ByVal IsTry As Boolean, Optional ByVal DefaultValue As Variant) As Variant
    On Error GoTo CatchErr
    ToDecimal = CDec(Expression): Exit Function
CatchErr:
    If IsTry Then
        ToDecimal = CDec(DefaultValue)
    Else
        throw Exps.InvalidCastException
    End If
End Function


Public Function TryCByte(Expression, ByRef outValue As Byte) As Boolean
    On Error GoTo CatchErr
    outValue = CByte(Expression): TryCByte = True: Exit Function
CatchErr: TryCByte = False
End Function
Public Function TryCInt(Expression, ByRef outValue As Integer) As Boolean
    On Error GoTo CatchErr
    outValue = CInt(Expression): TryCInt = True: Exit Function
CatchErr: TryCInt = False
End Function
Public Function TryCLng(Expression, ByRef outValue As Long) As Boolean
    On Error GoTo CatchErr
    outValue = CByte(Expression): TryCLng = True: Exit Function
CatchErr: TryCLng = False
End Function
Public Function TryCInt64(Expression, ByRef outValue As Int64) As Boolean
    On Error GoTo CatchErr
    outValue = CInt64(Expression): TryCInt64 = True: Exit Function
CatchErr: TryCInt64 = False
End Function
Public Function TryCSng(Expression, ByRef outValue As Single) As Boolean
    On Error GoTo CatchErr
    outValue = CSng(Expression): TryCSng = True: Exit Function
CatchErr: TryCSng = False
End Function
Public Function TryCDbl(Expression, ByRef outValue As Double) As Boolean
    On Error GoTo CatchErr
    outValue = CDbl(Expression): TryCDbl = True: Exit Function
CatchErr: TryCDbl = False
End Function
Public Function TryCCur(Expression, ByRef outValue As Currency) As Boolean
    On Error GoTo CatchErr
    outValue = CCur(Expression): TryCCur = True: Exit Function
CatchErr: TryCCur = False
End Function
Public Function TryCDate(Expression, ByRef outValue As Date) As Boolean
    On Error GoTo CatchErr
    outValue = CDate(Expression): TryCDate = True: Exit Function
CatchErr: TryCDate = False
End Function
Public Function TryCDec(Expression, ByRef outValue As Variant) As Boolean
    On Error GoTo CatchErr
    outValue = CDec(Expression): TryCDec = True: Exit Function
CatchErr: TryCDec = False
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
