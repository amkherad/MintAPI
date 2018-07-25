VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "StringParser is used to provide basic string methods ,and also provides MintAPI string parser engine."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Argument"


'Ideas from Me , Qt , .net Framework , M-Intelligence

Option Base 0
Option Explicit
Const CLASSID As String = "StringParser"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


'Implements ITargetStream
Implements ICloneable
'Implements IAliasable

Public Enum StringDirections
    RightToLeft
    LeftToRight
End Enum
Public Enum StringTypes
    stNull = 0
    stNumeric = 1
    stAlphabetic = 2
    stSpace = 3
    stNewLine = 4
    stSpecialCharacters = 5
    stAll = 6
End Enum

Dim Str As String
Dim StrLen As Long
Dim isSpecial As Boolean
Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String

Friend Sub Initialize(Optional Str As String, Optional Special As Boolean = False)
    If Not IsMissing(Str) Then Call SetString(Str)
    isSpecial = Special
End Sub

Public Sub SetString(Value As String)
    Str = Value
    StrLen = Len(Str)
End Sub
Public Property Get Value() As String
    Value = Str
End Property
Public Property Get IsSpecialCharacters() As Boolean
    IsSpecialCharacters = isSpecial
End Property

Public Property Get Length() As Long
    Length = StrLen
End Property
Public Function Duplicate() As StringParser
    Dim sp As New StringParser
    Call sp.Initialize(Str, isSpecial)
    Set Duplicate = sp
End Function



Public Function Argument(ParamArray Args() As Variant) As StringParser
    Dim sp As New StringParser
    Dim cArgs() As Variant
    cArgs = Args
    Call sp.Initialize(funcwArgs(Str, cArgs), True)
    Set Argument = sp
End Function
Public Function A(ParamArray Args() As Variant) As String
    Dim cArgs() As Variant
    cArgs = Args
    A = funcwArgs(Str, cArgs)
End Function

Public Function SubString(StartIndex As Long, Optional Length As Long = -1) As String
    If Length = -1 Then
        SubString = Mid(Str, StartIndex + 1)
    Else
        SubString = Mid(Str, StartIndex + 1, Length)
    End If
End Function
Public Function Replace(strWhat As String, strWith As String, Optional StartIndex As Long = 0, Optional Count As Long = -1) As String
    Replace = VBA.Replace(Str, strWhat, strWith, StartIndex + 1, Count)
End Function
Public Function IndexOf(strWhat As String, Optional StartIndex As Long = 0) As Long
    IndexOf = InStr(StartIndex + 1, Str, strWhat)
End Function
Public Function LastIndexOf(strWhat As String, Optional StartIndex As Long = -1) As Long
    LastIndexOf = InStrRev(Str, strWhat, IIf(StartIndex < 0, StartIndex, StartIndex + 1))
End Function
Public Function HowManyOf(strWhat As String, Optional StartIndex As Long = 0) As Long
    Dim i As Long, c As Long, hLen As Long
    hLen = Len(strWhat)
    For i = StartIndex + 1 To Len(Str) - hLen
        If Mid(Str, i, hLen) = strWhat Then
            c = c + 1
        End If
    Next
    HowManyOf = c
End Function
Public Function Trim() As String
    Trim = VBA.Trim$(Str)
End Function
Public Function TrimLeft() As String
    TrimLeft = VBA.LTrim$(Str)
End Function
Public Function TrimRight() As String
    TrimRight = VBA.RTrim$(Str)
End Function
Public Function TrimArr(TrimStrs() As String) As String
    If LenB(Str) = 0 Then Exit Function
    If ArraySize(TrimStrs) <= 0 Then TrimArr = Str: Exit Function
    Dim isExt As Boolean
    Dim strBuff As String, strBuff2 As String
    Dim i As Long, j As Long
    Dim ln As Long, Str1 As String, StrLen As Integer, lIndex As Long
    ln = Len(Str)
    i = 1
    While i <= ln
ContinueWhile:
    isExt = False
        For j = 0 To UBound(TrimStrs)
            Str1 = CStr(TrimStrs(j))
            StrLen = Len(Str1)
            If i + StrLen - 1 <= ln Then
                Dim bff As String
                bff = Mid(Str, i, StrLen)
                If bff = Str1 Then
                    isExt = True
                    lIndex = i
                    i = i + StrLen
                    GoTo ContinueWhile
                End If
            End If
        Next
        If Not isExt Then
            strBuff = Mid(Str, i, ln - i + 1)
            GoTo ExitFunc1
        End If
        i = i + 1
    Wend
    strBuff = vbNullString
ExitFunc1:
    ln = Len(strBuff)
    lIndex = ln
    i = ln
    While (i > 0)
ContinueWhile2:
    isExt = False
        For j = 0 To UBound(TrimStrs)
            Str1 = CStr(TrimStrs(j))
            StrLen = Len(Str1)
            If i - StrLen > 0 Then
            'If (i + strLen - 1) <= ln Then
                If Mid(strBuff, i - StrLen + 1, StrLen) = Str1 Then
                    isExt = True
                    lIndex = (i - StrLen)
                    i = i - StrLen
                    GoTo ContinueWhile2
                End If
            End If
        Next
        If Not isExt Then
            strBuff2 = Mid(strBuff, 1, lIndex)
            GoTo ExitFunc2
        End If
        i = i - 1
    Wend
    strBuff2 = vbNullString
ExitFunc2:
    TrimArr = strBuff2
End Function
Public Function TrimLeftArr(TrimStrs() As String) As String
    If LenB(Str) = 0 Then Exit Function
    If ArraySize(TrimStrs) <= 0 Then TrimLeftArr = Str: Exit Function
    Dim isExt As Boolean
    Dim i As Long, j As Long
    Dim ln As Long, Str1 As String, StrLen As Integer, lIndex As Long
    ln = Len(Str)
    i = 1
    While i <= ln
ContinueWhile:
    isExt = False
        For j = 0 To UBound(TrimStrs)
            Str1 = CStr(TrimStrs(j))
            StrLen = Len(Str1)
            If i + StrLen - 1 <= ln Then
                Dim bff As String
                bff = Mid(Str, i, StrLen)
                If bff = Str1 Then
                    isExt = True
                    lIndex = i
                    i = i + StrLen
                    GoTo ContinueWhile
                End If
            End If
        Next
        If Not isExt Then
            TrimLeftArr = Mid(Str, i, ln - i + 1)
            Exit Function
        End If
        i = i + 1
    Wend
    TrimLeftArr = vbNullString
End Function
Public Function TrimRightArr(TrimStrs() As String) As String
    If LenB(Str) = 0 Then Exit Function
    If ArraySize(TrimStrs) <= 0 Then TrimRightArr = Str: Exit Function
    Dim isExt As Boolean
    Dim i As Long, j As Long
    Dim ln As Long, Str1 As String, StrLen As Integer, lIndex As Long
    ln = Len(Str)
    lIndex = ln
    i = ln
    While (i > 0)
ContinueWhile:
    isExt = False
        For j = 0 To UBound(TrimStrs)
            Str1 = CStr(TrimStrs(j))
            StrLen = Len(Str1)
            If i - StrLen > 0 Then
            'If (i + strLen - 1) <= ln Then
                If Mid(Str, i - StrLen + 1, StrLen) = Str1 Then
                    isExt = True
                    lIndex = (i - StrLen)
                    i = i - StrLen
                    GoTo ContinueWhile
                End If
            End If
        Next
        If Not isExt Then
            TrimRightArr = Mid(Str, 1, lIndex)
            Exit Function
        End If
        i = i - 1
    Wend
    TrimRightArr = vbNullString
End Function

Private Function IsWhiteSpace(Str As String) As Boolean
    Select Case Str
        Case " ", vbTab, vbCr, vbLf, vbCrLf, " " 'alt+255
            IsWhiteSpace = True
    End Select
End Function
Public Function IsNullOrEmpty(Optional OptionalString) As Boolean
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    IsNullOrEmpty = (LenB(VL) = 0) Or (VL = Empty)
End Function
Public Function IsNullOrWhiteSpace(Optional OptionalString) As Boolean
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    IsNullOrWhiteSpace = IsNullOrEmpty(VL)
    If IsNullOrWhiteSpace Then Exit Function
    Dim l As Long
    l = Len(VL)
    Dim i As Long
    For i = 0 To l
        If Not IsWhiteSpace(Mid(VL, i, 1)) Then
            IsNullOrWhiteSpace = False
            Exit Function
        End If
    Next
    IsNullOrWhiteSpace = True
End Function

Public Function Join(TargetArray, Optional Delimiter) As String
    Join = VBA.Join(TargetArray, Delimiter)
End Function
Public Function fromStringArray(TargetArray, Optional Delimiter) As String
    fromStringArray = VBA.Join(TargetArray, Delimiter)
End Function
Public Function Split(Optional Delimiter, Optional Limit As Long = -1) As String()
    Split = VBA.Split(Str, Delimiter, Limit)
End Function
Public Function SplitArr(Optional Delimiters, Optional Limit As Long = -1) As String()
    Dim Dels() As String, i As Long, Lb As Long, DL
    Dim DelsCount As Long
    If IsArray(Delimiters) Then
        DelsCount = ArraySize(Delimiters)
        ReDim Dels(DelsCount - 1)
        Lb = LBound(Delimiters)
        For i = 0 To DelsCount - 1
            Dels(i) = CStr(Delimiters(i + Lb))
        Next
    Else
        ReDim Dels(0)
        DelsCount = 1
        Dels(0) = CStr(Delimiters)
    End If
    
    Dim RetVal() As String, CuStr As String, execOnce As Boolean
    Dim retValCount As Long
    
    For i = 1 To StrLen
        execOnce = False
        For Each DL In Dels
            If Mid(Str, i, Len(DL)) = DL Then
                ReDim Preserve RetVal(retValCount)
                RetVal(retValCount) = CStr(CuStr)
                retValCount = retValCount + 1
                CuStr = vbNullString
                execOnce = True
            End If
        Next
        If Not execOnce Then
            CuStr = CuStr & Mid(Str, i, 1)
        End If
    Next
    If LenB(CuStr) <> 0 Then
        ReDim Preserve RetVal(retValCount)
        RetVal(retValCount) = CStr(CuStr)
        retValCount = retValCount + 1
    End If
    
    SplitArr = RetVal
End Function
Public Function toStringArray(Optional Delimiter, Optional Limit As Long = -1) As String()
    toStringArray = VBA.Split(Str, Delimiter, Limit)
End Function
Public Function Concat(strConcat As String, Optional Separator As String = vbNullString) As String
    Concat = Str & Separator & strConcat
End Function
Public Function NewLine() As String
    NewLine = vbCrLf
End Function
Public Function Format(strFormat As String) As String
    Format = VBA.Format$(Str, strFormat)
End Function
Public Function Compare(strWhat As String) As Boolean
    Compare = (Str Like strWhat)
End Function
Public Function Remove(StartIndex As Long, Optional Length As Long = -1) As String
    Dim ln As Long
    ln = Len(Str)
    If Length = -1 Then
        Remove = Mid(Str, StartIndex + 1)
    Else
        Remove = Mid(Str, 1, StartIndex)
        Remove = Remove & Mid(Str, StartIndex + 1 + Length)
    End If
End Function
Public Function Reverse(Optional Length As Long = -1) As String
    If Length = -1 Then
        Reverse = StrReverse(Str)
    Else
        Reverse = StrReverse(Left(Length, Length))
    End If
End Function
Public Function Insert(strInsert As String, InsertIndex As Long, Optional Length As Long = -1) As String
    Dim fP As String, sp As String, ln As Long, outLen As Long
    fP = Left(Str, InsertIndex + 1)
    sp = Mid(Str, InsertIndex + 2)
    Insert = fP & strInsert & sp
    ln = Len(Insert)
    If InsertIndex <> -1 Then
        outLen = Length
        If outLen = -1 Then outLen = ln
        If outLen > ln Then
            Insert = StringParser(Insert).PadRight(outLen, " ")
        Else
            Insert = Mid(Insert, 1, outLen)
        End If
    End If
End Function
Public Function PadLeft(TotalLength As Long, FillWith As String) As String
    If TotalLength <= 0 Then Exit Function
    Dim ln As Long
    ln = Len(Str)
    If ln < TotalLength Then
        PadLeft = String(TotalLength - ln, FillWith) & Str
        PadLeft = Mid(PadLeft, 1, TotalLength)
    ElseIf ln > TotalLength Then
        PadLeft = Mid(Str, ln - TotalLength + 1)
    Else
        PadLeft = Str
    End If
End Function
Public Function PadRight(TotalLength As Long, FillWith As String) As String
    Dim ln As Long
    ln = Len(Str)
    If ln < TotalLength Then
        PadRight = Str & String(TotalLength - ln, FillWith)
        PadRight = Mid(PadRight, 1, TotalLength)
    ElseIf ln > TotalLength Then
        PadRight = Mid(Str, 1, TotalLength)
    Else
        PadRight = Str
    End If
End Function
Public Function Contains(strWhat As String, Optional StartIndex As Long = 0) As Boolean
    If LenB(Str) = 0 Then Exit Function
    Contains = (InStr(StartIndex + 1, Str, strWhat) > 0)
End Function
Public Function StartsWith(strWhat As String) As Boolean
    StartsWith = (Left(Str, Len(strWhat)) = strWhat)
End Function
Public Function EndsWith(strWhat As String) As Boolean
    EndsWith = (Right(Str, Len(strWhat)) = strWhat)
End Function

Public Function ToString() As String
    ToString = Str
End Function

Public Function ToInteger(Optional OptionalString) As Integer
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    ToInteger = CInt(VL)
End Function
Public Function ToLong(Optional OptionalString) As Long
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    ToLong = CLng(VL)
End Function
Public Function ToDouble(Optional OptionalString) As Double
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    ToDouble = CDbl(VL)
End Function
Public Function ToSingle(Optional OptionalString) As Single
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    ToSingle = CSng(VL)
End Function
Public Function ToDate(Optional OptionalString) As Date
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    ToDate = CDate(VL)
End Function
Public Function ToCurrency(Optional OptionalString) As Currency
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    ToCurrency = CCur(VL)
End Function
Public Function ToByte(Optional OptionalString) As Byte
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    ToByte = CByte(VL)
End Function
Public Function ToArray(Optional OptionalString) As Variant()
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))

End Function
Public Function ToHex(Optional OptionalString) As String
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    ToHex = Hex$(VL)
End Function
Public Function ToOctal(Optional OptionalString) As String
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    ToOctal = Oct$(VL)
End Function
Public Function ToBinary(Optional OptionalString) As String ' binary number 1010001011101
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))

End Function
Public Function ToByteArray(Optional OptionalString) As Byte()
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    ToByteArray = StringToByteArray(VL)
End Function
Public Function FromHex(ByVal Value As String) As StringParser
    If UCase(Left(Value, 2)) = "&H" Then
        Set FromHex = StringParser(CLng(Value))
    ElseIf UCase(Left(Value, 2)) = "0X" Then
        Set FromHex = StringParser(CLng("&H" & Mid(Value, 2)))
    Else
        Set FromHex = StringParser(CLng("&H" & Value))
    End If
End Function
Public Function FromOctal(ByVal Value As String) As StringParser
    If UCase(Left(Value, 2)) = "&O" Then
        Set FromOctal = StringParser(CLng(Value))
    ElseIf UCase(Left(Value, 1)) = "0" Then
        Set FromOctal = StringParser(CLng("&O" & Mid(Value, 2)))
    Else
        Set FromOctal = StringParser(CLng("&O" & Value))
    End If
End Function
Public Function FromBinary(ByVal Value As String) As StringParser ' binary number 1010001011101
    
End Function

Public Function ValidatePath(Optional CheckForDrive As Boolean = False, Optional OptionalString) As Boolean
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    ValidatePath = CheckPathValidation(VL, False, CheckForDrive)
End Function

Public Function toAbsolutePath(Optional CheckForDrive As Boolean = True, Optional OptionalString) As String
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), Str, CStr(OptionalString))
    If Not CheckPathValidation(VL, True, CheckForDrive) Then throw Exps.InvalidPathException
    toAbsolutePath = VL
End Function


Public Function GetRandomAllString(ByVal Length As Long) As String
If Length <= 0 Then throw Exps.ZeroNegativeArgumentException("Length Must Be Greater Than Zero."), Me  ' check if length less than 0.
    Dim X As String, i As Long
    Call Randomize
    For i = 1 To Length
        X = X & Chr(CLng(Rnd * 255))
    Next
    GetRandomAllString = X
End Function
Public Function GetRandomString(ByVal Length As Long, ByVal stringType As StringTypes) As String
If Length <= 0 Then throw Exps.ZeroNegativeArgumentException("Length Must Be Greater Than Zero."), Me  ' check if length less than 0. ' check if length less than 0.
End Function
Public Function GetRandomAlphabeticString(ByVal Length As Long) As String
If Length <= 0 Then throw Exps.ZeroNegativeArgumentException("Length Must Be Greater Than Zero."), Me   ' check if length less than 0.
    Dim X As String, i As Long, B As Byte
    Call Randomize
    For i = 1 To Length
        B = CLng(Rnd * 50)
        If B < 25 Then
            B = B + 65
        Else
            B = (B - 25) + 97
        End If
        X = X & Chr(B)
    Next
    GetRandomAlphabeticString = X
End Function
Public Function GetRandomNumericString(ByVal Length As Long) As String
If Length <= 0 Then throw Exps.ZeroNegativeArgumentException("Length Must Be Greater Than Zero."), Me  ' check if length less than 0.
    Dim Str As String, i As Long
    Call Randomize
    For i = 1 To Length
        Str = Str & CStr(Fix(Rnd * 10))
    Next
    GetRandomNumericString = Str
End Function

' pattern : pattern of exported string. * = single character, ? = single alphabetic character , # = single numeric character , ! = space character
' Length : length of exported string.
' addto : add extra string to?
' additionalStringType : extra added strin type.
Public Function GetRandomStringByTemplate(ByVal Pattern As String, ByVal Length As Long, Optional ByVal addto As StringDirections = LeftToRight, Optional ByVal additionalStringType As StringTypes = stAll) As String
If Length <= 0 Then throw Exps.ZeroNegativeArgumentException("Length Must Be Greater Than Zero."), Me   ' check if length less than 0.
    Dim Str As String, StrLen As Long, i As Long, oStrLength As Long 'Other Strings Length
    StrLen = Len(Pattern)
    Dim Char As String * 1, B As Long
    oStrLength = Length - StrLen
    Call Randomize
    For i = 1 To StrLen
        Char = Mid(Pattern, i, 1)
        Select Case Char
            Case "*"
                Str = Str & Chr(CLng(Rnd * 255))
            Case "?"
                B = Fix(Rnd * 50)
                If B < 25 Then
                    B = B + 65
                Else
                    B = (B - 25) + 97
                End If
                Str = Str & Chr(B)
            Case "#"
                Str = Str & CStr(CByte(Rnd * 10))
            Case "!"
                B = Round(Rnd * 3)
                If B = 0 Then
                    Str = Str & " "
                ElseIf B = 1 Then
                    Str = Str & vbTab
                ElseIf B = 2 Then
                    Str = Str & vbCrLf
                Else
                    Str = Str & vbNewLine
                End If
            Case Else
                Str = Str & Char
        End Select
    Next
    If oStrLength > 0 Then
        If addto = LeftToRight Then
            GetRandomStringByTemplate = GetRandomString(oStrLength, additionalStringType) & Str
        Else
            GetRandomStringByTemplate = Str & GetRandomString(oStrLength, additionalStringType)
        End If
    Else
        GetRandomStringByTemplate = Str
    End If
End Function

Private Function ICloneable_Clone() As Object
    Dim sp As New StringParser
    Call sp.Initialize(Str, isSpecial)
    Set ICloneable_Clone = sp
End Function

Private Function ITargetStream_Available(Dir As StreamDirection) As Long
    Dim Val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        Val = ((Len(Str) > 0) And inState)
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        Val = Val Or (outState)
    End If
    ITargetStream_Available = Val
End Function

'<NOT CLOSABLE>
Private Sub ITargetStream_CloseStream(Optional Direction As StreamDirection = 3&)
    If (Direction And sdInStream) = sdInStream Then
        inState = False
    End If
    If (Direction And sdOutStream) = sdOutStream Then
        outState = False
    End If
End Sub

Private Sub ITargetStream_OpenStream(Optional Direction As StreamDirection = StreamDirection.sdBoth)
    If (Direction And sdInStream) = sdInStream Then
        inState = True
    End If
    If (Direction And sdOutStream) = sdOutStream Then
        outState = True
    End If
End Sub
Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean
    Dim Val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        Val = inState
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        Val = Val Or outState
    End If
    ITargetStream_getState = Val
End Function

Private Sub ITargetStream_inStream(Data, Optional Length As Long = -1)
    If Not inState Then throw Exps.InvalidStatusException
    'Call mint_put_byte_array_to(Data, d, Length)
End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    ITargetStream_readBytes = StringToByteArray(Mid(Str, IIf(StartAt = -1, 1, StartAt + 1), Length))
End Function
Private Sub ITargetStream_outStream(Data, Optional Length As Long = -1)
    If Not outState Then throw Exps.InvalidStatusException
    If VarType(Data) = VBObject Then
        If TypeOf Data Is ObjectBuffer Then
            Dim objBuffer As ObjectBuffer
            Set objBuffer = Data
            If objBuffer.IsStreamingValue Then
                Call objBuffer.thrownotargs(1)
                If objBuffer.ArgType(0) = vbString Then
                    Dim strFirst As String
                    strFirst = objBuffer.First
                    If strFirst = OBJECTBUFFER_VALUES_STREAMING_CLEAR Then
                        Str = ""
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_NEWLINE Then
                        Str = Str & vbCrLf
                        Exit Sub
                    Else
                        throw Exps.InvalidArgumentTypeException("Streaming Command '" & strFirst & "' Is Not Available For ByteArray.")
                    End If
                Else
                    throw Exps.InvalidArgumentTypeException("Streaming Type Is Not Available For ByteArray.")
                End If
            End If
        End If
    End If
    Str = Str & CStr(Data)
End Sub

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, inState, inState_LOCK, Reserved)
End Sub
Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, outState, outState_LOCK, Reserved)
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
