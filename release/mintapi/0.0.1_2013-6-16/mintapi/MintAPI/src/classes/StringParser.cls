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
Option Base 0
Option Explicit
Const CLASSID As String = "StringParser"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements ITargetStream
Implements IClassClonable
Implements IClassTexer
Implements IAliasable

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

Dim str As String
Dim strLen As Long
Dim isSpecial As Boolean
Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String

Friend Sub Initialize(Optional str As String, Optional Special As Boolean = False)
    If Not IsMissing(str) Then Call SetString(str)
    isSpecial = Special
End Sub

Public Sub SetString(Value As String)
    str = Value
    strLen = Len(str)
End Sub
Public Property Get Value() As String
    Value = str
End Property
Public Property Get IsSpecialCharacters() As Boolean
    IsSpecialCharacters = isSpecial
End Property

Public Property Get Length() As Long
    Length = strLen
End Property
Public Function Duplicate() As StringParser
    Dim sp As New StringParser
    Call sp.Initialize(str, isSpecial)
    Set Duplicate = sp
End Function



Public Function Argument(ParamArray Args() As Variant) As StringParser
    Dim sp As New StringParser
    Dim cArgs() As Variant
    cArgs = Args
    Call sp.Initialize(funcwArgs(str, cArgs), True)
    Set Argument = sp
End Function
Public Function A(ParamArray Args() As Variant) As String
    Dim cArgs() As Variant
    cArgs = Args
    A = funcwArgs(str, cArgs)
End Function

Public Function SubString(StartIndex As Long, Optional Length As Long = -1) As String
    If Length = -1 Then
        SubString = mID(str, StartIndex + 1)
    Else
        SubString = mID(str, StartIndex + 1, Length)
    End If
End Function
Public Function Replace(strWhat As String, strWith As String, Optional StartIndex As Long = 0, Optional Count As Long = -1) As String
    Replace = VBA.Replace(str, strWhat, strWith, StartIndex + 1, Count)
End Function
Public Function IndexOf(strWhat As String, Optional StartIndex As Long = 0) As Long
    IndexOf = InStr(StartIndex + 1, str, strWhat)
End Function
Public Function LastIndexOf(strWhat As String, Optional StartIndex As Long = -1) As Long
    LastIndexOf = InStrRev(str, strWhat, IIf(StartIndex < 0, StartIndex, StartIndex + 1))
End Function
Public Function HowManyOf(strWhat As String, Optional StartIndex As Long = 0) As Long
    Dim i As Long, C As Long, hLen As Long
    hLen = Len(strWhat)
    For i = StartIndex + 1 To Len(str) - hLen
        If mID(str, i, hLen) = strWhat Then
            C = C + 1
        End If
    Next
    HowManyOf = C
End Function
Public Function Trim() As String
    Trim = VBA.Trim$(str)
End Function
Public Function TrimLeft() As String
    TrimLeft = VBA.LTrim$(str)
End Function
Public Function TrimRight() As String
    TrimRight = VBA.RTrim$(str)
End Function
Public Function TrimArr(trimStrs() As String) As String
    If str = "" Then Exit Function
    If ArraySize(trimStrs) <= 0 Then TrimArr = str: Exit Function
    Dim isExt As Boolean
    Dim strBuff As String, strBuff2 As String
    Dim i As Long, j As Long
    Dim ln As Long, str1 As String, strLen As Integer, lIndex As Long
    ln = Len(str)
    i = 1
    While i <= ln
ContinueWhile:
    isExt = False
        For j = 0 To UBound(trimStrs)
            str1 = CStr(trimStrs(j))
            strLen = Len(str1)
            If i + strLen - 1 <= ln Then
                Dim bff As String
                bff = mID(str, i, strLen)
                If bff = str1 Then
                    isExt = True
                    lIndex = i
                    i = i + strLen
                    GoTo ContinueWhile
                End If
            End If
        Next
        If Not isExt Then
            strBuff = mID(str, i, ln - i + 1)
            GoTo ExitFunc1
        End If
        i = i + 1
    Wend
    strBuff = ""
ExitFunc1:
    ln = Len(strBuff)
    lIndex = ln
    i = ln
    While (i > 0)
ContinueWhile2:
    isExt = False
        For j = 0 To UBound(trimStrs)
            str1 = CStr(trimStrs(j))
            strLen = Len(str1)
            If i - strLen > 0 Then
            'If (i + strLen - 1) <= ln Then
                If mID(strBuff, i - strLen + 1, strLen) = str1 Then
                    isExt = True
                    lIndex = (i - strLen)
                    i = i - strLen
                    GoTo ContinueWhile2
                End If
            End If
        Next
        If Not isExt Then
            strBuff2 = mID(strBuff, 1, lIndex)
            GoTo ExitFunc2
        End If
        i = i - 1
    Wend
    strBuff2 = ""
ExitFunc2:
    TrimArr = strBuff2
End Function
Public Function TrimLeftArr(trimStrs() As String) As String
    If str = "" Then Exit Function
    If ArraySize(trimStrs) <= 0 Then TrimLeftArr = str: Exit Function
    Dim isExt As Boolean
    Dim i As Long, j As Long
    Dim ln As Long, str1 As String, strLen As Integer, lIndex As Long
    ln = Len(str)
    i = 1
    While i <= ln
ContinueWhile:
    isExt = False
        For j = 0 To UBound(trimStrs)
            str1 = CStr(trimStrs(j))
            strLen = Len(str1)
            If i + strLen - 1 <= ln Then
                Dim bff As String
                bff = mID(str, i, strLen)
                If bff = str1 Then
                    isExt = True
                    lIndex = i
                    i = i + strLen
                    GoTo ContinueWhile
                End If
            End If
        Next
        If Not isExt Then
            TrimLeftArr = mID(str, i, ln - i + 1)
            Exit Function
        End If
        i = i + 1
    Wend
    TrimLeftArr = ""
End Function
Public Function TrimRightArr(trimStrs() As String) As String
    If str = "" Then Exit Function
    If ArraySize(trimStrs) <= 0 Then TrimRightArr = str: Exit Function
    Dim isExt As Boolean
    Dim i As Long, j As Long
    Dim ln As Long, str1 As String, strLen As Integer, lIndex As Long
    ln = Len(str)
    lIndex = ln
    i = ln
    While (i > 0)
ContinueWhile:
    isExt = False
        For j = 0 To UBound(trimStrs)
            str1 = CStr(trimStrs(j))
            strLen = Len(str1)
            If i - strLen > 0 Then
            'If (i + strLen - 1) <= ln Then
                If mID(str, i - strLen + 1, strLen) = str1 Then
                    isExt = True
                    lIndex = (i - strLen)
                    i = i - strLen
                    GoTo ContinueWhile
                End If
            End If
        Next
        If Not isExt Then
            TrimRightArr = mID(str, 1, lIndex)
            Exit Function
        End If
        i = i - 1
    Wend
    TrimRightArr = ""
End Function

Private Function IsWhiteSpace(str As String) As Boolean
    Select Case str
        Case " ", vbTab, vbCr, vbLf, vbCrLf, " " 'alt+255
            IsWhiteSpace = True
    End Select
End Function
Public Function IsNullOrEmpty(Optional OptionalString) As Boolean
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    IsNullOrEmpty = (VL = "") Or (VL = Empty)
End Function
Public Function IsNullOrWhiteSpace(Optional OptionalString) As Boolean
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    IsNullOrWhiteSpace = IsNullOrEmpty(VL)
    If IsNullOrWhiteSpace Then Exit Function
    Dim l As Long
    l = Len(VL)
    Dim i As Long
    For i = 0 To l
        If Not IsWhiteSpace(mID(VL, i, 1)) Then
            IsNullOrWhiteSpace = False
            Exit Function
        End If
    Next
    IsNullOrWhiteSpace = True
End Function

Public Function Join(targetArray, Optional Delimiter) As String
    Join = VBA.Join(targetArray, Delimiter)
End Function
Public Function fromStringArray(targetArray, Optional Delimiter) As String
    fromStringArray = VBA.Join(targetArray, Delimiter)
End Function
Public Function Split(Optional Delimiter, Optional Limit As Long = -1) As String()
    Split = VBA.Split(str, Delimiter, Limit)
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
    
    Dim retVal() As String, CuStr As String, execOnce As Boolean
    Dim retValCount As Long
    
    For i = 1 To strLen
        execOnce = False
        For Each DL In Dels
            If mID(str, i, Len(DL)) = DL Then
                ReDim Preserve retVal(retValCount)
                retVal(retValCount) = CStr(CuStr)
                retValCount = retValCount + 1
                CuStr = ""
                execOnce = True
            End If
        Next
        If Not execOnce Then
            CuStr = CuStr & mID(str, i, 1)
        End If
    Next
    If CuStr <> "" Then
        ReDim Preserve retVal(retValCount)
        retVal(retValCount) = CStr(CuStr)
        retValCount = retValCount + 1
    End If
    
    SplitArr = retVal
End Function
Public Function toStringArray(Optional Delimiter, Optional Limit As Long = -1) As String()
    toStringArray = VBA.Split(str, Delimiter, Limit)
End Function
Public Function Concat(strConcat As String, Optional Separator As String = "") As String
    Concat = str & Separator & strConcat
End Function
Public Function NewLine() As String
    NewLine = vbCrLf
End Function
Public Function Format(strFormat As String) As String
    Format = VBA.Format$(str, strFormat)
End Function
Public Function Compare(strWhat As String) As Boolean
    Compare = (str Like strWhat)
End Function
Public Function Remove(StartIndex As Long, Optional Length As Long = -1) As String
    Dim ln As Long
    ln = Len(str)
    If Length = -1 Then
        Remove = mID(str, StartIndex + 1)
    Else
        Remove = mID(str, 1, StartIndex)
        Remove = Remove & mID(str, StartIndex + 1 + Length)
    End If
End Function
Public Function Reverse(Optional Length As Long = -1) As String
    If Length = -1 Then
        Reverse = StrReverse(str)
    Else
        Reverse = StrReverse(Left(Length, Length))
    End If
End Function
Public Function Insert(strInsert As String, InsertIndex As Long, Optional Length As Long = -1) As String
    Dim fP As String, sp As String, ln As Long, outLen As Long
    fP = Left(str, InsertIndex + 1)
    sp = mID(str, InsertIndex + 2)
    Insert = fP & strInsert & sp
    ln = Len(Insert)
    If InsertIndex <> -1 Then
        outLen = Length
        If outLen = -1 Then outLen = ln
        If outLen > ln Then
            Insert = StringParser(Insert).PadRight(outLen, " ")
        Else
            Insert = mID(Insert, 1, outLen)
        End If
    End If
End Function
Public Function PadLeft(TotalLength As Long, FillWith As String) As String
    If TotalLength <= 0 Then Exit Function
    Dim ln As Long
    ln = Len(str)
    If ln < TotalLength Then
        PadLeft = String(TotalLength - ln, FillWith) & str
        PadLeft = mID(PadLeft, 1, TotalLength)
    ElseIf ln > TotalLength Then
        PadLeft = mID(str, ln - TotalLength + 1)
    Else
        PadLeft = str
    End If
End Function
Public Function PadRight(TotalLength As Long, FillWith As String) As String
    Dim ln As Long
    ln = Len(str)
    If ln < TotalLength Then
        PadRight = str & String(TotalLength - ln, FillWith)
        PadRight = mID(PadRight, 1, TotalLength)
    ElseIf ln > TotalLength Then
        PadRight = mID(str, 1, TotalLength)
    Else
        PadRight = str
    End If
End Function
Public Function Contains(strWhat As String, Optional StartIndex As Long = 0) As Boolean
    If str = "" Then Exit Function
    Contains = (InStr(StartIndex + 1, str, strWhat) > 0)
End Function
Public Function StartsWith(strWhat As String) As Boolean
    StartsWith = (Left(str, Len(strWhat)) = strWhat)
End Function
Public Function EndsWith(strWhat As String) As Boolean
    EndsWith = (Right(str, Len(strWhat)) = strWhat)
End Function

Public Function toLower(Optional OptionalString) As String
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toLower = LCase$(VL)
End Function
Public Function toUpper(Optional OptionalString) As String
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toUpper = UCase$(VL)
End Function
Public Function toString() As String
    toString = str
End Function

Public Function IsNumber() As Boolean
    IsNumber = VBA.IsNumeric(str)
End Function
Public Function IsNumeric() As Boolean
    Dim i As Long, ln As Long, chC As Byte
    ln = Len(str)
    If ln = 0 Then Exit Function
    For i = 1 To Len(str)
        chC = Asc(mID(str, i, 1))
        If (chC < 48) Or (chC > 57) Then
            IsNumeric = False
            Exit Function
        End If
    Next
    IsNumeric = True
End Function
Public Function IsAlphabetic() As Boolean
    Dim i As Long, ln As Long, chC As Byte
    ln = Len(str)
    If ln = 0 Then Exit Function
    For i = 1 To Len(str)
        chC = Asc(mID(str, i, 1))
        If ((chC < 65) Or (chC > 90 And chC < 97) Or (chC > 122)) Then
            IsAlphabetic = False
            Exit Function
        End If
    Next
    IsAlphabetic = True
End Function



Public Function toInteger(Optional OptionalString) As Integer
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toInteger = CInt(VL)
End Function
Public Function toLong(Optional OptionalString) As Long
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toLong = CLng(VL)
End Function
Public Function toDouble(Optional OptionalString) As Double
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toDouble = CDbl(VL)
End Function
Public Function toSingle(Optional OptionalString) As Single
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toSingle = CSng(VL)
End Function
Public Function toDate(Optional OptionalString) As Date
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toDate = CDate(VL)
End Function
Public Function toData(Optional OptionalString) As IData
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    Dim B As New BinaryData

    Set toData = B
End Function
Public Function toCurrency(Optional OptionalString) As Currency
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toCurrency = CCur(VL)
End Function
Public Function toByte(Optional OptionalString) As Byte
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toByte = CByte(VL)
End Function
Public Function toArray(Optional OptionalString) As Variant()
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))

End Function
Public Function toHex(Optional OptionalString) As String
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toHex = Hex$(VL)
End Function
Public Function toOctal(Optional OptionalString) As String
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toOctal = Oct$(VL)
End Function
Public Function toBinary(Optional OptionalString) As String ' binary number 1010001011101
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))

End Function
Public Function toByteArray(Optional OptionalString) As Byte()
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    toByteArray = baseMethods.StringToByteArray(VL)
End Function
Public Function fromHex(ByVal Value As String) As StringParser
    If UCase(Left(Value, 2)) = "&H" Then
        Set fromHex = StringParser(CLng(Value))
    ElseIf UCase(Left(Value, 2)) = "0X" Then
        Set fromHex = StringParser(CLng("&H" & mID(Value, 2)))
    Else
        Set fromHex = StringParser(CLng("&H" & Value))
    End If
End Function
Public Function fromOctal(ByVal Value As String) As StringParser
    If UCase(Left(Value, 2)) = "&O" Then
        Set fromOctal = StringParser(CLng(Value))
    ElseIf UCase(Left(Value, 1)) = "0" Then
        Set fromOctal = StringParser(CLng("&O" & mID(Value, 2)))
    Else
        Set fromOctal = StringParser(CLng("&O" & Value))
    End If
End Function
Public Function fromBinary(ByVal Value As String) As StringParser ' binary number 1010001011101
    
End Function

Public Function ValidatePath(Optional CheckForDrive As Boolean = False, Optional OptionalString) As Boolean
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    ValidatePath = CheckPathValidation(VL, False, CheckForDrive)
End Function

Public Function toAbsolutePath(Optional CheckForDrive As Boolean = True, Optional OptionalString) As String
    Dim VL As String
    VL = IIf(IsMissing(OptionalString), str, CStr(OptionalString))
    If Not CheckPathValidation(VL, True, CheckForDrive) Then throw InvalidPathException
    toAbsolutePath = VL
End Function


Public Function ValidateSocketAddress(SocketAddress As SocketAddress) As Boolean
    
End Function
Public Function ValidateIP(IP As IPAddress) As Boolean
    
End Function
Public Function GetIPVersion(IP As IPAddress) As Long
    
End Function
Public Function ValidatePort(SocketPort As Long) As Boolean
    
End Function


Public Function ParseDirectory(Optional OptionalPath) As Directory
    Dim d As New Directory
    Dim VL As String
    VL = IIf(IsMissing(OptionalPath), str, CStr(OptionalPath))
    Call d.Initialize(VL)
    Set ParseDirectory = d
End Function
Public Function ParseFile(Optional OptionalPath) As File
    Dim f As New File
    Dim VL As String
    VL = IIf(IsMissing(OptionalPath), str, CStr(OptionalPath))
    Call f.Initialize(VL)
    Set ParseFile = f
End Function
Public Function ParseStepString(Optional OptionalPath)
    
End Function
Public Function ParseRegistry(Optional OptionalPath)
    Dim R As New Registry
    Dim VL As String
    VL = IIf(IsMissing(OptionalPath), str, CStr(OptionalPath))
    Call R.Initialize(VL)
    Set ParseRegistry = R
End Function
Public Function ParseSkinFile(Optional OptionalPath)

End Function
Public Function ParsePlugin(Optional OptionalPath) As IPlugin
    Dim VL As String
    VL = IIf(IsMissing(OptionalPath), str, CStr(OptionalPath))
    Set ParsePlugin = LoadPluginFromFile(VL)
End Function
Public Function ParseSkin(Optional OptionalPath)

End Function
Public Function ParseData(Optional OptionalPath) As IData

End Function
Public Function ParseBlockReader(Optional OptionalPath)

End Function
Public Function ParseBlockWriter(Optional OptionalPath)

End Function


Public Function GetRandomAllString(ByVal Length As Long) As String
If Length <= 0 Then throw ZeroNegativeArgumentException("Length Must Be Greater Than Zero."), Me  ' check if length less than 0.
    Dim X As String, i As Long
    Call Randomize
    For i = 1 To Length
        X = X & Chr(CLng(Rnd * 255))
    Next
    GetRandomAllString = X
End Function
Public Function GetRandomString(ByVal Length As Long, ByVal stringType As StringTypes) As String
If Length <= 0 Then throw ZeroNegativeArgumentException("Length Must Be Greater Than Zero."), Me  ' check if length less than 0. ' check if length less than 0.
End Function
Public Function GetRandomAlphabeticString(ByVal Length As Long) As String
If Length <= 0 Then throw ZeroNegativeArgumentException("Length Must Be Greater Than Zero."), Me   ' check if length less than 0.
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
If Length <= 0 Then throw ZeroNegativeArgumentException("Length Must Be Greater Than Zero."), Me  ' check if length less than 0.
    Dim str As String, i As Long
    Call Randomize
    For i = 1 To Length
        str = str & CStr(Fix(Rnd * 10))
    Next
    GetRandomNumericString = str
End Function

' pattern : pattern of exported string. * = single character, ? = single alphabetic character , # = single numeric character , ! = space character
' Length : length of exported string.
' addto : add extra string to?
' additionalStringType : extra added strin type.
Public Function GetRandomStringByTemplate(ByVal Pattern As String, ByVal Length As Long, Optional ByVal addto As StringDirections = LeftToRight, Optional ByVal additionalStringType As StringTypes = stAll) As String
If Length <= 0 Then throw ZeroNegativeArgumentException("Length Must Be Greater Than Zero."), Me   ' check if length less than 0.
    Dim str As String, strLen As Long, i As Long, oStrLength As Long 'Other Strings Length
    strLen = Len(Pattern)
    Dim Char As String * 1, B As Long
    oStrLength = Length - strLen
    Call Randomize
    For i = 1 To strLen
        Char = mID(Pattern, i, 1)
        Select Case Char
            Case "*"
                str = str & Chr(CLng(Rnd * 255))
            Case "?"
                B = Fix(Rnd * 50)
                If B < 25 Then
                    B = B + 65
                Else
                    B = (B - 25) + 97
                End If
                str = str & Chr(B)
            Case "#"
                str = str & CStr(CByte(Rnd * 10))
            Case "!"
                B = Round(Rnd * 3)
                If B = 0 Then
                    str = str & " "
                ElseIf B = 1 Then
                    str = str & vbTab
                ElseIf B = 2 Then
                    str = str & vbCrLf
                Else
                    str = str & vbNewLine
                End If
            Case Else
                str = str & Char
        End Select
    Next
    If oStrLength > 0 Then
        If addto = LeftToRight Then
            GetRandomStringByTemplate = GetRandomString(oStrLength, additionalStringType) & str
        Else
            GetRandomStringByTemplate = str & GetRandomString(oStrLength, additionalStringType)
        End If
    Else
        GetRandomStringByTemplate = str
    End If
End Function

Private Function IClassClonable_Clone() As Object
    Dim sp As New StringParser
    Call sp.Initialize(str, isSpecial)
    Set IClassClonable_Clone = sp
End Function

Private Sub IClassTexer_fromByteArray(Value() As Byte)
    
End Sub
Private Sub IClassTexer_fromString(Value As String)
    
End Sub
Private Property Get IClassTexer_texHeader(texType As mint_TexingType) As String
    IClassTexer_texHeader = _
        mint_tex_createheader(texType, CLASSID, CLASSALIAS, CLASSVERSIONSTRING, 1)
End Property
Private Function IClassTexer_toByteArray() As Byte()
    
End Function
Private Function IClassTexer_toString() As String
    
End Function

Private Function ITargetStream_Available(Dir As StreamDirection) As Long
    Dim val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        val = ((Len(str) > 0) And inState)
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        val = val Or (outState)
    End If
    ITargetStream_Available = val
End Function

'<NOT CLOSABLE>
Private Sub ITargetStream_CloseStream(Optional direction As StreamDirection = 3&)
    If (direction And sdInStream) = sdInStream Then
        inState = False
    End If
    If (direction And sdOutStream) = sdOutStream Then
        outState = False
    End If
End Sub

Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = StreamDirection.sdBoth)
    If (direction And sdInStream) = sdInStream Then
        inState = True
    End If
    If (direction And sdOutStream) = sdOutStream Then
        outState = True
    End If
End Sub
Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean
    Dim val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        val = inState
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        val = val Or outState
    End If
    ITargetStream_getState = val
End Function

Private Sub ITargetStream_inStream(Data, Optional Length As Long = -1)
    If Not inState Then throw InvalidStatusException
    'Call mint_put_byte_array_to(Data, d, Length)
End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    ITargetStream_readBytes = StringToByteArray(mID(str, IIf(StartAt = -1, 1, StartAt + 1), Length))
End Function
Private Sub ITargetStream_outStream(Data, Optional Length As Long = -1)
    If Not outState Then throw InvalidStatusException
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
                        str = ""
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_NEWLINE Then
                        str = str & vbCrLf
                        Exit Sub
                    Else
                        throw InvalidArgumentTypeException("Streaming Command '" & strFirst & "' Is Not Available For ByteArray.")
                    End If
                Else
                    throw InvalidArgumentTypeException("Streaming Type Is Not Available For ByteArray.")
                End If
            End If
        End If
    End If
    str = str & CStr(Data)
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
