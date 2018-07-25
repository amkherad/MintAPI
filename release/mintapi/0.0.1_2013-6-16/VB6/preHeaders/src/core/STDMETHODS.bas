Attribute VB_Name = "STDMETHODS"
'@PROJECT_LICENSE
Option Explicit
Option Base 0
Const CLASSID As String = "STDMETHODS"

'-------------------------------------------------------

'' All Methods Grouped In:

'USE_STRINGMETHODS
'USE_RANDOMGENERATORS
'USE_ACTIVEXMETHODS
'USE_ARRAYMETHODS
'USE_TOOLS
'USE_USERINTERFACEMETHODS
'USE_ENVIRONMENT
'USE_NETCONNECTION
'USE_REGISTRYMETHODS

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

Private Enum ERR_Type
    ERR_Exception = -999999999
    ERR_Reflection = -999999998
    ERR_Abort = -999999997
End Enum

Public Enum RegistryHiveKeys
    rgHiveKEY_CLASSES_ROOT = HKEY_CLASSES_ROOT
    rgHiveKEY_CURRENT_USER = HKEY_CURRENT_USER
    rgHiveKEY_LOCAL_MACHINE = HKEY_LOCAL_MACHINE
    rgHiveKEY_USERS = HKEY_USERS
    rgHiveKEY_PERFORMANCE_DATA = HKEY_PERFORMANCE_DATA
    rgHiveKEY_CURRENT_CONFIG = HKEY_CURRENT_CONFIG
    rgHiveKEY_DYN_DATA = HKEY_DYN_DATA
End Enum

Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_SUCCESS = 0&

Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4

Const KEY_READ = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                          ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                          ' SYNCHRONIZE))
Const REG_OPENED_EXISTING_KEY = &H2
Const KEY_WRITE = &H20006  '((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or
                           ' KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Const MAXWin9xLength As Long = 255

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

Private Sub throw(ByRef LibraryName As String, ByRef ModuleName As String, ByVal SourceMethodName As String, ByVal ExceptionDescription As String, Optional ByVal Arguments As String = "", Optional ByVal ErrorNumber As Long, Optional ByVal ExceptionType As Long, Optional HelpFile, Optional HelpContext)
    Call err.Raise(IIf(IsMissing(ErrorNumber), 500, ErrorNumber), _
                   IIf(LibraryName = "", "STDMETHODS", LibraryName) & IIf(ModuleName = "", "", "::" & ModuleName) & IIf(SourceMethodName = "", IIf(Arguments = "", "", " =" & Arguments), "::" & SourceMethodName & "(" & IIf(Arguments = "", "", Arguments) & ")"), _
                   ExceptionDescription, _
                   HelpFile, _
                   HelpContext)
End Sub

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

#If USE_STRINGMETHODS Then

Public Function STD_ConvertStringToByteArray(str As String, Optional Length As Long = -1) As Byte()
    If Length = 0 Then Exit Function
    Dim B() As Byte
    Dim strLen As Long
    strLen = Len(str)
    If strLen = 0 Then Exit Function
    If Length < 0 Then Length = strLen
    ReDim B(Length)
    Dim i As Long
    For i = 1 To Length
        B(i - 1) = Asc(Mid(str, i, 1))
    Next
    STD_ConvertStringToByteArray = B()
End Function
Public Function STD_ConvertByteArrayToString(B() As Byte, Optional Length As Long = -1) As String
    If Length = 0 Then Exit Function
    Dim str As String, arrsize As Long
    On Error GoTo zeroLength
    arrsize = UBound(B) - LBound(B) + 1
zeroLength:
    If arrsize = 0 Then Exit Function
    If Length < 0 Then Length = arrsize
    Dim i As Long
    For i = LBound(B) To UBound(B)
        str = str & Chr(B(i))
    Next
    STD_ConvertByteArrayToString = str
End Function

Public Function STD_CountCharacters(ByRef Source As String, ByVal ValueToFind As String) As Long
    Dim i As Long, vtfLen As Long, counter As Long
    vtfLen = Len(ValueToFind)
    For i = 1 To Len(Source) - vtfLen
        If Mid$(Source, i, vtfLen) = ValueToFind Then
            counter = counter + 1
        End If
    Next
    STD_CountCharacters = counter
End Function

Public Function STD_MakeTrueLength(ByRef Source As String, ByVal Length As Long, ByVal FillWith As String, Optional ByVal EndOfString As Boolean = True) As String
    Dim SLen As Long
    SLen = Len(Source)
    If SLen = Length Then MakeTrueLength = Source: Exit Function
    Dim i As Long, Value As String
    If SLen < Length Then
        If EndOfString Then
            STD_MakeTrueLength = Left(Source, Length)
            Exit Function
        Else
            STD_MakeTrueLength = Right(Source, Length)
            Exit Function
        End If
    Else
        Dim AttachedValue As String
        Value = Source
        SLen = (Length - SLen)
        Do While SLen >= Len(AttachedValue)
            AttachedValue = AttachedValue & FillWith
        Loop
        If Len(AttachedValue) > SLen Then
            AttachedValue = Left(AttachedValue, SLen)
        End If
        If EndOfString Then
            STD_MakeTrueLength = Source & AttachedValue
            Exit Function
        Else
            STD_MakeTrueLength = AttachedValue & Source
            Exit Function
        End If
    End If
End Function

#End If

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

#If USE_RANDOMGENERATORS Then

Public Function STD_getRandomAllString(ByVal Length As Long) As String
If Length <= 0 Then Call err.Raise(500, "Length", "Length must be greater than zero.") ' check if length less than 0.
    Dim X As String, i As Long
    For i = 1 To Length
        X = X & Chr(CLng(Rnd * 255))
    Next
    STD_getRandomAllString = X
End Function
Public Function STD_getRandomString(ByVal Length As Long, ByVal stringType As StringTypes) As String
If Length <= 0 Then Call err.Raise(500, "Length", "Length must be greater than zero.") ' check if length less than 0.
Select Case stringType
    Case StringTypes.stAll
        STD_getRandomString = STD_getRandomAllString(Length)
    Case StringTypes.stAlphabetic
        STD_getRandomString = STD_getRandomAlphabeticString(Length)
    Case StringTypes.stNewLine
        STD_getRandomString = String(Length, vbCr)
    Case StringTypes.stNull
        STD_getRandomString = String(Length, Chr(0))
    Case StringTypes.stNumeric
        STD_getRandomString = STD_getRandomNumericString(Length)
    Case StringTypes.stSpace
        STD_getRandomString = Space(Length)
    Case StringTypes.stSpecialCharacters
        STD_getRandomString = STD_getRandomAllString(Length)
End Select
End Function
Public Function STD_getRandomAlphabeticString(ByVal Length As Long) As String
If Length <= 0 Then Call err.Raise(500, "Length", "Length must be greater than zero.") ' check if length less than 0.
    Dim X As String, i As Long, B As Byte
    For i = 1 To Length
        B = CLng(Rnd * 50)
        If B < 25 Then
            B = B + 65
        Else
            B = B + 97
        End If
        X = X & Chr(B)
    Next
    STD_getRandomAlphabeticString = X
End Function
Public Function STD_getRandomNumericString(ByVal Length As Long) As String
If Length <= 0 Then Call err.Raise(500, "Length", "Length Must Be Greater Than Zero.") ' check if length less than 0.
    Dim str As String, i As Long
    For i = 1 To Length
        str = str & CStr(Fix(Rnd * 10))
    Next
    STD_getRandomNumericString = str
End Function

Public Function STD_getRandomStringByTemplate(ByVal pattern As String, ByVal Length As Long, Optional ByVal addto As StringDirections = LeftToRight, Optional ByVal additionalStringType As StringTypes = stAll) As String
If Length <= 0 Then Call err.Raise(500, "Length", "Length Must Be Greater Than Zero.") ' check if length less than 0.
    Dim str As String, strLen As Long, i As Long, oStrLength As Long 'Other Strings Length
    strLen = Len(pattern)
    Dim char As String * 1, B As Long
    oStrLength = Length - strLen
    For i = 1 To strLen
        char = Mid(pattern, i, 1)
        Select Case char
            Case "*"
                str = str & Chr(CLng(Rnd * 255))
            Case "?"
                B = CLng(Rnd * 50)
                If B < 25 Then
                    B = B + 65
                Else
                    B = B + 97
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
                str = str & char
        End Select
    Next
    If oStrLength > 0 Then
        If addto = LeftToRight Then
            STD_getRandomStringByTemplate = STD_getRandomString(oStrLength, additionalStringType) & str
        Else
            STD_getRandomStringByTemplate = str & STD_getRandomString(oStrLength, additionalStringType)
        End If
    Else
        STD_getRandomStringByTemplate = str
    End If
End Function

#End If

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

#If USE_ACTIVEXMETHODS Then

Public Function STD_SplitVBCommandValue(ByVal Command___ As String) As String()
If Trim(Command___) = "" Then GoTo raiseArgumentIncorrectError
    Dim str() As String, Count As Long
    str = Split(Command___, """ """)
    On Error GoTo zeroLength
        Count = (UBound(str) - LBound(str)) + 1
zeroLength:
    If Count <= 0 Then GoTo raiseArgumentIncorrectError
    str(LBound(str)) = Mid(str(LBound(str)), InStr(str(LBound(str)), """") + 1, Len(str(LBound(str))) - 1)
    str(UBound(str)) = Mid(str(UBound(str)), 1, Len(str(UBound(str))) - (Len(str(UBound(str))) - InStr(str(UBound(str)), """")) - 1)
    STD_SplitVBCommandValue = str()
    Exit Function
raiseArgumentIncorrectError:
    Call Throw("", CLASSID, "SplitVBCommandValue", "Argument Incorrect.", "ByVal Command As String")
End Function

Public Function STD_ActiveXReg(fName As String, Func As Long) As Long
    Dim regLib As Long, process As Long, succeed As Long
    Dim h1 As Long, xc As Long, ID As Long
    Dim P As String
    
    Select Case Func
        Case REGISTER: P = "DllRegisterServer"
        Case UNREGISTER: P = "DllUnregisterServer"
        Case Else: STD_ActiveXReg = INVALID: Exit Function
    End Select
    regLib = API_LoadLibrary(fName)
    If regLib = 0 Then
        STD_ActiveXReg = NOTFOUND
        Exit Function
    End If
    process = API_GetProcAddress(regLib, P)
    If process = 0 Then
        STD_ActiveXReg = NOTACTX
    Else
        Dim ASA As API_SECURITY_ATTRIBUTES
        h1 = API_CreateThread(ASA, 0&, ByVal process, ByVal 0&, 0&, ID)
        
        If h1 = 0 Then
            STD_ActiveXReg = NOTHREAD
        Else
            succeed = _
                (API_WaitForSingleObject(h1, 10000) = 0)
            If succeed Then
                Call API_CloseHandle(h1)
                STD_ActiveXReg = SUCCESS
            Else
                Call API_GetExitCodeThread(h1, xc)
                Call API_ExitThread(xc)
                STD_ActiveXReg = FAILURE
            End If
        End If
    End If
    Call API_FreeLibrary(regLib)
End Function

'Handle differently if it is an EXE
Private Function STD_ItIsNotEXE(ByVal path As String, Func As Long) As Boolean
    STD_ItIsNotEXE = False
    
    path = Trim$(path)
    If Len(path) < 5 Then Exit Function
    
    If LCase$(getFileExtention(path)) <> "exe" Then
        STD_ItIsNotEXE = True
        Exit Function
    End If
    
    Dim sWorkDir As String, sFile As String, sCommand As String
    
    sFile = path
    sWorkDir = vbNullString
    sCommand = IIf(Func = REGISTER, "/regserver", "/unregserver")

    STD_ItIsNotEXE = str$(API_ShellExecute(0, "open", sFile, sCommand, sWorkDir, 1))
End Function

#End If

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

#If USE_ARRAYMETHODS Then

Public Function STD_IsArray(vArray As Variant) As Boolean
    STD_IsArray = (VarType(vArray) And vbArray) = vbArray
End Function
Public Function STD_IsEmptyArray(vArray As Variant) As Boolean
'##BLOCK_DESCRIPTION Returns true if the variant passed is an empty array
    STD_IsEmptyArray = (arraySize(vArray) = 0)
End Function
Public Function STD_ArraySize(targetArray) As Long
If Not (VarType(targetArray) And vbArray) = vbArray Then Call Throw("", CLASSID, "ArraySize", "Only Arrays Accepted.")
    On Error GoTo zeroLength
    STD_ArraySize = UBound(targetArray) - LBound(targetArray) + 1
zeroLength:
End Function
Public Function STD_MaxValue(vArray As Variant) As Variant
        '##BLOCK_DESCRIPTION Returns the maximum value in the array passed.
        '##PARAMETER_DESCRIPTION vArray An array - use Array() function.
        Dim max As Variant
        Dim i As Long, vt As VariantTypeConstants
        If Not IsArray(vArray) Then Call err.Raise(500, "MaxValue", "MaxValue requires an array as parameter")
        If Not arraySize(vArray) > 0 Then Call err.Raise(500, "MaxValue", "Empty array")
        For i = 0 To UBound(vArray)
                If i = 0 Then
                        max = vArray(0)
                        vt = VarType(vArray(0))
                Else
                        If vArray(i) > max Then max = vArray(i)
                        If Not VarType(vArray(i)) = vt Then Call err.Raise(500, "MaxValue", "Array items must be of the same type")
                End If
        Next i
        STD_MaxValue = max
End Function
Public Function STD_MinValue(vArray As Variant) As Variant
        '##BLOCK_DESCRIPTION Returns the minimum value in the array passed.
        '##PARAMETER_DESCRIPTION vArray An array - use Array() function.
        Dim Min As Variant
        Dim i As Long, vt As VariantTypeConstants
        If Not IsArray(vArray) Then Call err.Raise(500, "MinValue", "MinValue requires an array as parameter")
        If Not arraySize(vArray) > 0 Then Call err.Raise(500, "MinValue", "Empty array")
        For i = 0 To UBound(vArray)
                If i = 0 Then
                        Min = vArray(0)
                        vt = VarType(vArray(0))
                Else
                        If Not VarType(vArray(i)) = vt Then Call err.Raise(500, "MinValue", "Array items must be of the same type")
                        If vArray(i) < Min Then Min = vArray(i)
                End If
        Next i
        STD_MinValue = Min
End Function

#End If

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

#If USE_FILESYSTEM Then

Public Function STD_getSubDirectories(ByVal path As String) As String()
    Dim str() As String, strCount As Long
    Dim C As String
    C = Dir(path)
    If C = "" Then Exit Function
    While C <> ""
        ReDim Preserve str(strCount)
        str(strCount) = C
        strCount = strCount + 1
        C = Dir
    Wend
    STD_getSubDirectories = str()
End Function
Public Function STD_CountSubDirectories(ByVal path As String) As Long
    Dim strCount As Long
    Dim C As String
    C = Dir(path)
    If C = "" Then Exit Function
    While C <> ""
        strCount = strCount + 1
        C = Dir
    Wend
    STD_CountSubDirectories = strCount
End Function

Public Property Get STD_Directory() As String
    STD_Directory = FileSystem.CurDir
End Property
Public Property Let STD_Directory(ByVal Value As String)
    Call FileSystem.ChDir(Value)
End Property
Public Property Get STD_Drive() As String
    STD_Drive = Left(FileSystem.CurDir, 3)
End Property
Public Property Let STD_Drive(ByVal Value As String)
    Call FileSystem.ChDrive(Value)
End Property

#End If

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

#If USE_TOOLS Then

Public Function STD_Compare(str1 As String, str2 As String) As Boolean
    STD_Compare = (str1 Like str2)
End Function
Public Sub STD_ResetVar(ByRef Var)
' resets a variable with the appropriate value
' depending of its type

If IsObject(Var) Then
    Set Var = Nothing
Else
    If IsArray(Var) Then
        Erase Var
    Else
        Var = Empty
    End If
End If

End Sub
Public Function STD_getObject(ByVal Address As Long) As Object
    Dim X As Object
    Call API_CopyMemory(X, Address, 4)
    Set STD_getObject = X
End Function

Public Function STD_getFileName(ByVal path As String) As String
On Error GoTo err
    Dim slashIndex As Long, backslashIndex As Long
    slashIndex = InStrRev(path, "/")
    backslashIndex = InStrRev(path, "\")
    slashIndex = IIf(slashIndex >= backslashIndex, slashIndex, backslashIndex)
    If slashIndex = 0 Then Call Throw("", CLASSID, "getFileName", "Not True Path.")
    backslashIndex = Len(path)
    STD_getFileName = Right(path, backslashIndex - slashIndex)
Exit Function
err:
End Function
Public Function STD_getFileExtention(ByVal path As String) As String
On Error GoTo err
    Dim slashIndex As Long, backslashIndex As Long
    slashIndex = InStrRev(path, "/")
    backslashIndex = InStrRev(path, "\")
    slashIndex = IIf(slashIndex >= backslashIndex, slashIndex, backslashIndex)
    If slashIndex = 0 Then Call Throw("", CLASSID, "getFileExtention", "Not True Path.")
    backslashIndex = InStrRev(path, ".")
    If backslashIndex = 0 Then
        STD_getFileExtention = ""
        Exit Function
    End If
    If slashIndex > backslashIndex Then
        STD_getFileExtention = ""
        Exit Function
    Else
        slashIndex = Len(path)
        STD_getFileExtention = Right(path, slashIndex - backslashIndex)
    End If
Exit Function
err:
End Function
Public Function STD_getFileNameOnly(ByVal path As String) As String
On Error GoTo err
    Dim slashIndex As Long, backslashIndex As Long, fLen As Long
    fLen = Len(path)
    slashIndex = InStrRev(path, "/")
    backslashIndex = InStrRev(path, "\")
    slashIndex = IIf(slashIndex >= backslashIndex, slashIndex, backslashIndex)
    If slashIndex = 0 Then Call Throw("", CLASSID, "getFileNameOnly", "Not True Path.")
    backslashIndex = InStrRev(path, ".")
    If backslashIndex = 0 Then
        STD_getFileNameOnly = Right(path, fLen - slashIndex)
        Exit Function
    End If
    If slashIndex > backslashIndex Then
        STD_getFileNameOnly = Right(path, fLen - slashIndex)
        Exit Function
    Else
        STD_getFileNameOnly = Mid(path, slashIndex + 1, backslashIndex - slashIndex - 1)
    End If
Exit Function
err:
End Function

#End If

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

#If USE_USERINTERFACEMETHODS Then

Public Function STD_DialogMessage(ByVal Message As String, Optional ByVal Flags As VbMsgBoxStyle, Optional ByVal Title As String, Optional ByVal ParenthWnd As Long = 0) As VbMsgBoxResult
    If IsMissing(Title) Then Title = App.Title
    STD_DialogMessage = API_MessageBox(ParenthWnd, Message, Title, Flags)
End Function
Public Sub STD_ScrollControl(ByVal hwnd As Long, ByVal Direction As Long, ByVal Action As Long, ByVal Amount As Long)
    Dim position As Integer
    ' What direction are we going
    If Direction = SBS_HORZ Then
        ' What action are we taking (Jumping or Relative)
        If Action = API_SCROLLACTION_RELATIVE Then
            position = API_GetScrollPos(hwnd, SBS_HORZ) + Amount
        Else
            position = Amount
        End If
        ' Make it so
        If (API_SetScrollPos(hwnd, SBS_HORZ, position, True) <> -1) Then
            Call API_PostMessageA(hwnd, WM_HSCROLL, SB_THUMBPOSITION + &H10000 * position, Nothing)
        Else
            Call Throw("", CLASSID, "STD_ScrollControl", "Can't Change Scroll Value.")
        End If

    Else

        ' What action are we taking (Jumping or Relative)
        If Action = API_SCROLLACTION_RELATIVE Then
            position = API_GetScrollPos(hwnd, SBS_VERT) + Amount
        Else
            position = Amount
        End If

        ' Make it so
        If (API_SetScrollPos(hwnd, SBS_VERT, position, True) <> -1) Then
            Call API_PostMessageA(hwnd, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * position, Nothing)
        Else
            Call Throw("", CLASSID, "STD_ScrollControl", "Can't Change Scroll Value.")
        End If
    End If
End Sub
Public Function STD_ScreenShot() As StdPicture
Dim hWndSrc As Long, hSrcDC As Long, res As Long, pic As StdPicture
    hWndSrc = API_GetDesktopWindow()
    hSrcDC = API_GetDC(hWndSrc)
    res = API_BitBlt(pic.hdc, 0, 0, Screen.Width, Screen.Width, hSrcDC, 0, 0, &HCC0020)
    res = API_ReleaseDC(hWndSrc, hSrcDC)
End Function

#End If

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

#If USE_ENVIRONMENT Then

Public Function STD_HaveSoundCard() As Boolean
    STD_HaveSoundCard = (API_waveOutGetNumDevs > 0)
End Function
Public Function STD_getLocaleString(ByVal lLocaleNum As Long) As String
    'Generic routine to get the locale string from the Operating system.
    Dim lBuffSize As String
    Dim sBuffer As String
    Dim lRet As Long

    lBuffSize = 256
    sBuffer = String(lBuffSize, vbNullChar)

    'Get the information from the registry
    lRet = API_GetLocaleInfo(LOCALE_USER_DEFAULT, lLocaleNum, sBuffer, lBuffSize)
    'If lRet > 0 then success - lret is the size of the string returned
    If lRet > 0 Then
        STD_getLocaleString = Left$(sBuffer, lRet - 1)
    End If
End Function
' CGLocale Class
' This DLL allows you to obtain the Regional Settings for your System.
' This will ensure that all displays etc are correct for the country of use
'
Public Property Get STD_DateFormat() As String
    ' This function will return the Locale date format for the system. Note that the
    ' returned Year is always formatted to 'YYYY' regardless, to ensure Y2k compliance.
    Dim sDateFormat As String
    On Error GoTo vbErrorHandler
    sDateFormat = getLocaleString(LOCALE_SSHORTDATE)

    ' Make sure we always have YYYY format for y2k
    If InStr(1, sDateFormat, "YYYY", vbTextCompare) = 0 Then
        Replace sDateFormat, "YY", "YYYY"
    End If
    STD_DateFormat = sDateFormat
Exit Property
vbErrorHandler:
    Call Throw("", CLASSID, "STD_DateFormat", err.Description)
    'err.Raise err.Number, "LocaleSettings GetDateFormat", err.Description
End Property
Public Property Get STD_TimeFormat() As String
    'This function returns the locale's defined Time Format.
    STD_TimeFormat = getLocaleString(LOCALE_STIMEFORMAT)
Exit Property
vbErrorHandler:
    Call Throw("", CLASSID, "STD_TimeFormat", err.Description)
    'err.Raise err.Number, "LocaleSettings GetTimeFormat", err.Description
End Property
Public Function STD_NumberFormat() As String
' This function returns the Locales defined Decimal Number format
    Dim lBuffLen As Long
    Dim sBuffer As String
    Dim sDecimal As String
    Dim sThousand As String
    Dim lResult As Long
    Dim sNumFormat As String

    On Error GoTo vbErrorHandler

    'Setup a buffer to receive the settings
    lBuffLen = 128
    sBuffer = String(lBuffLen, vbNullChar)

    lResult = API_GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sBuffer, lBuffLen)
    If lResult <= 0 Then Exit Function

    sDecimal = Left$(sBuffer, lResult - 1)

    sBuffer = String(lBuffLen, vbNullChar)
    lResult = API_GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, sBuffer, lBuffLen)
    If lResult <= 0 Then Exit Function

    sThousand = Left$(sBuffer, lResult - 1)

    STD_NumberFormat = "###" & sThousand & "###" & sDecimal & "######"
Exit Function
    
vbErrorHandler:
    Call Throw("", CLASSID, "NumberFormat", err.Description)
    'err.Raise err.Number, "LocaleSettings GetNumberFormat", err.Description
End Function
Public Function STD_ThousandSpecifier() As String
    'This function returns the correct Thousand Specifier for the system Locale
    STD_ThousandSpecifier = STD_getLocaleString(LOCALE_STHOUSAND)
End Function
Public Function STD_DecimalSpecifier() As String
    'This function returns the correct Decimal Specifier for the system Locale
    STD_DecimalSpecifier = STD_getLocaleString(LOCALE_SDECIMAL)
End Function
Public Function STD_CurrencySpecifier() As String
    'This function returns the correct Currency Specifier for the system Locale
    STD_CurrencySpecifier = STD_getLocaleString(LOCALE_SCURRENCY)
End Function
Public Function STD_SystemLanguageID() As Long
    'Returns the System Language ID for the machine
    STD_SystemLanguageID = API_GetSystemDefaultLangID
End Function
Public Function STD_SystemLanguageName() As String
    'Returns the System Language Name eg : English (United Kingdom)
    Dim lLangID As Long
    Dim sBuffer As String
    Dim lBuffSize As Long
    Dim lRet As Long

    On Error GoTo vbErrorHandler

    lLangID = API_GetSystemDefaultLangID
    'Setup a buffer to receive the settings
    lBuffSize = 50
    sBuffer = String(lBuffSize, vbNullChar)
    lRet = API_VerLanguageName(lLangID, sBuffer, lBuffSize)
    If lRet > 0 Then
        STD_SystemLanguageName = Left$(sBuffer, lRet)
    End If
Exit Function
vbErrorHandler:
    Call Throw("", CLASSID, "DateFormat", err.Description)
    'err.Raise err.Number, "LocaleSettings GetSysLanguageName", err.Description
End Function
Public Function STD_ShortMonthName(ByVal iMonthNum As Integer) As String
    'Returns the short-month-name for the specified Month Number
    'eg 1=Jan, 2=Feb (on English machines)
    STD_ShortMonthName = STD_getLocaleString(LOCALE_SABBREVMONTHNAME1 - 1 + iMonthNum)
End Function
Public Function STD_MonthName(ByVal iMonthNum As Integer) As String
    'Returns the Full-Month-Name for the specified month number
    'eg. 1=January, 2=February (on english machines)
    STD_MonthName = STD_getLocaleString(LOCALE_SMONTHNAME1 + iMonthNum - 1)
End Function
Public Function STD_ShortDayName(ByVal iDayNum As Integer) As String
    'Returns the Short-Day-Name for the specified Day Number
    'eg. 1=Mon, 2=Tue (on english machines)
    STD_ShortDayName = STD_getLocaleString(LOCALE_SABBREVDAYNAME1 + iDayNum - 1)
End Function
Public Function STD_DayName(ByVal iDayNum As Integer) As String
    'Returns the Full Day Name for the specified Day number
    'eg. 1=Monday, 2=Tuesday (on english machines)
    STD_DayName = STD_getLocaleString(LOCALE_SDAYNAME1 + iDayNum - 1)
End Function
Public Function STD_Country() As String
    'Returns the Country Name eg. 'United Kingdom'
    STD_Country = STD_getLocaleString(LOCALE_SENGCOUNTRY)
End Function
Public Function STD_LanguageName() As String
    'Returns the Native Language Name eg. 'English'
    STD_LanguageName = STD_getLocaleString(LOCALE_SNATIVELANGNAME)
End Function
Public Function STD_NativeCountryName() As String
    STD_NativeCountryName = STD_getLocaleString(LOCALE_SNATIVECTRYNAME)
End Function
Public Function STD_PositiveSign() As String
    'Returns the symbol used for the positive sign eg. +
    STD_PositiveSign = STD_getLocaleString(LOCALE_SPOSITIVESIGN)
End Function
Public Function STD_NegativeSign() As String
' Returns the symbol used for the negative sign eg. -
    STD_NegativeSign = STD_getLocaleString(LOCALE_SNEGATIVESIGN)
End Function

#End If

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

#If USE_NETCONNECTION Then

Public Function STD_InternetConnectionStatus() As API_ConnectionSettings
    Dim TRasCon(255) As API_RASCONN95
    Dim lg As Long
    Dim lpcon As Long
    Dim retval As Long
    Dim Tstatus As API_RASCONNSTATUS95
    '
    TRasCon(0).dwSize = 412
    lg = 256 * TRasCon(0).dwSize
    '
    retval = API_RasEnumConnections(TRasCon(0), lg, lpcon)

    If retval <> 0 Then
        Call Throw("", CLASSID, "InternetConnectionStatus", "Error In Method RasEnumConnectionsA() In RasApi32.dll.")
        Exit Function
    End If
    
    Tstatus.dwSize = 160
    retval = API_RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)

    If Tstatus.RasConnState = NET_CONNECTED_VALUE Then
        STD_InternetConnectionStatus = csConnected
    Else
        STD_InternetConnectionStatus = csNotConnected
    End If
End Function

#End If

'-------------------------------------------------------
'-------------------------------------------------------
'-------------------------------------------------------

#If USE_REGISTRYMETHODS Then

Public Function STD_regCreateRegistryKey(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String) As Boolean
Dim Handle As Long, disposition As Long
        Dim sa As API_SECURITY_ATTRIBUTES
If API_RegCreateKeyEx(hKey, KeyName, 0, 0, 0, 0, sa, Handle, disposition) Then
    'Err.Raise 1001, , "Unable to create the registry key"
    STD_regCreateRegistryKey = False
Else
    ' Return True if the key already existed.
    STD_regCreateRegistryKey = (disposition = REG_OPENED_EXISTING_KEY)
    ' Close the key.
    Call API_RegCloseKey(Handle)
    STD_regCreateRegistryKey = True
End If

End Function
' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Return True if a Registry key exists
Public Function STD_regCheckRegistryKey(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String) As Boolean
Dim Handle As Long
Dim Ret As Long
' Try to open the key

Ret = API_RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, Handle)
Select Case Ret
    Case 0:
        ' The key exists
        STD_regCheckRegistryKey = True
        ' Close it before exiting
        API_RegCloseKey Handle
    Case 5:
        Call Throw("", CLASSID, "CheckRegistryKey", "Access Denied.")
        STD_regCheckRegistryKey = False
    Case Else:
        'Call Throw("", CLASSID, "CheckRegistryKey", "Access Denied.")
        STD_regCheckRegistryKey = False
End Select
End Function
' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Enumerate registry keys under a given key
' returns a collection of strings
Public Function STD_regEnumRegistryKeys(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String) As Collection
Dim Handle As Long
Dim Length As Long
Dim Index As Long
Dim subkeyName As String

' initialize the result collection
Set STD_regEnumRegistryKeys = New Collection

' Open the key, exit if not found
If Len(KeyName) Then
    If API_RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, Handle) Then Exit Function
    ' in all case the subsequent functions use hKey
    hKey = Handle
End If

Do
    ' this is the max length for a key name
    Length = 260
    subkeyName = Space$(Length)
    ' get the N-th key, exit the loop if not found
    If API_RegEnumKey(hKey, Index, subkeyName, Length) Then Exit Do
    
    ' add to the result collection
    subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
    Call STD_regEnumRegistryKeys.Add(subkeyName)
    ' prepare to query for next key
    Index = Index + 1
Loop
' Close the key, if it was actually opened
If Handle Then Call API_RegCloseKey(Handle)
End Function
Public Function STD_regSeekRegistryValue(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, _
    ByVal EntryName As String, ByRef EntryValue) As Boolean
Dim re() As RegistryEntry
Dim reLen As Long
Dim i As Long

re = STD_regEnumRegistryValues(hKey, KeyName)

If Not STD_IsEmptyArray(re) Then
    reLen = STD_ArraySize(re)
    Call STD_ResetVar(EntryValue)
    If reLen > 0 Then
        For i = 0 To reLen - 1
            If re(i).EntryName = EntryName Then
                EntryValue = re(i).EntryValue
                STD_regSeekRegistryValue = True
                Exit Function
            End If
        Next i
    End If
End If
End Function
Private Function STD_GetModuleName(ByRef SourceModule) As String

    Select Case VarType(SourceModule)
        Case vbObject:
        STD_GetModuleName = TypeName(SourceModule)
    Case vbString:
        STD_GetModuleName = SourceModule
    Case Else:
        STD_GetModuleName = "<UndefinedModule>"
End Select

End Function
Public Function STD_regEnumRegistryValues(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String) As RegistryEntry()
' ritorna un collezione di coppie (nome,valore) dove
' valore è del tipo relativo a quello rappresentato nel registro
Dim Handle As Long
Dim Length As Long
Dim Index As Long
Dim subkeyName As String, res As Long
Dim ValName As String, lenValName As Long
Const ValLen As Long = 1024
Dim ValType As Long
Dim DataBuffer() As Byte, lenDataBuffer As Long
Const DataBufferLen As Long = 4096
Dim byteArrayItem() As Byte
Dim stringItem As String
Dim longItem As Long
Dim Item As Variant
Dim i As Long, S As String
Dim retval() As RegistryEntry, rEntry As RegistryEntry

' Open the key, exit if not found
If Len(KeyName) Then
    If API_RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, Handle) Then Exit Function
    ' in all case the subsequent functions use hKey
    hKey = Handle
End If

Do
    ValName = Space$(ValLen)
    lenValName = ValLen
    ReDim DataBuffer(DataBufferLen)
    lenDataBuffer = DataBufferLen
    res = API_RegEnumValue(hKey, Index, ValName, lenValName, 0&, ValType, DataBuffer(0), lenDataBuffer)
    If res = ERROR_SUCCESS Then
        
        Call ResetVar(Item)
        Erase byteArrayItem
        stringItem = ""
        longItem = 0
        
        Select Case ValType
            Case REG_BINARY ' 3 ==> ritorna un array di byte
                If lenDataBuffer > 0 Then
                    ReDim byteArrayItem(lenDataBuffer - 1)
                    For i = 0 To lenDataBuffer - 1
                        byteArrayItem(i) = DataBuffer(i)
                    Next i
                End If
                Item = byteArrayItem
                
            Case REG_DWORD ' 4 ==> ritorna un long
                If lenDataBuffer > 0 Then
                    S = Space$(lenDataBuffer)
                    longItem = 0
                    For i = 1 To lenDataBuffer
                        On Error Resume Next
                            longItem = longItem + 256 ^ (i - 1) * DataBuffer(i - 1)
                        On Error GoTo 0
                    Next i
                End If
                Item = longItem
                
            Case Else
                If lenDataBuffer > 0 Then
                    stringItem = Space$(lenDataBuffer - 1)
                    For i = 1 To lenDataBuffer - 1
                        Mid$(stringItem, i, 1) = Chr$(DataBuffer(i - 1))
                    Next i
                End If
                Item = stringItem
                
        End Select
                
        ReDim Preserve retval(Index)
        Set rEntry = New RegistryEntry
        rEntry.EntryName = Left(ValName, lenValName)
        rEntry.EntryValue = Item
        Set retval(Index) = rEntry
        Index = Index + 1
    End If
Loop While (res = ERROR_SUCCESS)
If res = ERROR_NO_MORE_ITEMS Then 'ok
   STD_regEnumRegistryValues = retval
End If
End Function
Public Function STD_regReadEntryValue(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, _
            ByVal EntryName As String, ByRef EntryValue As Variant) As Boolean
        ' cerca nella chiave specificata na voce; se la trova ritorna true
        ' ed aggiorna il valore di EntryValue
        Dim res() As RegistryEntry
        Dim i As Long
        Dim v As Variant, B() As Byte
        On Error GoTo xERR

    res = STD_regEnumRegistryValues(hKey, KeyName)
    If Not IsEmptyArray(res) Then
            For i = LBound(res) To UBound(res)
            If Trim(LCase(res(i).EntryName)) = Trim(LCase(EntryName)) Then
                v = res(i).EntryValue
                STD_regReadEntryValue = True
                    Select Case VarType(EntryValue)
                    Case vbString:  EntryValue = CStr(v)
                    Case vbBoolean: EntryValue = CBool(v)
                    Case vbLong:    EntryValue = CLng(v)
                    Case vbInteger: EntryValue = CInt(v)
                    Case vbSingle:    EntryValue = CSng(v)
                    Case vbDouble:    EntryValue = CDbl(v)
                    Case vbCurrency:    EntryValue = CCur(v)
                    Case vbArray + vbByte: EntryValue = v
                    Case Else: Call Throw("", CLASSID, "STD_regSetRegistryValue", "Unsupported Value Type :" & TypeName(Value), 1001)
                    STD_regReadEntryValue = False
                    End Select
                    'EntryValue = v
                    Exit Function
                End If
        Next i
        End If
Exit Function
xERR:
    STD_regReadEntryValue = False
End Function
' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Delete a registry key
'
' Under Windows NT it doesn't work if the key contains subkeys

Public Function STD_regDeleteRegistryKey(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String) As Boolean
    STD_regDeleteRegistryKey = (API_RegDeleteKey(hKey, KeyName) = 0)
End Function
' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Delete a registry value
' Return True if successful, False if the value hasn't been found
Public Function STD_regDeleteRegistryValue(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, ByVal ValueName As String) As Boolean
Dim Handle As Long
Dim Ret As Long

' Open the key, exit if not found
If API_RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, Handle) Then Exit Function
Call err.Clear
Ret = API_RegDeleteValue(Handle, ValueName)
' Delete the value (returns 0 if success)
'Debug.Print Ret, Err.LastDllError
STD_regDeleteRegistryValue = (Ret = 0)
' Close the handle
Call API_RegCloseKey(Handle)
End Function
' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Write or Create a Registry value
' returns True if successful
'
' Use KeyName = "" for the default value
'
' Value can be an integer value (REG_DWORD), a string (REG_SZ)
' or an array of binary (REG_BINARY). Raises an error otherwise.
Public Function STD_regSetRegistryValue(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, ByVal ValueName As String, Value As Variant) As Boolean
Dim Handle As Long
Dim lngValue As Long
Dim StrValue As String
Dim binValue() As Byte
Dim Length As Long
Dim retval As Long

' Open the key, exit if not found
If API_RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, Handle) <> 0 Then Exit Function

retval = -1

' three cases, according to the data type in Value
Select Case VarType(Value)
    Case vbInteger, vbLong
        lngValue = CLng(Value)
        retval = API_RegSetValueEx(Handle, ValueName, 0, REG_DWORD, lngValue, 4)
    Case vbString, vbBoolean
        StrValue = CStr(Value)
        retval = API_RegSetValueEx(Handle, ValueName, 0, REG_SZ, ByVal StrValue, _
            Len(StrValue))
    Case vbArray + vbByte
        binValue = Value
        Length = UBound(binValue) - LBound(binValue) + 1
        retval = API_RegSetValueEx(Handle, ValueName, 0&, REG_BINARY, _
            binValue(LBound(binValue)), Length)
    Case Else
        Call API_RegCloseKey(Handle)
        Call Throw("", CLASSID, "STD_regSetRegistryValue", "Unsupported Value Type :" & TypeName(Value), 1001)
End Select

' Close the key and signal success
API_RegCloseKey Handle
' signal success if the value was written correctly
STD_regSetRegistryValue = (retval = 0)

End Function
Public Function STD_regEraseRegistryTree(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, ByRef MaxLevelsErased As Long, Optional ByRef TotKeys As Long = 0, Optional ByRef DelKeys As Long = 0) As Boolean
    ' elimina una chiave e tutte le sue sottochiavi
    ' ritorna true se tutte la chiave e le sottochivi  sono state eliminate
    ' totKeys = numero chiavi navigate, DelKeys = numero chiavi cancellate
    ' MaxLevelErased = numero massimo di livello delle sottochiavi
    
    Dim regEntries As Collection
    Dim rEntry As Variant
    Dim MaxLE As Long, ActLevel As Long, PassLevel As Long
    Static notFirstTime As Boolean, Ret As Boolean
    Dim ActLevelDelCount As Long
        
    If Not notFirstTime Then
        If Not STD_regCheckRegistryKey(hKey, ByVal KeyName) Then
            STD_regEraseRegistryTree = False
            Exit Function
        End If
        notFirstTime = True
        MaxLevelsErased = 0 ' init first time
        TotKeys = 1
        DelKeys = 0
    End If
    
    Set regEntries = STD_regEnumRegistryKeys(hKey, KeyName)
        
    If regEntries.Count = 0 Then ' it's a leaf
        If STD_regCheckRegistryKey(hKey, KeyName) Then
            Ret = regDeleteRegistryKey(hKey, KeyName)
            If Ret Then DelKeys = DelKeys + 1
            STD_regEraseRegistryTree = Ret
        End If
        Exit Function
    Else
        TotKeys = TotKeys + regEntries.Count
        ActLevel = MaxLevelsErased + 1
        MaxLE = ActLevel
        ActLevelDelCount = 0
        For Each rEntry In regEntries
            PassLevel = ActLevel
            Ret = STD_regEraseRegistryTree(hKey, KeyName & "\" & rEntry, PassLevel, TotKeys, DelKeys)
            If Ret Then ActLevelDelCount = ActLevelDelCount + 1
            If (PassLevel > MaxLE) And Ret Then MaxLE = PassLevel
        Next
        MaxLevelsErased = MaxLE
        Ret = STD_regDeleteRegistryKey(hKey, KeyName)
        If Ret Then DelKeys = DelKeys + 1
        If ActLevel = 1 Then
            STD_regEraseRegistryTree = (TotKeys = DelKeys)
            notFirstTime = False
        Else
            STD_regEraseRegistryTree = (ActLevelDelCount = regEntries.Count)
        End If
    End If
End Function
Public Function STD_regSeekRegistryLeafs(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, _
    Optional ByVal MaxEntries As Long = -1) As Collection
    ' ritorna una collection di stringhe
Dim REGKEY As Variant, SeekKeys As Collection, SeekKey As Variant
Dim EnumKeys As Collection
Dim RetValue As New Collection
Dim Key As String, IsLeaf As Boolean, AddIT As Boolean
Dim TotKeys As New Collection
Static KeysFound As Long
Static Level As Long

Level = Level + 1
Set EnumKeys = STD_regEnumRegistryKeys(hKey, KeyName)

If EnumKeys Is Nothing Then
    ' è una foglia
    IsLeaf = True
Else
    IsLeaf = (EnumKeys.Count = 0)
End If

If IsLeaf Then
    AddIT = IIf(MaxEntries < 0, True, (KeysFound < MaxEntries))
    If AddIT Then
        Call TotKeys.Add(KeyName)
        KeysFound = KeysFound + 1
    End If
Else
    For Each REGKEY In EnumKeys
        Key = AppendToPath(KeyName, REGKEY)
        If MaxEntries < 0 Then
            Set SeekKeys = STD_regSeekRegistryLeafs(hKey, Key)
        Else
            Set SeekKeys = STD_regSeekRegistryLeafs(hKey, Key, MaxEntries)
        End If

        If SeekKeys.Count > 0 Then
            For Each SeekKey In SeekKeys
                Call TotKeys.Add(SeekKey)
            Next
        End If
    Next
End If
Set STD_regSeekRegistryLeafs = TotKeys

Level = Level - 1
If Level = 0 Then KeysFound = 0

End Function
Public Function STD_regFlushRegistryChanges(ByVal hKey As RegistryHiveKeys) As Boolean
    STD_regFlushRegistryChanges = (API_RegFlushKey(hKey) = 0)
End Function

#End If


' INT8 ENTRYPOINT toINT8(INT8 Value){return Value;}
' INT16 ENTRYPOINT INT8plusINT8(INT8 v1,INT8 v2){return v1 + v2;}
' INT16 ENTRYPOINT INT8minusINT8(INT8 v1,INT8 v2){return v1 - v2;}
' INT16 ENTRYPOINT INT8intsubINT8(INT8 v1,INT8 v2){return v1 / v2;}
' INT16 ENTRYPOINT INT8subINT8(INT8 v1,INT8 v2){return (float)v1 / (float)v2;}
' INT16 ENTRYPOINT INT8modINT8(INT8 v1,INT8 v2){return v1 % v2;}
' INT16 ENTRYPOINT INT8mulINT8(INT8 v1,INT8 v2){return v1 * v2;}
' INT64 ENTRYPOINT INT8powINT8(INT8 v1,INT8 v2){return pow((double)v1 , (double)v2);}

' INT16 ENTRYPOINT toINT16(INT16 Value){return Value;}
' INT32 ENTRYPOINT INT16plusINT16(INT16 v1,INT16 v2){return v1 + v2;}
' INT32 ENTRYPOINT INT16minusINT16(INT16 v1,INT16 v2){return v1 - v2;}
' INT32 ENTRYPOINT INT16intsubINT16(INT16 v1,INT16 v2){return v1 / v2;}
' INT32 ENTRYPOINT INT16subINT16(INT16 v1,INT16 v2){return (double)v1 / (double)v2;}
' INT32 ENTRYPOINT INT16modINT16(INT16 v1,INT16 v2){return v1 % v2;}
' INT32 ENTRYPOINT INT16mulINT16(INT16 v1,INT16 v2){return v1 * v2;}
' INT64 ENTRYPOINT INT16powINT16(INT16 v1,INT16 v2){return pow((double)v1 , (double)v2);}

' INT32 ENTRYPOINT toINT32(INT32 Value){return Value;}
' INT64 ENTRYPOINT INT32plusINT32(INT32 v1,INT32 v2){return v1 + v2;}
' INT64 ENTRYPOINT INT32minusINT32(INT32 v1,INT32 v2){return v1 - v2;}
' INT64 ENTRYPOINT INT32intsubINT32(INT32 v1,INT32 v2){return v1 / v2;}
' INT64 ENTRYPOINT INT32subINT32(INT32 v1,INT32 v2){return (double)v1 / (double)v2;}
' INT64 ENTRYPOINT INT32modINT32(INT32 v1,INT32 v2){return v1 % v2;}
' INT64 ENTRYPOINT INT32mulINT32(INT32 v1,INT32 v2){return v1 * v2;}
' INT64 ENTRYPOINT INT32powINT32(INT32 v1,INT32 v2){return pow((double)v1 , (double)v2);}

' INT64 ENTRYPOINT toINT64(INT64 Value){return Value;}
' INT64 ENTRYPOINT INT64plusINT64(INT64 v1,INT64 v2){return v1 +v2;}
' INT64 ENTRYPOINT INT64minusINT64(INT64 v1,INT64 v2){return v1 - v2;}
' INT64 ENTRYPOINT INT64intsubINT64(INT64 v1,INT64 v2){return v1 / v2;}
' INT64 ENTRYPOINT INT64subINT64(INT64 v1,INT64 v2){return (double)v1 / (double)v2;}
' INT64 ENTRYPOINT INT64modINT64(INT64 v1,INT64 v2){return v1 % v2;}
' INT64 ENTRYPOINT INT64mulINT64(INT64 v1,INT64 v2){return v1 * v2;}
' INT64 ENTRYPOINT INT64powINT64(INT64 v1,INT64 v2){return pow((double)v1 , (double)v2);}