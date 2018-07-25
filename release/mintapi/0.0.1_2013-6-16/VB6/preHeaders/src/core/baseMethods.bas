Attribute VB_Name = "baseMethods"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "baseMethods"


Private Const LOCALE_SDECIMAL As Long = &HE
Private Const LOCALE_SLONGDATE As Long = &H20
Private Const LOCALE_SSHORTDATE As Long = &H1F
Private Const LOCALE_SCURRENCY As Long = &H14
Private Const LOCALE_STHOUSAND As Long = &HF
Private Const LOCALE_SINTLSYMBOL As Long = &H15
Private Const LOCALE_STIMEFORMAT As Long = &H1003
Private Const LOCALE_SPOSITIVESIGN = &H50
Private Const LOCALE_SNEGATIVESIGN = &H51
Private Const LOCALE_SCOUNTRY As Long = &H6
Private Const LOCALE_SDAYNAME1 As Long = &H2A
Private Const LOCALE_SDAYNAME2 As Long = &H2B
Private Const LOCALE_SDAYNAME3 As Long = &H2C
Private Const LOCALE_SDAYNAME4 As Long = &H2D
Private Const LOCALE_SDAYNAME5 As Long = &H2E
Private Const LOCALE_SDAYNAME6 As Long = &H2F
Private Const LOCALE_SDAYNAME7 As Long = &H30
Private Const LOCALE_SENGCOUNTRY As Long = &H1002
Private Const LOCALE_SENGLANGUAGE As Long = &H1001
Private Const LOCALE_SLANGUAGE As Long = &H2
Private Const LOCALE_SMONTHNAME1 As Long = &H38
Private Const LOCALE_SMONTHNAME10 As Long = &H41
Private Const LOCALE_SMONTHNAME11 As Long = &H42
Private Const LOCALE_SMONTHNAME12 As Long = &H43
Private Const LOCALE_SMONTHNAME2 As Long = &H39
Private Const LOCALE_SMONTHNAME3 As Long = &H3A
Private Const LOCALE_SMONTHNAME4 As Long = &H3B
Private Const LOCALE_SMONTHNAME5 As Long = &H3C
Private Const LOCALE_SMONTHNAME6 As Long = &H3D
Private Const LOCALE_SMONTHNAME7 As Long = &H3E
Private Const LOCALE_SMONTHNAME8 As Long = &H3F
Private Const LOCALE_SMONTHNAME9 As Long = &H40
Private Const LOCALE_SABBREVCTRYNAME = &H7
Private Const LOCALE_SABBREVDAYNAME1 = &H31
Private Const LOCALE_SABBREVDAYNAME3 = &H33
Private Const LOCALE_SABBREVDAYNAME2 = &H32
Private Const LOCALE_SABBREVDAYNAME4 = &H34
Private Const LOCALE_SABBREVDAYNAME5 = &H35
Private Const LOCALE_SABBREVDAYNAME6 = &H36
Private Const LOCALE_SABBREVDAYNAME7 = &H37
Private Const LOCALE_SABBREVLANGNAME = &H3
Private Const LOCALE_SABBREVMONTHNAME1 = &H44
Private Const LOCALE_SABBREVMONTHNAME10 = &H4D
Private Const LOCALE_SABBREVMONTHNAME11 = &H4E
Private Const LOCALE_SABBREVMONTHNAME12 = &H4F
Private Const LOCALE_SABBREVMONTHNAME13 = &H100F
Private Const LOCALE_SABBREVMONTHNAME2 = &H45
Private Const LOCALE_SABBREVMONTHNAME3 = &H46
Private Const LOCALE_SABBREVMONTHNAME4 = &H47
Private Const LOCALE_SABBREVMONTHNAME5 = &H48
Private Const LOCALE_SABBREVMONTHNAME6 = &H49
Private Const LOCALE_SABBREVMONTHNAME7 = &H4A
Private Const LOCALE_SABBREVMONTHNAME8 = &H4B
Private Const LOCALE_SABBREVMONTHNAME9 = &H4C
Private Const LOCALE_SNATIVECTRYNAME = &H8
Private Const LOCALE_SNATIVELANGNAME = &H4

Private Const LOCALE_USER_DEFAULT As Long = &H400

Public Type BASEMETHODS_SAFEPATH_COLUMN
    Value As String
    Include As Boolean
End Type
Public Type BASEMETHODS_SAFEPATH
    cols() As BASEMETHODS_SAFEPATH_COLUMN
    colsCount As Long
End Type

Public Type API_OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Declare Function API_VarPtr Lib "msvbvm60" Alias "VarPtr" (Ptr As Any) As Long
Private Declare Function API_VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub API_baseMethods_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function API_baseMethods_GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As API_OSVERSIONINFO) As Long
Private Declare Function API_baseMethods_GetVersion Lib "kernel32" Alias "GetVersion" () As Long
Private Declare Function API_baseMethods_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function API_baseMethods_GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function API_baseMethods_GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function API_baseMethods_GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function API_baseMethods_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function API_baseMethods_SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function API_baseMethods_GetUserDefaultLangID Lib "kernel32" Alias "GetUserDefaultLangID" () As Integer
Private Declare Function API_baseMethods_GetSystemDefaultLCID Lib "kernel32" Alias "GetSystemDefaultLCID" () As Long
Private Declare Function API_baseMethods_GetUserDefaultLCID Lib "kernel32" Alias "GetUserDefaultLCID" () As Long
Private Declare Function API_baseMethods_SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function API_baseMethods_GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectory" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Declare Function API_baseMethods_CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
'Private Declare Function API_baseMethods_CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function API_baseMethods_RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function API_baseMethods_GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function API_baseMethods_GetUserProfileDirectory Lib "userenv" (ByVal hToken As Long, ByVal lpProfileDir As String, ByRef lpcchSize As Long) As Long
Private Declare Function API_baseMethods_GetSiteDirectory Lib "advapi32" Alias "GetSiteDirectoryA" (ByVal hToken As Long, ByVal pszSiteDirectory As String, ByVal uSize As Long) As Long
Private Declare Function API_baseMethods_GetSystemWindowsDirectory Lib "kernel32" Alias "GetSystemWindowsDirectoryA" (ByVal lpBuffer As String, ByVal uSize As Long) As Long
Private Declare Function API_baseMethods_GetProfilesDirectory Lib "userenv" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
'Private Declare Function API_baseMethods_GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long


Private Declare Function API_baseMethods_SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Private Declare Function API_baseMethods_GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal lLocale As Long, ByVal lLocaleType As Long, ByVal sLCData As String, ByVal lBufferLength As Long) As Long
Private Declare Function API_baseMethods_GetSystemDefaultLangID Lib "kernel32" Alias "GetSystemDefaultLangID" () As Integer
Private Declare Function API_baseMethods_VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long

Public Enum ConcatPathEndMode
    cpeEndsWithSlash
    cpeNoEndsWithSlash
    cpeNotMatter
End Enum

Public Enum API_PlatformID
    pltID_Win32s = 0
    pltID_Win95_98 = 1
    pltID_WinNT = 2
End Enum
Public Enum API_OSVersion
    OSV_DOS = 0
    OSV_Win3X = 1
    OSV_Win95 = 2
    OSV_Win98 = 3
    OSV_WinNT = 4
    OSV_WinNT3X = OSV_WinNT
    OSV_WinNT40 = OSV_WinNT
    OSV_Win2000 = 5
    OSV_WinXP = OSV_Win2000
    OSV_WinServer2003 = OSV_WinXP
    OSV_WinVista = 6
    OSV_Win7 = 7
    OSV_WinServer2008 = OSV_Win7
    OSV_Win8 = 8
    OSV_Win9 = 9
    OSV_Higher = &H80
    OSV_Unknown = &H7F
End Enum

Public Enum API_equalToFlags
    API_AllMustEqual = 1
    API_SomeEqual = 2
    API_SomeNotEqual = 4
    API_AllMustNotEqual = 8

    API_SomeCompare = API_SomeEqual Or API_SomeNotEqual
    API_AllCompare = API_AllMustEqual Or API_AllMustNotEqual
    API_NotValue = 32
    API_AllValues = API_AllMustEqual Or API_SomeEqual Or API_SomeNotEqual Or API_AllMustNotEqual
End Enum


Dim inited As Boolean
Dim bmVersionInfoRecord As API_OSVERSIONINFO

Public Sub Initialize()
    If inited Then Exit Sub
    Call baseConstants.Initialize
    Call baseExceptions.Initialize
    
    
    bmVersionInfoRecord.dwOSVersionInfoSize = Len(bmVersionInfoRecord)
    Call API_baseMethods_GetVersionEx(bmVersionInfoRecord)
    
    
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    inited = False
End Sub

Public Function ArraySize(targetArray) As Long
If Not (VarType(targetArray) And vbArray) = vbArray Then throw InvalidArgumentTypeException
    On Error GoTo zeroLength
    ArraySize = (UBound(targetArray) - LBound(targetArray) + 1)
zeroLength:
End Function
Public Function IsArrayEmpty(targetArray) As Boolean
If Not (VarType(targetArray) And vbArray) = vbArray Then throw InvalidArgumentTypeException
    On Error GoTo zeroLength
    IsArrayEmpty = (UBound(targetArray) - LBound(targetArray) + 1) <= 0
    Exit Function
zeroLength:
    IsArrayEmpty = True
End Function
Public Sub EmptyArray(targetArray)
If Not (VarType(targetArray) And vbArray = vbArray) Then throw InvalidArgumentTypeException
    Erase targetArray
End Sub
Public Function IsEmptyVariable(targetVariable) As Boolean
    Select Case VarType(targetVariable)
        Case vbArray
            IsEmptyVariable = (IsArrayEmpty(targetVariable))
        Case VBObject
            IsEmptyVariable = (targetVariable Is Nothing)
        Case Else
            IsEmptyVariable = (targetVariable = Empty)
    End Select
End Function

Public Function new_(Arg)
    If VarType(Arg) = VBObject Then
        Set new_ = Arg
    Else
            new_ = Arg
    End If
End Function

Public Function GetLPSTR(lpStr As String) As String
    Dim Index As Long
    Index = InStr(1, lpStr, Chr(0)) - 1
    If Index <= 0 Then Exit Function
    GetLPSTR = Left(lpStr, Index)
End Function
Public Function GetLPSTRWL(lpStr As String, Length As Long) As String 'With Length
    GetLPSTRWL = Left(lpStr, Length)
End Function

'Check If EnumerateValue Has One Of The flags Conditions When Comparing It To args.
'ex: if equalTo(SomeEqual,x,10,29,4) then //means if x equal to one of the 10,29,4 returns true else false
'flags:Compaire Mode.
'EnumerateValue:Value To Compaire.
'args:Array To Compaire With EnumerateValue.
'returns:bool
Public Function equalTo(Flags As API_equalToFlags, EnumerateValue, ParamArray Args() As Variant) As Boolean
    Dim i As Long
    On Error GoTo errZeroArgs
    i = LBound(Args)
    GoTo notZeroArgs
errZeroArgs:
    If (Flags And API_AllMustNotEqual = API_AllMustNotEqual) Or (Flags And API_SomeNotEqual = API_SomeNotEqual) Then
        equalTo = (EnumerateValue <> 0)
    Else
        equalTo = (EnumerateValue = 0)
    End If
    Exit Function
notZeroArgs:
    On Error Resume Next
    Select Case VarType(EnumerateValue)
    Case VbVarType.vbArray 'VarType(EnumerateValue)
        Select Case Flags And API_equalToFlags.API_AllValues
            Case API_AllMustEqual: For i = LBound(Args) To UBound(Args)
                If Not ArrayCompare(EnumerateValue, Args(i)) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case API_AllMustNotEqual: For i = LBound(Args) To UBound(Args)
                If ArrayCompare(EnumerateValue, Args(i)) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case API_SomeEqual: For i = LBound(Args) To UBound(Args)
                If ArrayCompare(EnumerateValue, Args(i)) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case API_SomeNotEqual: For i = LBound(Args) To UBound(Args)
                If Not ArrayCompare(EnumerateValue, Args(i)) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case Else: throw UnknownValueException("Unknown Flags argument value.")
        End Select 'Flags And equalToFlags.AllValues
    Case VbVarType.VBObject 'VarType(EnumerateValue)
        Select Case Flags And API_equalToFlags.API_AllValues
            Case API_AllMustEqual: For i = LBound(Args) To UBound(Args)
                If Not EnumerateValue Is Args(i) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case API_AllMustNotEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue Is Args(i) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case API_SomeEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue Is Args(i) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case API_SomeNotEqual: For i = LBound(Args) To UBound(Args)
                If Not EnumerateValue Is Args(i) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case Else: throw UnknownValueException("Unknown Flags argument value.")
        End Select 'Flags And equalToFlags.AllValues
    Case Else 'VarType(EnumerateValue)
        Select Case Flags And API_AllValues
            Case AllMustEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue <> CLng(Args(i)) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case API_AllMustNotEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue = CLng(Args(i)) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case API_SomeEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue = CLng(Args(i)) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case API_SomeNotEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue <> CLng(Args(i)) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case Else: throw UnknownValueException("Unknown Flags argument value.")
        End Select 'Flags And equalToFlags.AllValues
    End Select 'VarType(EnumerateValue)
    If Flags And API_NotValue = API_NotValue Then equalTo = Not equalTo
End Function
Public Function equalToArr(Flags As API_equalToFlags, EnumerateValue, Args() As Variant) As Boolean
    Dim i As Long
    On Error GoTo errZeroArgs
    i = LBound(Args)
    GoTo notZeroArgs
errZeroArgs:
    If (Flags And API_AllMustNotEqual = API_AllMustNotEqual) Or (Flags And API_SomeNotEqual = API_SomeNotEqual) Then
        equalToArr = (EnumerateValue <> 0)
    Else
        equalToArr = (EnumerateValue = 0)
    End If
    Exit Function
notZeroArgs:
    On Error Resume Next
    Select Case VarType(EnumerateValue)
    Case VbVarType.vbArray 'VarType(EnumerateValue)
        Select Case Flags And API_AllValues
            Case API_AllMustEqual: For i = LBound(Args) To UBound(Args)
                If Not ArrayCompare(EnumerateValue, Args(i)) Then: equalToArr = False: Exit Function
            Next: equalToArr = True
            Case API_AllMustNotEqual: For i = LBound(Args) To UBound(Args)
                If ArrayCompare(EnumerateValue, Args(i)) Then: equalToArr = False: Exit Function
            Next: equalToArr = True
            Case API_SomeEqual: For i = LBound(Args) To UBound(Args)
                If ArrayCompare(EnumerateValue, Args(i)) Then: equalToArr = True: Exit Function
            Next: equalToArr = False
            Case API_SomeNotEqual: For i = LBound(Args) To UBound(Args)
                If Not ArrayCompare(EnumerateValue, Args(i)) Then: equalToArr = True: Exit Function
            Next: equalToArr = False
            Case Else: throw UnknownValueException("Unknown Flags argument value.")
        End Select 'Flags And equalToArrFlags.AllValues
    Case VbVarType.VBObject 'VarType(EnumerateValue)
        Select Case Flags And API_AllValues
            Case API_AllMustEqual: For i = LBound(Args) To UBound(Args)
                If Not EnumerateValue Is Args(i) Then: equalToArr = False: Exit Function
            Next: equalToArr = True
            Case API_AllMustNotEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue Is Args(i) Then: equalToArr = False: Exit Function
            Next: equalToArr = True
            Case API_SomeEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue Is Args(i) Then: equalToArr = True: Exit Function
            Next: equalToArr = False
            Case API_SomeNotEqual: For i = LBound(Args) To UBound(Args)
                If Not EnumerateValue Is Args(i) Then: equalToArr = True: Exit Function
            Next: equalToArr = False
            Case Else: throw UnknownValueException("Unknown Flags argument value.")
        End Select 'Flags And equalToArrFlags.AllValues
    Case Else 'VarType(EnumerateValue)
        Select Case Flags And API_AllValues
            Case API_AllMustEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue <> CLng(Args(i)) Then: equalToArr = False: Exit Function
            Next: equalToArr = True
            Case API_AllMustNotEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue = CLng(Args(i)) Then: equalToArr = False: Exit Function
            Next: equalToArr = True
            Case API_SomeEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue = CLng(Args(i)) Then: equalToArr = True: Exit Function
            Next: equalToArr = False
            Case API_SomeNotEqual: For i = LBound(Args) To UBound(Args)
                If EnumerateValue <> CLng(Args(i)) Then: equalToArr = True: Exit Function
            Next: equalToArr = False
            Case Else: throw UnknownValueException("Unknown Flags argument value.")
        End Select 'Flags And equalToArrArrFlags.AllValues
    End Select 'VarType(EnumerateValue)
    If Flags And API_NotValue = API_NotValue Then equalToArr = Not equalToArr
End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
Public Function GetVersion() As Long
    GetVersion = API_baseMethods_GetVersion
End Function
Public Function GetVersionEx() As API_OSVERSIONINFO
    GetVersionEx = bmVersionInfoRecord
End Function

Public Function IsWow64Environment() As Boolean
    
End Function

Public Function OSVersion() As API_OSVersion
    On Error GoTo ErrorHandler
    Dim mjVersion As Long, mnVersion As Long, pltfmID As Long
    mjVersion = bmVersionInfoRecord.dwMajorVersion
    mnVersion = bmVersionInfoRecord.dwMinorVersion
    pltfmID = bmVersionInfoRecord.dwPlatformId
    Select Case mjVersion
        Case 1
            
        Case 2
            
        Case 3
            If pltfmID = pltID_Win32s Then
                OSVersion = OSV_Win3X
            ElseIf pltfmID = pltID_WinNT Then
                OSVersion = OSV_WinNT40
            Else
                OSVersion = OSV_WinNT
            End If
        Case 4
            If pltfmID = pltID_WinNT Then
                OSVersion = OSV_WinNT40
            Else
                OSVersion = IIf(mnVersion = 0, OSV_Win95, OSV_Win98)
            End If
        Case 5
            OSVersion = IIf(mnVersion = 0, OSV_Win2000, OSV_WinXP)
        Case 6
            OSVersion = IIf(mnVersion = 0, OSV_WinVista, OSV_Win7)
        Case Else
            OSVersion = OSV_Unknown
    End Select
Exit Function
ErrorHandler:
    OSVersion = OSV_Unknown
End Function

Public Function CurrentUser() As String
    Dim buf As String * SMALLLPSTR, bufSize As Long
    bufSize = SMALLLPSTR
    buf = String(SMALLLPSTR, Chr(0))
    Call API_baseMethods_GetUserName(buf, bufSize)
    CurrentUser = GetLPSTR(buf)
End Function
Public Function GetTempPath() As String
    Dim buf As String * LARGELPSTR, bufSize As Long
    bufSize = LARGELPSTR
    buf = String(LARGELPSTR, Chr(0))
    Call API_baseMethods_GetTempPath(bufSize, buf)
    GetTempPath = GetLPSTR(buf)
End Function
Public Function GetSystemPath() As String
    Dim buf As String * LARGELPSTR, bufSize As Long
    bufSize = LARGELPSTR
    buf = String(LARGELPSTR, Chr(0))
    Call API_baseMethods_GetSystemDirectory(buf, bufSize)
    GetSystemPath = GetLPSTR(buf)
End Function
Public Function GetWindowsPath() As String
    Dim buf As String * LARGELPSTR, bufSize As Long
    bufSize = LARGELPSTR
    buf = String(LARGELPSTR, Chr(0))
    Call API_baseMethods_GetWindowsDirectory(buf, bufSize)
    GetWindowsPath = GetLPSTR(buf)
End Function
Public Function GetUserPath() As String
    Dim buf As String * LARGELPSTR, bufSize As Long
    bufSize = LARGELPSTR
    buf = String(LARGELPSTR, Chr(0))
    Call API_baseMethods_GetProfilesDirectory(buf, bufSize)
    GetUserPath = GetLPSTR(buf)
End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
Public Function EnumDirectoryFiles(Path As String, Optional Attributes As VbFileAttribute) As String()
    Dim CP As String, Count As Long, retVal() As String
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
    CP = Dir(ConcatPath(Path, "", cpeEndsWithSlash), vbNormal Or Attributes)
    While CP <> ""
        ReDim Preserve retVal(Count)
        retVal(Count) = CP
        Count = Count + 1
        CP = Dir
    Wend
    EnumDirectoryFiles = retVal
End Function
Public Function CountDirectoryFiles(Path As String, Optional Attributes As VbFileAttribute) As Long
    Dim CP As String, Count As Long
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
    CP = Dir(ConcatPath(Path, "", cpeEndsWithSlash), Attributes)
    While CP <> ""
        Count = Count + 1
        CP = Dir
    Wend
    CountDirectoryFiles = Count
End Function
Public Function EnumDirectoryFolders(Path As String, Optional Attributes As VbFileAttribute = 0) As String()
    Dim CP As String, Count As Long, retVal() As String
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
    CP = Dir(Path, vbDirectory Or Attributes)
    If CP <> "" Then CP = Dir
    If CP <> "" Then CP = Dir
    While CP <> ""
        On Local Error GoTo cantOpenFile
        If (GetAttr(ConcatPath(Path, CP)) And vbDirectory) = vbDirectory Then
            On Error GoTo 0
            ReDim Preserve retVal(Count)
            retVal(Count) = CP
            Count = Count + 1
        End If
cantOpenFile:
        CP = Dir
    Wend
    EnumDirectoryFolders = retVal
End Function
Public Function CountDirectoryFolders(Path As String, Optional Attributes As VbFileAttribute = 0) As Long
    Dim CP As String, Count As Long
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
    CP = Dir(Path, vbDirectory Or Attributes)
    If CP <> "" Then CP = Dir
    If CP <> "" Then CP = Dir
    While CP <> ""
        On Local Error GoTo cantOpenFile
        If (GetAttr(ConcatPath(Path, CP)) And vbDirectory) = vbDirectory Then
            On Error GoTo 0
            Count = Count + 1
        End If
cantOpenFile:
        CP = Dir
    Wend
    CountDirectoryFolders = Count
End Function
Public Function CountSubDirectoriesFiles(Path As String, Optional Attributes As VbFileAttribute = 0) As Long
    Dim Count As Long, subDirs() As String, subCount As Long
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
    subDirs = EnumDirectoryFolders(Path, Attributes)
    subCount = ArraySize(subDirs)
    Dim i As Long
    Count = CountDirectoryFiles(Path, Attributes)
    For i = 0 To subCount - 1
        Count = Count + CountSubDirectoriesFiles(ConcatPath(Path, subDirs(i), cpeEndsWithSlash), Attributes)
    Next
    CountSubDirectoriesFiles = Count
End Function
Public Function EnumSubDirectoriesFiles(Path As String, Optional Attributes As VbFileAttribute = 0) As String()
    Dim subDirs() As String, subCount As Long, retVal() As String, bufVal() As String
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
    subDirs = EnumDirectoryFolders(Path, Attributes)
    subCount = ArraySize(subDirs)
    Dim i As Long
    retVal = EnumDirectoryFiles(Path, Attributes)
    For i = 0 To subCount - 1
        bufVal = EnumSubDirectoriesFiles(ConcatPath(Path, subDirs(i), cpeEndsWithSlash), Attributes)
        Call AppendArrayToArray(retVal, bufVal)
    Next
    EnumSubDirectoriesFiles = retVal()
End Function
Public Function GetFilePath(ByVal Path As String) As String
On Error GoTo Err
    Dim slashIndex As Long, backslashIndex As Long
    slashIndex = InStrRev(Path, "/")
    backslashIndex = InStrRev(Path, "\")
    slashIndex = IIf(slashIndex >= backslashIndex, slashIndex, backslashIndex)
    If slashIndex = 0 Then throw InvalidPathException
    backslashIndex = Len(Path)
    GetFilePath = Left(Path, backslashIndex - (backslashIndex - slashIndex + 1))
Exit Function
Err:
End Function
Public Function GetFileName(ByVal Path As String) As String
On Error GoTo Err
    Dim slashIndex As Long, backslashIndex As Long
    slashIndex = InStrRev(Path, "/")
    backslashIndex = InStrRev(Path, "\")
    slashIndex = IIf(slashIndex >= backslashIndex, slashIndex, backslashIndex)
    If slashIndex = 0 Then throw InvalidPathException
    backslashIndex = Len(Path)
    GetFileName = Right(Path, backslashIndex - slashIndex)
Exit Function
Err:
End Function
Public Function GetFileExtention(ByVal Path As String) As String
On Error GoTo Err
    Dim slashIndex As Long, backslashIndex As Long
    slashIndex = InStrRev(Path, "/")
    backslashIndex = InStrRev(Path, "\")
    slashIndex = IIf(slashIndex >= backslashIndex, slashIndex, backslashIndex)
    'If slashIndex = 0 Then throw InvalidPathException
    backslashIndex = InStrRev(Path, ".")
    If backslashIndex = 0 Then
        GetFileExtention = ""
        Exit Function
    End If
    If slashIndex > backslashIndex Then
        GetFileExtention = ""
        Exit Function
    Else
        slashIndex = Len(Path)
        GetFileExtention = Right(Path, slashIndex - backslashIndex)
    End If
Exit Function
Err:
End Function
Public Function GetFileNameOnly(ByVal Path As String) As String
On Error GoTo Err
    Dim slashIndex As Long, backslashIndex As Long, fLen As Long
    fLen = Len(Path)
    slashIndex = InStrRev(Path, "/")
    backslashIndex = InStrRev(Path, "\")
    slashIndex = IIf(slashIndex >= backslashIndex, slashIndex, backslashIndex)
    If slashIndex = 0 Then throw InvalidPathException
    backslashIndex = InStrRev(Path, ".")
    If backslashIndex = 0 Then
        GetFileNameOnly = Right(Path, fLen - slashIndex)
        Exit Function
    End If
    If slashIndex > backslashIndex Then
        GetFileNameOnly = Right(Path, fLen - slashIndex)
        Exit Function
    Else
        GetFileNameOnly = mID(Path, slashIndex + 1, backslashIndex - slashIndex - 1)
    End If
Exit Function
Err:
End Function
Public Function ConcatPath(ByVal Path As String, ByVal PathToAdd As String, Optional EndWithSlash As ConcatPathEndMode = ConcatPathEndMode.cpeNotMatter, Optional Slash As String = "/") As String
    Dim p As String, A As String
    A = PathToAdd
    p = Path
    If A <> "" Then
        If (Left(A, 1) = "/") And (Left(A, 1) = "\") Then
            A = mID(A, 2)
        End If
    Else
        ConcatPath = p
        GoTo checkLastSlash
    End If
    If p <> "" Then
        If Not ((Right(p, 1) = "/") And (Right(p, 1) = "\")) Then
            p = p & Slash
        End If
    Else
        ConcatPath = A
        GoTo checkLastSlash
    End If
    ConcatPath = p & A
    Call CheckPathValidation(ConcatPath, True, False)
checkLastSlash:
    If Not equalTo(SomeEqual, EndWithSlash, ConcatPathEndMode.cpeNoEndsWithSlash, ConcatPathEndMode.cpeEndsWithSlash) Then Exit Function
    If ConcatPath <> "" Then
        If Right(ConcatPath, 1) = Slash Then
            If EndWithSlash = cpeNoEndsWithSlash Then ConcatPath = mID(ConcatPath, 1, Len(ConcatPath) - 1)
        Else
            If EndWithSlash = cpeEndsWithSlash Then ConcatPath = ConcatPath & Slash
        End If
    End If
End Function
Public Function CheckPathValidation(Path As String, MakeTrueForm As Boolean, Optional CheckForDrive As Boolean = True) As Boolean
    If Len(Trim(Path)) = 0 Then Exit Function
    Dim charsIndex As Long
    charsIndex = InStr(1, Path, "*") + InStr(1, Path, "?") + InStr(1, Path, """") + InStr(1, Path, "<") + InStr(1, Path, ">")
    If charsIndex > 0 Then
        CheckPathValidation = False
        Exit Function
    End If
    If CheckForDrive Then
        If (Left(Path, 2) Like "?:") Then
            If Len(Path) > 2 Then
                If Not (Left(Path, 3) Like "?:[/,\,|]") Then
                    CheckPathValidation = False
                    Exit Function
                End If
            End If
        Else
            CheckPathValidation = False
            Exit Function
        End If
    End If
    If MakeTrueForm Then
        Path = RemovePathBads(Path, CheckForDrive)
    End If
    CheckPathValidation = True
End Function
Public Function RemovePathBads(ByVal Path As String, Optional CheckForDrive As Boolean = True) As String
    Dim DriveLetter As String * 1, IsDriveLetter As Boolean
    Path = Trim(Path)
    Path = Replace(Path, "\", "/")
    Path = Replace(Path, "|", "/")
    IsDriveLetter = False
    If (Left(Path, 2) Like "?:") Then
        DriveLetter = Left(Path, 1)
        IsDriveLetter = True
        If Len(Path) >= 3 Then
            If (Left(Path, 3) Like "?:/") Then
                On Error Resume Next
                If Len(Path) = 3 Then
                    RemovePathBads = Path
                    Exit Function
                End If
                Path = mID(Path, 4)
            Else
                throw InvalidPathException
            End If
        Else
            RemovePathBads = Path
            Exit Function
        End If
    Else
        If CheckForDrive Then throw InvalidPathException
    End If

'    Dim strBuff As String
'    strBuff = String(LARGELPSTR, Chr(0))
'    If API_baseMethods_GetFullPathName(Path, LARGELPSTR, strBuff, "") <> 0 Then
'        Path = GetLPSTR(strBuff)
'        MsgBox Path
'    End If
    
    Dim cols As BASEMETHODS_SAFEPATH
    cols = SplitPathToSafePath(Path)

    Dim i As Long, buf As String, doubleDot As Long
    doubleDot = 0
    For i = cols.colsCount - 1 To 0 Step -1
        buf = Trim(cols.cols(i).Value)
        If buf = "" Or buf = "." Then
            cols.cols(i).Include = False
        ElseIf buf = ".." Then
            cols.cols(i).Include = False
            doubleDot = doubleDot + 1
        ElseIf doubleDot > 0 Then
            cols.cols(i).Include = False
            doubleDot = doubleDot - 1
        Else
            cols.cols(i).Include = True
        End If
    Next

    Path = ""
    For i = 0 To cols.colsCount - 1
        If cols.cols(i).Include Then
            Path = Path & "/" & cols.cols(i).Value
        End If
    Next

    On Error Resume Next

    If Len(Path) > 1 Then _
        Path = mID(Path, 2) ' removes first /

    If Not IsDriveLetter Then
        RemovePathBads = Path
    Else
        RemovePathBads = DriveLetter & ":/" & Path  'this also add \ character between driveLetter and path.
    End If
End Function
Public Function SplitPathToSafePath(ByVal Path As String) As BASEMETHODS_SAFEPATH
    Dim cols As BASEMETHODS_SAFEPATH
    Dim strs() As String
    If InStr(1, Path, "|") > 0 Then Path = Replace(Path, "|", "/")
    If InStr(1, Path, "\") > 0 Then Path = Replace(Path, "\", "/")
    strs = Split(Path, "/")
    Dim i As Long, strsCount As Long, cIndex As Long
    strsCount = ArraySize(strs) 'zero based
    If strsCount = 0 Then GoTo zeroLength
    cols.colsCount = strsCount
    ReDim cols.cols(cols.colsCount - 1)
    For i = 0 To strsCount - 1
        cols.cols(i).Value = strs(i)
        cols.cols(i).Include = True
    Next
zeroLength:
    SplitPathToSafePath = cols
End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
Public Sub CopyByteArrayToByteArray(DestinationBA() As Byte, SourceBA() As Byte)
    Call API_baseMethods_CopyMemory(ByVal API_VarPtrArray(DestinationBA), ByVal API_VarPtrArray(SourceBA), 4)
End Sub
'MemoryToByteArray : copies specified memory address value to a bytearray.
'targetAddress : the address of memory to copy to byte array.
'SourceSize : source memory address content size to copy to byte array.
'Times : determines the times that source memory value with length equals to SourceSize copies to byte array.
'IT'S AN UNSAFE METHOD!!
Public Function MemoryToByteArray(ByVal targetAddress As Long, SourceSize As Long, Optional times As Long = 1) As Byte()
    Dim outLen As Long
    outLen = (SourceSize * times) - 1
    If outLen <= 0 Then throw InvalidArgumentValueException
    Dim retVal() As Byte
    ReDim retVal(outLen)

    Dim c_byte_value As Byte

    Dim i As Long
    For i = 0 To outLen
        Call API_baseMethods_CopyMemory(ByVal VarPtr(c_byte_value), ByVal (targetAddress + i), 1)
        retVal(i) = c_byte_value
    Next

    MemoryToByteArray = retVal
End Function
'IT'S AN UNSAFE METHOD!!
Public Sub CopyMemoryToByteArray(ByVal targetAddress As Long, SourceSize As Long, targetByteArray() As Byte, Optional times As Long = 1)
    Dim outLen As Long
    outLen = (SourceSize * times) - 1
    If outLen <= 0 Then throw InvalidArgumentValueException
    'Dim targetByteArray() As Byte
    ReDim targetByteArray(outLen)

    Dim c_byte_value As Byte

    Dim i As Long
    For i = 0 To outLen
        Call API_baseMethods_CopyMemory(ByVal VarPtr(c_byte_value), ByVal (targetAddress + i), 1)
        targetByteArray(i) = c_byte_value
    Next

    'targetByteArray = retVal
End Sub
'ByteArrayToMemory : copies specified byte array to memory target's address.
'targetArray : source array to copy to memory.
'BytesToCopy : number of bytes used to copy byte array.
'IT'S AN UNSAFE METHOD!!
Public Sub ByteArrayToMemory(ByVal targetAddress As Long, targetByteArray() As Byte, BytesToCopy As Long, Optional FillWithNull As Boolean = False)
    If BytesToCopy <= 0 Then Exit Sub
    Dim arrSize As Long
    arrSize = ArraySize(targetByteArray)
    If arrSize < BytesToCopy Then
        If Not FillWithNull Then throw InvalidStatusException("targetByteArray Length Is Less Than BytesToCopy And FillWithNull Is Not Allowed.")
    End If
    Dim i As Long, c_byte_value As Byte
    For i = 0 To BytesToCopy - 1
        If (i < arrSize) Then
            c_byte_value = targetByteArray(i)
        Else
            c_byte_value = 0
        End If
        Call API_baseMethods_CopyMemory(ByVal (targetAddress + i), ByVal VarPtr(c_byte_value), 1)
    Next
End Sub
'IT'S AN UNSAFE METHOD!!

Public Function StringToIntegerArray(str As String, Optional Length As Long = -1) As Integer()
    If Length = 0 Then Exit Function
    Dim strLen As Long
    Dim retVal() As Integer
    strLen = Len(str)
    If strLen = 0 Then Exit Function
    If Length < 0 Then Length = strLen
    ReDim retVal(Length - 1)
    Dim i As Long
    For i = 1 To Length
        retVal(i - 1) = AscW(mID(str, i, 1))
    Next
    StringToIntegerArray = retVal()
End Function
Public Function IntegerArrayToString(intArr() As Integer, Optional Length As Long = -1) As String
    If Length = 0 Then Exit Function
    Dim str As String, arrSize As Long
    arrSize = ArraySize(intArr)
    If arrSize = 0 Then Exit Function
    If Length < 0 Then Length = arrSize
    Dim i As Long, arrlBound As Long
    arrlBound = LBound(intArr)
    For i = 0 To Length - 1
        str = str & ChrW(intArr(i + arrlBound))
    Next
    IntegerArrayToString = str
End Function
Public Function StringToByteArray(str As String, Optional Length As Long = -1) As Byte()
    If Length = 0 Then Exit Function
    Dim strLen As Long
    Dim retVal() As Byte
    strLen = Len(str)
    If strLen = 0 Then Exit Function
    If Length < 0 Then Length = strLen
    ReDim retVal(Length - 1)
    Dim i As Long
    For i = 1 To Length
        retVal(i - 1) = Asc(mID(str, i, 1))
    Next
    StringToByteArray = retVal()
End Function
Public Function ByteArrayToString(B() As Byte, Optional Length As Long = -1) As String
    If Length = 0 Then Exit Function
    Dim str As String, arrSize As Long
    On Error GoTo zeroLength
    arrSize = UBound(B) - LBound(B) + 1
zeroLength:
    If arrSize = 0 Then Exit Function
    If Length < 0 Then Length = arrSize
    Dim i As Long, arrlBound As Long
    arrlBound = LBound(B)
    For i = 0 To Length - 1
        str = str & Chr(B(i + arrlBound))
    Next
    ByteArrayToString = str
End Function
Public Function ByteArrayToSafeString(B() As Byte, Optional Length As Long = -1) As String
    If Length = 0 Then Exit Function
    Dim str As String, arrSize As Long
    On Error GoTo zeroLength
    arrSize = UBound(B) - LBound(B) + 1
zeroLength:
    If arrSize = 0 Then Exit Function
    If Length < 0 Then Length = arrSize
    Dim i As Long, arrlBound As Long, cIndex As Long
    arrlBound = LBound(B)
    For i = 0 To Length - 1
        cIndex = i + arrlBound
        If B(cIndex) = 0 Then Exit For
        str = str & Chr(B(cIndex))
    Next
    ByteArrayToSafeString = str
End Function
Public Function LongToByteArray(lngNum As Long) As Byte()
    LongToByteArray = MemoryToByteArray(VarPtr(lngNum), 4)
End Function
Public Function ByteArrayToLong(B() As Byte) As Long
    Call ByteArrayToMemory(VarPtr(ByteArrayToLong), B, 4)
End Function
Public Function IntegerToByteArray(intNum As Integer) As Byte()
    IntegerToByteArray = MemoryToByteArray(VarPtr(intNum), 2)
End Function
Public Function ByteArrayToInteger(B() As Byte) As Integer
    Call ByteArrayToMemory(API_VarPtr(ByteArrayToInteger), B, 2)
End Function
Public Function ByteToByteArray(btByte As Byte) As Byte()
    ByteToByteArray = MemoryToByteArray(VarPtr(btByte), 1)
End Function
Public Function ByteArrayToByte(B() As Byte) As Byte
    Call ByteArrayToMemory(API_VarPtr(ByteArrayToByte), B, 1)
End Function
Public Function DateToByteArray(dtDate As Date) As Byte()
    DateToByteArray = MemoryToByteArray(VarPtr(dtDate), 8)
End Function
Public Function ByteArrayToDate(B() As Byte) As Date
    Call ByteArrayToMemory(API_VarPtr(ByteArrayToDate), B, 8)
End Function
Public Function CurrencyToByteArray(cyCurrency As Currency) As Byte()
    CurrencyToByteArray = MemoryToByteArray(VarPtr(cyCurrency), 8)
End Function
Public Function ByteArrayToCurrency(B() As Byte) As Currency
    Call ByteArrayToMemory(API_VarPtr(ByteArrayToCurrency), B, 8)
End Function
Public Function DoubleToByteArray(dblNum As Double) As Byte()
    DoubleToByteArray = MemoryToByteArray(VarPtr(dblNum), 8)
End Function
Public Function ByteArrayToDouble(B() As Byte) As Double
    Call ByteArrayToMemory(API_VarPtr(ByteArrayToDouble), B, 8)
End Function
Public Function SingleToByteArray(sngNum As Single) As Byte()
    SingleToByteArray = MemoryToByteArray(VarPtr(sngNum), 4)
End Function
Public Function ByteArrayToSingle(B() As Byte) As Single
    Call ByteArrayToMemory(API_VarPtr(ByteArrayToSingle), B, 4)
End Function
Public Function BooleanToByteArray(boolValue As Boolean) As Byte()
    BooleanToByteArray = MemoryToByteArray(VarPtr(boolValue), 2)
End Function
Public Function ByteArrayToBoolean(B() As Byte) As Boolean
    Call ByteArrayToMemory(API_VarPtr(ByteArrayToBoolean), B, 2)
End Function


Public Function IndexOfByteArray(Source() As Byte, Find() As Byte, Optional StartAt As Long = 0) As Long
    Dim i As Long, sj As Long 'Same j
    Dim sLen As Long, sLB As Long, fLB As Long, fLen As Long
    sLen = ArraySize(Source)
    If sLen <= 0 Then Exit Function
    fLen = ArraySize(Find)
    If fLen <= 0 Then Exit Function
    sLB = LBound(Source)
    fLB = LBound(Find)
    
    Dim areSame As Boolean
    
    For i = StartAt To sLen - 1
        areSame = True
        For sj = 0 To fLen - 1
            If (Source(sLB + i + sj) <> Find(fLB + sj)) Then
                areSame = False
                Exit For
            End If
        Next
        If areSame Then
            IndexOfByteArray = i
            Exit Function
        End If
    Next
    IndexOfByteArray = -1
End Function
Public Function LastIndexOfByteArray(Source() As Byte, Find() As Byte, Optional StartAt As Long = 0) As Long
    Dim i As Long, sj As Long 'Same j
    Dim sLen As Long, sLB As Long, fLB As Long, fLen As Long
    sLen = ArraySize(Source)
    If sLen <= 0 Then Exit Function
    fLen = ArraySize(Find)
    If fLen <= 0 Then Exit Function
    sLB = LBound(Source)
    fLB = LBound(Find)
    
    Dim areSame As Boolean
    
    For i = StartAt To sLen - 1
        areSame = True
        For sj = 0 To fLen - 1
            If (Source(sLB + i + sj) <> Find(fLB + sj)) Then
                areSame = False
                Exit For
            End If
        Next
        If areSame Then
            LastIndexOfByteArray = i
            Exit Function
        End If
    Next
    LastIndexOfByteArray = -1
End Function


Public Sub ByteArrayToArgsArray(B() As Byte, Args() As Variant)
    
End Sub
Public Function ArgsArrayToByteArray(Args() As Variant) As Byte()
    
End Function
Public Function ParamArrayArgsToByteArray(target) As Byte()
    
End Function
Public Sub ByteArrayToParamArrayArgs(targetByteArray() As Byte, target)
    
End Sub

Public Function ArrayToByteArray(targetArray) As Byte()
    
End Function

Public Function GetSubByteArray(targetByteArray() As Byte, ByVal StartIndex As Long, Optional ByVal Length As Long = -1) As Byte()
    If Length = 0 Then Exit Function
    Dim ln As Long, outArr() As Byte, outLen As Long
    ln = ArraySize(targetByteArray)
    If ln <= 0 Then Exit Function
    If StartIndex >= ln Then Exit Function
    If Length = -1 Then
        Length = ln - StartIndex
    ElseIf Length > ln - StartIndex Then
        Length = ln - StartIndex
    End If
    ReDim outArr(Length - 1)
    Dim i As Long, l_Bound As Long, step_outArrIndex As Long
    l_Bound = LBound(targetByteArray) + StartIndex
    outLen = (l_Bound + (Length - 1))
    step_outArrIndex = 0
    For i = l_Bound To outLen
        outArr(step_outArrIndex) = targetByteArray(i)
        step_outArrIndex = step_outArrIndex + 1
    Next
    GetSubByteArray = outArr
End Function
Public Function GetByteArraySomeLength(targetArray() As Byte, Length As Long) As Byte()
If Not ((VarType(targetArray) And (vbArray Or vbByte)) = (vbArray Or vbByte)) Then throw InvalidArgumentTypeException("Only Arrays Accepted.")
    Dim retVal()
    Dim outLen As Long
    
End Function
Public Function IsByteArrayNumeric(bArr() As Byte) As Boolean
    Dim cb As Byte
    Dim i As Long
    For i = 0 To ArraySize(bArr) - 1
        cb = bArr(i)
        If cb <= 48 Or cb >= 57 Then
            IsByteArrayNumeric = False
            Exit Function
        End If
    Next
    IsByteArrayNumeric = True
End Function
Public Function IsByteArrayAlphabetic(bArr() As Byte) As Boolean
    Dim cb As Byte
    Dim i As Long
    For i = 0 To ArraySize(bArr) - 1
        cb = bArr(i)
        If (cb <= 65 Or (cb >= 90 And cb <= 97) Or cb >= 122) Then
            IsByteArrayAlphabetic = False
            Exit Function
        End If
    Next
    IsByteArrayAlphabetic = True
End Function
Public Function IsByteArrayLikeAnother(fA() As Byte, sA() As Byte) As Boolean
    
End Function

Public Function ByteArrayToUpper(bArr() As Byte) As Byte()
    Dim retVal() As Byte
    retVal = bArr
    Dim i As Long
    For i = 0 To ArraySize(retVal) - 1
        If retVal(i) >= 97 And retVal(i) <= 122 Then
            retVal(i) = retVal(i) - 32
        End If
    Next
    ByteArrayToUpper = retVal
End Function
Public Function ByteArrayToLower(bArr() As Byte) As Byte()
    Dim retVal() As Byte
    retVal = bArr
    Dim i As Long
    For i = 0 To ArraySize(retVal) - 1
        If retVal(i) >= 65 And retVal(i) <= 90 Then
            retVal(i) = retVal(i) + 32
        End If
    Next
    ByteArrayToLower = retVal
End Function

Public Function FillArray(targetArray, whattofill, OutputSize As Long, Optional FillToLeft = True)
    If (VarType(targetArray) And vbArray) <> vbArray Then throw InvalidArgumentTypeException
    Dim targetLength As Long
    targetLength = ArraySize(targetArray)

End Function
Public Sub InsertArrayIndex(targetArray, Index, Item)
    Dim arrayType As VbVarType, itemType As VbVarType
    itemType = VarType(Item)
    arrayType = VarType(targetArray)
    If (arrayType And vbArray) <> vbArray Then throw InvalidArgumentTypeException("Only Arrays Accepted.")
    If (arrayType And vbVariant) <> vbVariant Then
        If arrayType <> (itemType Or vbArray) Then _
            throw InvalidArgumentTypeException("Both arguments type must be as one.")
    End If
    
    Dim varCount As Long, i As Long
    varCount = ArraySize(targetArray)
    
    If varCount = 0 Then Exit Sub
    
    If Index > varCount Then throw OutOfRangeException
    
    ReDim Preserve targetArray(varCount)
    If (arrayType And VBObject) = VBObject Then
        For i = varCount To Index + 1 Step -1
            Set targetArray(i) = targetArray(i - 1)
        Next
        Set targetArray(Index) = Item
    ElseIf (arrayType And vbVariant) = vbVariant Then
        If (itemType And VBObject) = VBObject Then
            For i = varCount To Index + 1 Step -1
                Set targetArray(i) = targetArray(i - 1)
            Next
            Set targetArray(Index) = Item
        Else
            For i = varCount To Index + 1 Step -1
                targetArray(i) = targetArray(i - 1)
            Next
            targetArray(Index) = Item
        End If
    Else
        For i = varCount To Index + 1 Step -1
            targetArray(i) = targetArray(i - 1)
        Next
        targetArray(Index) = Item
    End If
End Sub
Public Sub InsertArrayIndexArray(targetArray, Index, insertArray, Optional Length As Long = -1)
    Dim arrayType As VbVarType, insertType As VbVarType
    insertType = VarType(insertArray)
    arrayType = VarType(targetArray)
    If (arrayType And vbArray) <> vbArray Then throw InvalidArgumentTypeException("Only Arrays Accepted.")
    If (arrayType And vbVariant) <> vbVariant Then
        If arrayType <> insertType Then _
            throw InvalidArgumentTypeException("Both arguments type must be as one.")
    End If
    
    Dim varCount As Long, i As Long, sArrCtr As Long
    Dim howMuchToShift As Long, outLength As Long
    varCount = ArraySize(targetArray)
    howMuchToShift = ArraySize(insertArray)
    
    If howMuchToShift = 0 Then Exit Sub
    If varCount = 0 Then Exit Sub
    
    If Index > varCount Then throw OutOfRangeException
    
    If Length <> -1 Then howMuchToShift = Length
    If Length < -1 Then throw OutOfRangeException
    If Length > howMuchToShift Then throw OutOfRangeException
    
    outLength = varCount + howMuchToShift
    outLength = outLength - 1
    
    ReDim Preserve targetArray(outLength)
    If (arrayType And VBObject) = VBObject Then
        For i = outLength To Index + howMuchToShift Step -1
            Set targetArray(i) = targetArray(i - howMuchToShift)
        Next
        For i = 0 To howMuchToShift - 1
            Set targetArray(Index + i) = insertArray(i)
        Next
    ElseIf (arrayType And vbVariant) = vbVariant Then
        If (insertType And vbVariant) = vbVariant Then
            For i = outLength To Index + howMuchToShift Step -1
                targetArray(i) = targetArray(insertType)
            Next
            For i = 0 To howMuchToShift - 1
                insertType = i - howMuchToShift
                arrayType = VarType(insertArray(i))
                If arrayType = VBObject Then
                    Set targetArray(Index + i) = insertArray(i)
                Else
                        targetArray(Index + i) = insertArray(i)
                End If
            Next
        Else
            For i = outLength To Index + howMuchToShift Step -1
                targetArray(i) = targetArray(i - howMuchToShift)
            Next
            For i = 0 To howMuchToShift - 1
                targetArray(Index + i) = insertArray(i)
            Next
        End If
    Else
        For i = outLength To Index + howMuchToShift Step -1
            targetArray(i) = targetArray(i - howMuchToShift)
        Next
        For i = 0 To howMuchToShift - 1
            targetArray(Index + i) = insertArray(i)
        Next
    End If
End Sub
Public Sub AppendToArray(sourceArray, Item)
    Dim sourceUBound As Long
    If (VarType(sourceArray) And vbArray) <> vbArray Then throw InvalidArgumentTypeException
    If Not IsArrayEmpty(sourceArray) Then
        sourceUBound = UBound(sourceArray)
        sourceUBound = sourceUBound + 1
    Else
        sourceUBound = 0
    End If
    ReDim Preserve sourceArray(sourceUBound)
    If VarType(Item) = VBObject Then
        Set sourceArray(sourceUBound) = Item
    Else
            sourceArray(sourceUBound) = Item
    End If
End Sub
Public Sub AppendArrayToArray(sourceArray, targetArray)
    Dim sourceArrayVT As Long
    Dim targetArrayVT As Long
    sourceArrayVT = VarType(sourceArray)
    targetArrayVT = VarType(targetArray)
    If (sourceArrayVT And vbArray) <> vbArray Then throw InvalidArgumentTypeException
    If (targetArrayVT And vbArray) <> vbArray Then
        Call AppendToArray(sourceArray, targetArray)
        Exit Sub
    End If
    If targetArrayVT <> sourceArrayVT Then throw InvalidArgumentTypeException("Both arrays must be as the same type.")
    Dim sourceUBound As Long, sourceLBound As Long
    Dim targetUBound As Long, targetLBound As Long
    If IsArrayEmpty(targetArray) Then
        Exit Sub
    Else
        If IsArrayEmpty(sourceArray) Then
            sourceArray = targetArray
        End If
    End If
    sourceLBound = LBound(sourceArray)
    sourceUBound = UBound(sourceArray)
    targetLBound = LBound(targetArray)
    targetUBound = UBound(targetArray)
    Dim i As Long
    For i = targetLBound To targetUBound
        sourceUBound = sourceUBound + 1
        ReDim Preserve sourceArray(sourceUBound)
        If VarType(targetArray(i)) = VBObject Then
            Set sourceArray(sourceUBound) = targetArray(i)
        Else
                sourceArray(sourceUBound) = targetArray(i)
        End If
    Next
End Sub

Public Function ExpandVArrayToString(targetArray, Optional throwExceptionOnUnknownValues As Boolean = True) As String
    Dim targetArrayType As VbVarType, i As Long, retVal As String
    targetArrayType = VarType(targetArray)
    For i = LBound(targetArray) To UBound(targetArray)
        If (targetArrayType And vbArray) = vbArray Then
            retVal = retVal & ExpandVArrayToString(targetArray(i), throwExceptionOnUnknownValues)
        ElseIf (targetArrayType And VBObject) = VBObject Then
            If throwExceptionOnUnknownValues Then throw InvalidArgumentTypeException("One Or More Array Items Type Is Object.")
        Else
            retVal = retVal & CStr(targetArray(i))
        End If
    Next
    ExpandVArrayToString = retVal
End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
Public Function BinaryCompare(a1, a2, Optional LengthToCompare As Long = -1) As Boolean
    Dim a1Type As VbVarType
    a1Type = VarType(a1)
    If a1Type <> VarType(a2) Then: Exit Function
    If a1Type = vbArray Then
        BinaryCompare = ArrayCompare(a1, a2, LengthToCompare)
    ElseIf a1Type = VBObject Then
        BinaryCompare = (a1 Is a2)
    Else
        BinaryCompare = (a1 = a2)
    End If
End Function
Public Function ArrayCompare(a1, a2, Optional LengthToCompare As Long = -1) As Boolean
    Dim a1Type As VbVarType, a2Type As VbVarType
    a1Type = VarType(a1): a2Type = VarType(a2)
    If Not (((a1Type And vbArray) = vbArray) And ((a2Type And vbArray) = vbArray)) Then throw InvalidArgumentTypeException("Only Arrays Accepted.")
    If (a1Type <> a2Type) Then GoTo some_returnFalse 'throw InvalidArgumentTypeException("a1 Array Type Must Equal To a2 Array Type.")
    Dim a1Len As Long, a2Len As Long
    Dim l1Bound As Long, l2Bound As Long
    On Error GoTo a1_zeroLength
    l1Bound = LBound(a1)
    a1Len = UBound(a1) - l1Bound + 1
a1_zeroLength:
    On Error GoTo a2_zeroLength
    l2Bound = LBound(a2)
    a2Len = UBound(a2) - l2Bound + 1
a2_zeroLength:
    If a1Len = 0 And a2Len = 0 Then
        ArrayCompare = True
        Exit Function
    End If
    If (l1Bound <> l2Bound) Then GoTo some_returnFalse
    Dim ln As Long
    If LengthToCompare = -1 Then
        If a1Len <> a2Len Then GoTo some_returnFalse
        LengthToCompare = a1Len
    End If
    If a1Len < LengthToCompare Then GoTo some_returnFalse
    If a2Len < LengthToCompare Then GoTo some_returnFalse
    ln = LengthToCompare - 1
    On Error GoTo 0
    Dim i As Long
    For i = LBound(a1) To ln
        If a1(i) <> a2(i) Then GoTo some_returnFalse
    Next
    ArrayCompare = True
    Exit Function
some_returnFalse:
    ArrayCompare = False
End Function

Public Function RepeatString(str As String, Count As Long, Optional Splitter As String = "") As String
    Dim i As Long
    For i = 1 To Count
        RepeatString = RepeatString & str & Splitter
    Next
End Function




Public Function GetLocaleString(ByVal lLocaleNum As Long) As String
    'Generic routine to get the locale string from the Operating system.
    Dim lBuffSize As String
    Dim sBuffer As String
    Dim lRet As Long

    lBuffSize = 256
    sBuffer = String(lBuffSize, vbNullChar)

    'Get the information from the registry
    lRet = API_baseMethods_GetLocaleInfo(LOCALE_USER_DEFAULT, lLocaleNum, sBuffer, lBuffSize)
    'If lRet > 0 then success - lret is the size of the string returned
    If lRet > 0 Then
        GetLocaleString = Left$(sBuffer, lRet - 1)
    End If
End Function
Public Sub SetLocaleString(ByVal lLocaleNum As Long, strValue As String)
    Call API_baseMethods_SetLocaleInfo(LOCALE_USER_DEFAULT, lLocaleNum, strValue)
End Sub

Public Function LocaleDateFormat() As String
    ' This function will return the Locale date format for the system. Note that the
    ' returned Year is always formatted to 'YYYY' regardless, to ensure Y2k compliance.
    Dim sDateFormat As String
    On Error GoTo vbErrorHandler
    sDateFormat = GetLocaleString(LOCALE_SSHORTDATE)

    ' Make sure we always have YYYY format for y2k
    If InStr(1, sDateFormat, "YYYY", vbTextCompare) = 0 Then
        Replace sDateFormat, "YY", "YYYY"
    End If
    LocaleDateFormat = sDateFormat
Exit Function
vbErrorHandler:
    throw Exception(Err.Description)
    'err.Raise err.Number, "LocaleSettings GetDateFormat", err.Description
End Function
Public Sub SetLocaleDateFormat(Value As String)
    Call SetLocaleString(LOCALE_SSHORTDATE, Value)
End Sub
Public Function LocaleTimeFormat() As String
    'This function returns the locale's defined Time Format.
    LocaleTimeFormat = GetLocaleString(LOCALE_STIMEFORMAT)
Exit Function
vbErrorHandler:
    throw Exception(Err.Description)
    'err.Raise err.Number, "LocaleSettings GetTimeFormat", err.Description
End Function
Public Sub SetLocaleTimeFormat(Value As String)
    Call SetLocaleString(LOCALE_STIMEFORMAT, Value)
End Sub
Public Function LocaleNumberFormat() As String
' This function returns the Locales defined Decimal Number format
    Dim lBuffLen As Long
    Dim sBuffer As String
    Dim sDecimal As String
    Dim sThousand As String
    Dim LRESULT As Long
    Dim sNumFormat As String

    On Error GoTo vbErrorHandler

    'Setup a buffer to receive the settings
    lBuffLen = 128
    sBuffer = String(lBuffLen, vbNullChar)

    LRESULT = API_baseMethods_GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sBuffer, lBuffLen)
    If LRESULT <= 0 Then Exit Function

    sDecimal = Left$(sBuffer, LRESULT - 1)

    sBuffer = String(lBuffLen, vbNullChar)
    LRESULT = API_baseMethods_GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, sBuffer, lBuffLen)
    If LRESULT <= 0 Then Exit Function

    sThousand = Left$(sBuffer, LRESULT - 1)

    LocaleNumberFormat = ("###" & sThousand & "###" & sDecimal & "######")
Exit Function

vbErrorHandler:
    throw Exception(Err.Description)
    'err.Raise err.Number, "LocaleSettings GetNumberFormat", err.Description
End Function
Public Sub SetLocaleNumberFormat(Value As String)
    'Call SetLocaleString(LOCALE_SSHORTDATE, Value)
End Sub
Public Function LocaleThousandSpecifier() As String
    'This function returns the correct Thousand Specifier for the system Locale
    LocaleThousandSpecifier = GetLocaleString(LOCALE_STHOUSAND)
End Function
Public Sub SetLocaleThousandSpecifier(Value As String)
    Call SetLocaleString(LOCALE_STHOUSAND, Value)
End Sub
Public Function LocaleDecimalSpecifier() As String
    'This function returns the correct Decimal Specifier for the system Locale
    LocaleDecimalSpecifier = GetLocaleString(LOCALE_SDECIMAL)
End Function
Public Sub SetLocaleDecimalSpecifier(Value As String)
    Call SetLocaleString(LOCALE_SDECIMAL, Value)
End Sub
Public Function LocaleCurrencySpecifier() As String
    'This function returns the correct Currency Specifier for the system Locale
    LocaleCurrencySpecifier = GetLocaleString(LOCALE_SCURRENCY)
End Function
Public Sub SetLocaleCurrencySpecifier(Value As String)
    Call SetLocaleString(LOCALE_SCURRENCY, Value)
End Sub
Public Function LocaleSystemLanguageID() As Long
    'Returns the System Language ID for the machine
    LocaleSystemLanguageID = API_baseMethods_GetSystemDefaultLangID
End Function
Public Function LocaleSystemLanguageName() As String
    'Returns the System Language Name eg : English (United Kingdom)
    Dim lLangID As Long
    Dim sBuffer As String
    Dim lBuffSize As Long
    Dim lRet As Long

    On Error GoTo vbErrorHandler

    lLangID = API_baseMethods_GetSystemDefaultLangID
    'Setup a buffer to receive the settings
    lBuffSize = 50
    sBuffer = String(lBuffSize, vbNullChar)
    lRet = API_baseMethods_VerLanguageName(lLangID, sBuffer, lBuffSize)
    If lRet > 0 Then
        LocaleSystemLanguageName = Left$(sBuffer, lRet)
    End If
Exit Function
vbErrorHandler:
    throw Exception(Err.Description)
    'err.Raise err.Number, "LocaleSettings GetSysLanguageName", err.Description
End Function
Public Function LocaleShortMonthName(ByVal iMonthNum As Integer) As String
    'Returns the short-month-name for the specified Month Number
    'eg 1=Jan, 2=Feb (on English machines)
    LocaleShortMonthName = GetLocaleString(LOCALE_SABBREVMONTHNAME1 - 1 + iMonthNum)
End Function
Public Sub SetLocaleShortMonthName(ByVal iMonthNum As Integer, Value As String)
    Call SetLocaleString(LOCALE_SABBREVMONTHNAME1 - 1 + iMonthNum, Value)
End Sub
Public Function LocaleMonthName(ByVal iMonthNum As Integer) As String
    'Returns the Full-Month-Name for the specified month number
    'eg. 1=January, 2=February (on english machines)
    LocaleMonthName = GetLocaleString(LOCALE_SMONTHNAME1 + iMonthNum - 1)
End Function
Public Sub SetLocaleMonthName(ByVal iMonthNum As Integer, Value As String)
    Call SetLocaleString(LOCALE_SMONTHNAME1 - 1 + iMonthNum, Value)
End Sub
Public Function LocaleShortDayName(ByVal iDayNum As Integer) As String
    'Returns the Short-Day-Name for the specified Day Number
    'eg. 1=Mon, 2=Tue (on english machines)
    LocaleShortDayName = GetLocaleString(LOCALE_SABBREVDAYNAME1 + iDayNum - 1)
End Function
Public Sub SetLocaleShortDayName(ByVal iDayNum As Integer, Value As String)
    Call SetLocaleString(LOCALE_SABBREVDAYNAME1 - 1 + iDayNum, Value)
End Sub
Public Function LocaleDayName(ByVal iDayNum As Integer) As String
    'Returns the Full Day Name for the specified Day number
    'eg. 1=Monday, 2=Tuesday (on english machines)
    LocaleDayName = GetLocaleString(LOCALE_SDAYNAME1 + iDayNum - 1)
End Function
Public Sub SetLocaleDayName(ByVal iDayNum As Integer, Value As String)
    Call SetLocaleString(LOCALE_SDAYNAME1 - 1 + iDayNum, Value)
End Sub
Public Function LocaleCountry() As String
    'Returns the Country Name eg. 'United Kingdom'
    LocaleCountry = GetLocaleString(LOCALE_SENGCOUNTRY)
End Function
Public Sub SetLocaleCountry(Value As String)
    Call SetLocaleString(LOCALE_SENGCOUNTRY, Value)
End Sub
Public Function LocaleLanguageName() As String
    'Returns the Native Language Name eg. 'English'
    LocaleLanguageName = GetLocaleString(LOCALE_SNATIVELANGNAME)
End Function
Public Sub SetLocaleLanguageName(Value As String)
    Call SetLocaleString(LOCALE_SNATIVELANGNAME, Value)
End Sub
Public Function LocaleNativeCountryName() As String
    LocaleNativeCountryName = GetLocaleString(LOCALE_SNATIVECTRYNAME)
End Function
Public Sub SetLocaleNativeCountryName(Value As String)
    Call SetLocaleString(LOCALE_SNATIVECTRYNAME, Value)
End Sub
Public Function LocalePositiveSign() As String
    'Returns the symbol used for the positive sign eg. +
    LocalePositiveSign = GetLocaleString(LOCALE_SPOSITIVESIGN)
End Function
Public Sub SetLocalePositiveSign(Value As String)
    Call SetLocaleString(LOCALE_SPOSITIVESIGN, Value)
End Sub
Public Function LocaleNegativeSign() As String
' Returns the symbol used for the negative sign eg. -
    LocaleNegativeSign = GetLocaleString(LOCALE_SNEGATIVESIGN)
End Function
Public Sub SetLocaleNegativeSign(Value As String)
    Call SetLocaleString(LOCALE_SNEGATIVESIGN, Value)
End Sub
