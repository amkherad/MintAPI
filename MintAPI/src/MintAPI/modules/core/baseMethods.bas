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
    Cols() As BASEMETHODS_SAFEPATH_COLUMN
    colsCount As Long
End Type


Private Declare Function API_VarPtr Lib "msvbvm60" Alias "VarPtr" (Ptr As Any) As Long
Private Declare Function API_VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub API_baseMethods_CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Function API_baseMethods_GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As API_OSVERSIONINFO) As Long
Private Declare Function API_baseMethods_GetVersion Lib "Kernel32" Alias "GetVersion" () As Long
Private Declare Function API_baseMethods_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function API_baseMethods_GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function API_baseMethods_GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function API_baseMethods_GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function API_baseMethods_GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function API_baseMethods_SetComputerName Lib "Kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function API_baseMethods_GetUserDefaultLangID Lib "Kernel32" Alias "GetUserDefaultLangID" () As Integer
Private Declare Function API_baseMethods_GetSystemDefaultLCID Lib "Kernel32" Alias "GetSystemDefaultLCID" () As Long
Private Declare Function API_baseMethods_GetUserDefaultLCID Lib "Kernel32" Alias "GetUserDefaultLCID" () As Long
Private Declare Function API_baseMethods_SetCurrentDirectory Lib "Kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function API_baseMethods_GetCurrentDirectory Lib "Kernel32" Alias "GetCurrentDirectory" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Declare Function API_baseMethods_CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
'Private Declare Function API_baseMethods_CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function API_baseMethods_RemoveDirectory Lib "Kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function API_baseMethods_GetFullPathName Lib "Kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function API_baseMethods_GetUserProfileDirectory Lib "userenv" (ByVal hToken As Long, ByVal lpProfileDir As String, ByRef lpcchSize As Long) As Long
Private Declare Function API_baseMethods_GetSiteDirectory Lib "advapi32" Alias "GetSiteDirectoryA" (ByVal hToken As Long, ByVal pszSiteDirectory As String, ByVal uSize As Long) As Long
Private Declare Function API_baseMethods_GetSystemWindowsDirectory Lib "Kernel32" Alias "GetSystemWindowsDirectoryA" (ByVal lpBuffer As String, ByVal uSize As Long) As Long
Private Declare Function API_baseMethods_GetProfilesDirectory Lib "userenv" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
'Private Declare Function API_baseMethods_GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long


Private Declare Function API_baseMethods_SetLocaleInfo Lib "Kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Private Declare Function API_baseMethods_GetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" (ByVal lLocale As Long, ByVal lLocaleType As Long, ByVal sLCData As String, ByVal lBufferLength As Long) As Long
Private Declare Function API_baseMethods_GetSystemDefaultLangID Lib "Kernel32" Alias "GetSystemDefaultLangID" () As Integer
Private Declare Function API_baseMethods_VerLanguageName Lib "Kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long

Public Enum ConcatPathEndMode
    cpeEndsWithSlash
    cpeNoEndsWithSlash
    cpeNotMatter
End Enum



Public Enum API_EqualToFlags
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

Public Sub Initialize()
    If inited Then Exit Sub
    
    inited = True
End Sub
'Public Sub Dispose(Optional ByVal Force As Boolean = False)
'    If Not inited Then Exit Sub
'    inited = False
'End Sub

Public Function GetLPSTR(lpStr As String) As String
    Dim Index As Long
    Index = InStr(1, lpStr, Chr(0)) - 1
    If Index <= 0 Then Exit Function
    GetLPSTR = Left(lpStr, Index)
End Function
Public Function GetLPSTRWL(lpStr As String, Length As Long) As String 'With Length
    GetLPSTRWL = Left(lpStr, Length)
End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------

'Public Function osVersion() As API_OSVersion
'    On Error GoTo ErrorHandler
'    Dim mjVersion As Long, mnVersion As Long, pltfmID As Long
'    mjVersion = bmVersionInfoRecord.dwMajorVersion
'    mnVersion = bmVersionInfoRecord.dwMinorVersion
'    pltfmID = bmVersionInfoRecord.dwPlatformId
'    Select Case mjVersion
'        Case 1
'
'        Case 2
'
'        Case 3
'            If pltfmID = pltID_Win32s Then
'                OSVersion = OSV_Win3X
'            ElseIf pltfmID = pltID_WinNT Then
'                OSVersion = OSV_WinNT40
'            Else
'                OSVersion = OSV_WinNT
'            End If
'        Case 4
'            If pltfmID = pltID_WinNT Then
'                OSVersion = OSV_WinNT40
'            Else
'                OSVersion = IIf(mnVersion = 0, OSV_Win95, OSV_Win98)
'            End If
'        Case 5
'            OSVersion = IIf(mnVersion = 0, OSV_Win2000, OSV_WinXP)
'        Case 6
'            OSVersion = IIf(mnVersion = 0, OSV_WinVista, OSV_Win7)
'        Case Else
'            OSVersion = OSV_Unknown
'    End Select
'Exit Function
'ErrorHandler:
'    OSVersion = OSV_Unknown
'End Function

Public Function CurrentUser() As String
    Dim Buf As String * SMALLLPSTR, bufSize As Long
    bufSize = SMALLLPSTR
    Buf = String(SMALLLPSTR, Chr(0))
    Call API_baseMethods_GetUserName(Buf, bufSize)
    CurrentUser = GetLPSTR(Buf)
End Function
Public Function GetTempPath() As String
    Dim Buf As String * LARGELPSTR, bufSize As Long
    bufSize = LARGELPSTR
    Buf = String(LARGELPSTR, Chr(0))
    Call API_baseMethods_GetTempPath(bufSize, Buf)
    GetTempPath = GetLPSTR(Buf)
End Function
Public Function GetSystemPath() As String
    Dim Buf As String * LARGELPSTR, bufSize As Long
    bufSize = LARGELPSTR
    Buf = String(LARGELPSTR, Chr(0))
    Call API_baseMethods_GetSystemDirectory(Buf, bufSize)
    GetSystemPath = GetLPSTR(Buf)
End Function
Public Function GetWindowsPath() As String
    Dim Buf As String * LARGELPSTR, bufSize As Long
    bufSize = LARGELPSTR
    Buf = String(LARGELPSTR, Chr(0))
    Call API_baseMethods_GetWindowsDirectory(Buf, bufSize)
    GetWindowsPath = GetLPSTR(Buf)
End Function
Public Function GetUserPath() As String
    Dim Buf As String * LARGELPSTR, bufSize As Long
    bufSize = LARGELPSTR
    Buf = String(LARGELPSTR, Chr(0))
    Call API_baseMethods_GetProfilesDirectory(Buf, bufSize)
    GetUserPath = GetLPSTR(Buf)
End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'Public Function EnumDirectoryFiles(Path As String, Optional Attributes As VbFileAttribute) As String()
'    Dim CP As String, Count As Long, RetVal() As String
'    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
'    CP = Dir(ConcatPath(Path, "", cpeEndsWithSlash), vbNormal Or Attributes)
'    While CP <> ""
'        ReDim Preserve RetVal(Count)
'        RetVal(Count) = CP
'        Count = Count + 1
'        CP = Dir
'    Wend
'    EnumDirectoryFiles = RetVal
'End Function
'Public Function CountDirectoryFiles(Path As String, Optional Attributes As VbFileAttribute) As Long
'    Dim CP As String, Count As Long
'    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
'    CP = Dir(ConcatPath(Path, "", cpeEndsWithSlash), Attributes)
'    While CP <> ""
'        Count = Count + 1
'        CP = Dir
'    Wend
'    CountDirectoryFiles = Count
'End Function
'Public Function EnumDirectoryFolders(Path As String, Optional Attributes As VbFileAttribute = 0) As String()
'    Dim CP As String, Count As Long, RetVal() As String
'    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
'    CP = Dir(Path, vbDirectory Or Attributes)
'    If CP <> "" Then CP = Dir
'    If CP <> "" Then CP = Dir
'    While CP <> ""
'        On Local Error GoTo cantOpenFile
'        If (GetAttr(ConcatPath(Path, CP)) And vbDirectory) = vbDirectory Then
'            On Error GoTo 0
'            ReDim Preserve RetVal(Count)
'            RetVal(Count) = CP
'            Count = Count + 1
'        End If
'cantOpenFile:
'        CP = Dir
'    Wend
'    EnumDirectoryFolders = RetVal
'End Function
'Public Function CountDirectoryFolders(Path As String, Optional Attributes As VbFileAttribute = 0) As Long
'    Dim CP As String, Count As Long
'    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
'    CP = Dir(Path, vbDirectory Or Attributes)
'    If CP <> "" Then CP = Dir
'    If CP <> "" Then CP = Dir
'    While CP <> ""
'        On Local Error GoTo cantOpenFile
'        If (GetAttr(ConcatPath(Path, CP)) And vbDirectory) = vbDirectory Then
'            On Error GoTo 0
'            Count = Count + 1
'        End If
'cantOpenFile:
'        CP = Dir
'    Wend
'    CountDirectoryFolders = Count
'End Function
'Public Function CountSubDirectoriesFiles(Path As String, Optional Attributes As VbFileAttribute = 0) As Long
'    Dim Count As Long, subDirs() As String, subCount As Long
'    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
'    subDirs = EnumDirectoryFolders(Path, Attributes)
'    subCount = ArraySize(subDirs)
'    Dim i As Long
'    Count = CountDirectoryFiles(Path, Attributes)
'    For i = 0 To subCount - 1
'        Count = Count + CountSubDirectoriesFiles(ConcatPath(Path, subDirs(i), cpeEndsWithSlash), Attributes)
'    Next
'    CountSubDirectoriesFiles = Count
'End Function
'Public Function EnumSubDirectoriesFiles(Path As String, Optional Attributes As VbFileAttribute = 0) As String()
'    Dim subDirs() As String, subCount As Long, RetVal() As String, bufVal() As String
'    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "/"
'    subDirs = EnumDirectoryFolders(Path, Attributes)
'    subCount = ArraySize(subDirs)
'    Dim i As Long
'    RetVal = EnumDirectoryFiles(Path, Attributes)
'    For i = 0 To subCount - 1
'        bufVal = EnumSubDirectoriesFiles(ConcatPath(Path, subDirs(i), cpeEndsWithSlash), Attributes)
'        Call AppendArrayToArray(RetVal, bufVal)
'    Next
'    EnumSubDirectoriesFiles = RetVal()
'End Function
'Public Function GetFilePath(ByVal Path As String) As String
'On Error GoTo Err
'    Dim SlashIndex As Long, backslashIndex As Long
'    SlashIndex = InStrRev(Path, "/")
'    backslashIndex = InStrRev(Path, "\")
'    SlashIndex = IIf(SlashIndex >= backslashIndex, SlashIndex, backslashIndex)
'    If SlashIndex = 0 Then throw Exps.InvalidPathException(Path)
'    backslashIndex = Len(Path)
'    GetFilePath = Left(Path, backslashIndex - (backslashIndex - SlashIndex + 1))
'Exit Function
'Err:
'End Function
'Public Function GetFileName(ByVal Path As String) As String
'On Error GoTo Err
'    Dim SlashIndex As Long, backslashIndex As Long
'    SlashIndex = InStrRev(Path, "/")
'    backslashIndex = InStrRev(Path, "\")
'    SlashIndex = IIf(SlashIndex >= backslashIndex, SlashIndex, backslashIndex)
'    If SlashIndex = 0 Then throw Exps.InvalidPathException(Path)
'    backslashIndex = Len(Path)
'    GetFileName = Right(Path, backslashIndex - SlashIndex)
'Exit Function
'Err:
'End Function
'Public Function GetFileExtention(ByVal Path As String) As String
'On Error GoTo Err
'    Dim SlashIndex As Long, backslashIndex As Long
'    SlashIndex = InStrRev(Path, "/")
'    backslashIndex = InStrRev(Path, "\")
'    SlashIndex = IIf(SlashIndex >= backslashIndex, SlashIndex, backslashIndex)
'    'If slashIndex = 0 Then throw Exps.InvalidPathException
'    backslashIndex = InStrRev(Path, ".")
'    If backslashIndex = 0 Then
'        GetFileExtention = ""
'        Exit Function
'    End If
'    If SlashIndex > backslashIndex Then
'        GetFileExtention = ""
'        Exit Function
'    Else
'        SlashIndex = Len(Path)
'        GetFileExtention = Right(Path, SlashIndex - backslashIndex)
'    End If
'Exit Function
'Err:
'End Function
'Public Function GetFileNameOnly(ByVal Path As String) As String
'On Error GoTo Err
'    Dim SlashIndex As Long, backslashIndex As Long, fLen As Long
'    fLen = Len(Path)
'    SlashIndex = InStrRev(Path, "/")
'    backslashIndex = InStrRev(Path, "\")
'    SlashIndex = IIf(SlashIndex >= backslashIndex, SlashIndex, backslashIndex)
'    If SlashIndex = 0 Then throw Exps.InvalidPathException
'    backslashIndex = InStrRev(Path, ".")
'    If backslashIndex = 0 Then
'        GetFileNameOnly = Right(Path, fLen - SlashIndex)
'        Exit Function
'    End If
'    If SlashIndex > backslashIndex Then
'        GetFileNameOnly = Right(Path, fLen - SlashIndex)
'        Exit Function
'    Else
'        GetFileNameOnly = Mid(Path, SlashIndex + 1, backslashIndex - SlashIndex - 1)
'    End If
'Exit Function
'Err:
'End Function
'Public Function ConcatPath(ByVal Path As String, ByVal PathToAdd As String, Optional EndWithSlash As ConcatPathEndMode = ConcatPathEndMode.cpeNotMatter, Optional Slash As String = "/") As String
'    Dim p As String, A As String
'    A = PathToAdd
'    p = Path
'    If A <> "" Then
'        If (Left(A, 1) = "/") And (Left(A, 1) = "\") Then
'            A = Mid(A, 2)
'        End If
'    Else
'        ConcatPath = p
'        GoTo checkLastSlash
'    End If
'    If p <> "" Then
'        If Not ((Right(p, 1) = "/") And (Right(p, 1) = "\")) Then
'            p = p & Slash
'        End If
'    Else
'        ConcatPath = A
'        GoTo checkLastSlash
'    End If
'    ConcatPath = p & A
'    Call CheckPathValidation(ConcatPath, True, False)
'checkLastSlash:
'    If Not EqualTo(SomeEqual, EndWithSlash, ConcatPathEndMode.cpeNoEndsWithSlash, ConcatPathEndMode.cpeEndsWithSlash) Then Exit Function
'    If ConcatPath <> "" Then
'        If Right(ConcatPath, 1) = Slash Then
'            If EndWithSlash = cpeNoEndsWithSlash Then ConcatPath = Mid(ConcatPath, 1, Len(ConcatPath) - 1)
'        Else
'            If EndWithSlash = cpeEndsWithSlash Then ConcatPath = ConcatPath & Slash
'        End If
'    End If
'End Function
'Public Function CheckPathValidation(Path As String, MakeTrueForm As Boolean, Optional CheckForDrive As Boolean = True) As Boolean
'    If Len(Trim(Path)) = 0 Then Exit Function
'    Dim charsIndex As Long
'    charsIndex = InStr(1, Path, "*") + InStr(1, Path, "?") + InStr(1, Path, """") + InStr(1, Path, "<") + InStr(1, Path, ">")
'    If charsIndex > 0 Then
'        CheckPathValidation = False
'        Exit Function
'    End If
'    If CheckForDrive Then
'        If (Left(Path, 2) Like "?:") Then
'            If Len(Path) > 2 Then
'                If Not (Left(Path, 3) Like "?:[/,\,|]") Then
'                    CheckPathValidation = False
'                    Exit Function
'                End If
'            End If
'        Else
'            CheckPathValidation = False
'            Exit Function
'        End If
'    End If
'    If MakeTrueForm Then
'        Path = RemovePathBads(Path, CheckForDrive)
'    End If
'    CheckPathValidation = True
'End Function
'Public Function RemovePathBads(ByVal Path As String, Optional CheckForDrive As Boolean = True) As String
'    Dim DriveLetter As String * 1, IsDriveLetter As Boolean
'    Path = Trim(Path)
'    Path = Replace(Path, "\", "/")
'    Path = Replace(Path, "|", "/")
'    IsDriveLetter = False
'    If (Left(Path, 2) Like "?:") Then
'        DriveLetter = Left(Path, 1)
'        IsDriveLetter = True
'        If Len(Path) >= 3 Then
'            If (Left(Path, 3) Like "?:/") Then
'                On Error Resume Next
'                If Len(Path) = 3 Then
'                    RemovePathBads = Path
'                    Exit Function
'                End If
'                Path = Mid(Path, 4)
'            Else
'                throw Exps.InvalidPathException
'            End If
'        Else
'            RemovePathBads = Path
'            Exit Function
'        End If
'    Else
'        If CheckForDrive Then throw Exps.InvalidPathException
'    End If
'
'    Dim Cols As BASEMETHODS_SAFEPATH
'    Cols = SplitPathToSafePath(Path)
'
'    Dim i As Long, Buf As String, doubleDot As Long
'    doubleDot = 0
'    For i = Cols.colsCount - 1 To 0 Step -1
'        Buf = Trim(Cols.Cols(i).Value)
'        If Buf = "" Or Buf = "." Then
'            Cols.Cols(i).Include = False
'        ElseIf Buf = ".." Then
'            Cols.Cols(i).Include = False
'            doubleDot = doubleDot + 1
'        ElseIf doubleDot > 0 Then
'            Cols.Cols(i).Include = False
'            doubleDot = doubleDot - 1
'        Else
'            Cols.Cols(i).Include = True
'        End If
'    Next
'
'    Path = ""
'    For i = 0 To Cols.colsCount - 1
'        If Cols.Cols(i).Include Then
'            Path = Path & "/" & Cols.Cols(i).Value
'        End If
'    Next
'
'    On Error Resume Next
'
'    If Len(Path) > 1 Then _
'        Path = Mid(Path, 2) ' removes first /
'
'    If Not IsDriveLetter Then
'        RemovePathBads = Path
'    Else
'        RemovePathBads = DriveLetter & ":/" & Path  'this also add \ character between driveLetter and path.
'    End If
'End Function
'Public Function SplitPathToSafePath(ByVal Path As String) As BASEMETHODS_SAFEPATH
'    Dim Cols As BASEMETHODS_SAFEPATH
'    Dim strs() As String
'    If InStr(1, Path, "|") > 0 Then Path = Replace(Path, "|", "/")
'    If InStr(1, Path, "\") > 0 Then Path = Replace(Path, "\", "/")
'    strs = Split(Path, "/")
'    Dim i As Long, StrsCount As Long, cIndex As Long
'    StrsCount = ArraySize(strs) 'zero based
'    If StrsCount = 0 Then GoTo zeroLength
'    Cols.colsCount = StrsCount
'    ReDim Cols.Cols(Cols.colsCount - 1)
'    For i = 0 To StrsCount - 1
'        Cols.Cols(i).Value = strs(i)
'        Cols.Cols(i).Include = True
'    Next
'zeroLength:
'    SplitPathToSafePath = Cols
'End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------



'Public Function IndexOfByteArray(Source() As Byte, Find() As Byte, Optional StartAt As Long = 0) As Long
'    Dim i As Long, sj As Long 'Same j
'    Dim sLen As Long, sLB As Long, fLB As Long, fLen As Long
'    sLen = ArraySize(Source)
'    If sLen <= 0 Then Exit Function
'    fLen = ArraySize(Find)
'    If fLen <= 0 Then Exit Function
'    sLB = LBound(Source)
'    fLB = LBound(Find)
'
'    Dim areSame As Boolean
'
'    For i = StartAt To sLen - 1
'        areSame = True
'        For sj = 0 To fLen - 1
'            If (Source(sLB + i + sj) <> Find(fLB + sj)) Then
'                areSame = False
'                Exit For
'            End If
'        Next
'        If areSame Then
'            IndexOfByteArray = i
'            Exit Function
'        End If
'    Next
'    IndexOfByteArray = -1
'End Function
'Public Function LastIndexOfByteArray(Source() As Byte, Find() As Byte, Optional StartAt As Long = 0) As Long
'    Dim i As Long, sj As Long 'Same j
'    Dim sLen As Long, sLB As Long, fLB As Long, fLen As Long
'    sLen = ArraySize(Source)
'    If sLen <= 0 Then Exit Function
'    fLen = ArraySize(Find)
'    If fLen <= 0 Then Exit Function
'    sLB = LBound(Source)
'    fLB = LBound(Find)
'
'    Dim areSame As Boolean
'
'    For i = StartAt To sLen - 1
'        areSame = True
'        For sj = 0 To fLen - 1
'            If (Source(sLB + i + sj) <> Find(fLB + sj)) Then
'                areSame = False
'                Exit For
'            End If
'        Next
'        If areSame Then
'            LastIndexOfByteArray = i
'            Exit Function
'        End If
'    Next
'    LastIndexOfByteArray = -1
'End Function


'Public Sub ByteArrayToArgsArray(B() As Byte, Args() As Variant)
'
'End Sub
'Public Function ArgsArrayToByteArray(Args() As Variant) As Byte()
'
'End Function
'Public Function ParamArrayArgsToByteArray(Target) As Byte()
'
'End Function
'Public Sub ByteArrayToParamArrayArgs(targetByteArray() As Byte, Target)
'
'End Sub
'
'Public Function ArrayToByteArray(TargetArray) As Byte()
'
'End Function

Public Function GetSubByteArray(targetByteArray() As Byte, ByVal StartIndex As Long, Optional ByVal Length As Long = -1) As Byte()
    If Length = 0 Then Exit Function
    If Length < 0 Then throw Exps.InvalidArgumentException("Length")
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
Public Function GetByteArraySomeLength(TargetArray() As Byte, Length As Long) As Byte()
    If Length = 0 Then Exit Function
    If Length < 0 Then throw Exps.InvalidArgumentException("Length")
    Dim RetVal() As Byte
    Dim outLen As Long
    outLen = ArraySize(TargetArray)
    If Length > 0 Then
        If Length > outLen Then
            throw Exps.IndexOutOfRangeException("Length is greater than actual targetArray size.")
        Else
            outLen = Length
        End If
    End If
    If outLen <= 0 Then Exit Function
    ReDim RetVal(outLen - 1)
    Dim i As Long, l_Bound As Long
    l_Bound = LBound(TargetArray)
    For i = 0 To outLen - 1
        RetVal(i) = TargetArray(i + l_Bound)
    Next
'    Dim refDestination As Long, refSource As Long
'    Call API_CopyMemory(ByVal VarPtr(refDestination), ByVal API_VarPtrArray(retVal), 4)
'    Call API_CopyMemory(ByVal VarPtr(refSource), ByVal API_VarPtrArray(targetArray), 4)
'    Call API_CopyMemory(ByVal refDestination, ByVal refSource, outLen)
    GetByteArraySomeLength = RetVal
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
Public Function IsByteArrayLikeAnother(fA() As Byte, SA() As Byte) As Boolean
    
End Function

Public Function ByteArrayToUpper(bArr() As Byte) As Byte()
    Dim RetVal() As Byte
    RetVal = bArr
    Dim i As Long
    For i = 0 To ArraySize(RetVal) - 1
        If RetVal(i) >= 97 And RetVal(i) <= 122 Then
            RetVal(i) = RetVal(i) - 32
        End If
    Next
    ByteArrayToUpper = RetVal
End Function
Public Function ByteArrayToLower(bArr() As Byte) As Byte()
    Dim RetVal() As Byte
    RetVal = bArr
    Dim i As Long
    For i = 0 To ArraySize(RetVal) - 1
        If RetVal(i) >= 65 And RetVal(i) <= 90 Then
            RetVal(i) = RetVal(i) + 32
        End If
    Next
    ByteArrayToLower = RetVal
End Function

Public Function FillArray(TargetArray, whattofill, OutputSize As Long, Optional FillToLeft = True)
    If (VarType(TargetArray) And vbArray) <> vbArray Then throw Exps.ArrayExpectedException
    Dim targetLength As Long
    targetLength = ArraySize(TargetArray)

End Function
Public Sub InsertArrayIndex(TargetArray, Index, Item)
    Dim ArrayType As VbVarType, ItemType As VbVarType
    ItemType = VarType(Item)
    ArrayType = VarType(TargetArray)
    If (ArrayType And vbArray) <> vbArray Then throw Exps.ArrayExpectedException("TargetArray")
    If (ArrayType And vbVariant) <> vbVariant Then
        If ArrayType <> (ItemType Or vbArray) Then _
            throw Exps.ArgumentTypeMismatch("TargetArray and Item")
    End If
    
    Dim varCount As Long, i As Long
    varCount = ArraySize(TargetArray)
    
    If varCount = 0 Then Exit Sub
    
    If Index > varCount Then throw Exps.IndexOutOfRangeException
    
    ReDim Preserve TargetArray(varCount)
    If (ArrayType And VBObject) = VBObject Then
        For i = varCount To Index + 1 Step -1
            Set TargetArray(i) = TargetArray(i - 1)
        Next
        Set TargetArray(Index) = Item
    ElseIf (ArrayType And vbVariant) = vbVariant Then
        If (ItemType And VBObject) = VBObject Then
            For i = varCount To Index + 1 Step -1
                Set TargetArray(i) = TargetArray(i - 1)
            Next
            Set TargetArray(Index) = Item
        Else
            For i = varCount To Index + 1 Step -1
                TargetArray(i) = TargetArray(i - 1)
            Next
            TargetArray(Index) = Item
        End If
    Else
        For i = varCount To Index + 1 Step -1
            TargetArray(i) = TargetArray(i - 1)
        Next
        TargetArray(Index) = Item
    End If
End Sub
Public Sub InsertArrayIndexArray(TargetArray, Index, insertArray, Optional Length As Long = -1)
    Dim ArrayType As VbVarType, insertType As VbVarType
    insertType = VarType(insertArray)
    ArrayType = VarType(TargetArray)
    If (ArrayType And vbArray) <> vbArray Then throw Exps.ArrayExpectedException("TargetArray")
    If (ArrayType And vbVariant) <> vbVariant Then
        If ArrayType <> insertType Then _
            throw Exps.ArgumentException("Both arguments type must be as one.")
    End If
    
    Dim varCount As Long, i As Long, sArrCtr As Long
    Dim howMuchToShift As Long, outLength As Long
    varCount = ArraySize(TargetArray)
    howMuchToShift = ArraySize(insertArray)
    
    If howMuchToShift = 0 Then Exit Sub
    If varCount = 0 Then Exit Sub
    
    If Index > varCount Then throw Exps.IndexOutOfRangeException
    
    If Length <> -1 Then howMuchToShift = Length
    If Length < -1 Then throw Exps.IndexOutOfRangeException
    If Length > howMuchToShift Then throw Exps.IndexOutOfRangeException
    
    outLength = varCount + howMuchToShift
    outLength = outLength - 1
    
    ReDim Preserve TargetArray(outLength)
    If (ArrayType And VBObject) = VBObject Then
        For i = outLength To Index + howMuchToShift Step -1
            Set TargetArray(i) = TargetArray(i - howMuchToShift)
        Next
        For i = 0 To howMuchToShift - 1
            Set TargetArray(Index + i) = insertArray(i)
        Next
    ElseIf (ArrayType And vbVariant) = vbVariant Then
        If (insertType And vbVariant) = vbVariant Then
            For i = outLength To Index + howMuchToShift Step -1
                TargetArray(i) = TargetArray(insertType)
            Next
            For i = 0 To howMuchToShift - 1
                insertType = i - howMuchToShift
                ArrayType = VarType(insertArray(i))
                If ArrayType = VBObject Then
                    Set TargetArray(Index + i) = insertArray(i)
                Else
                        TargetArray(Index + i) = insertArray(i)
                End If
            Next
        Else
            For i = outLength To Index + howMuchToShift Step -1
                TargetArray(i) = TargetArray(i - howMuchToShift)
            Next
            For i = 0 To howMuchToShift - 1
                TargetArray(Index + i) = insertArray(i)
            Next
        End If
    Else
        For i = outLength To Index + howMuchToShift Step -1
            TargetArray(i) = TargetArray(i - howMuchToShift)
        Next
        For i = 0 To howMuchToShift - 1
            TargetArray(Index + i) = insertArray(i)
        Next
    End If
End Sub
Public Sub AppendToArray(sourceArray, Item)
    Dim sourceUBound As Long
    If (VarType(sourceArray) And vbArray) <> vbArray Then throw Exps.InvalidArgumentException
    If Not IsEmptyArray(sourceArray) Then
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
Public Sub AppendArrayToArray(sourceArray, TargetArray)
    Dim sourceArrayVT As Long
    Dim targetArrayVT As Long
    sourceArrayVT = VarType(sourceArray)
    targetArrayVT = VarType(TargetArray)
    If (sourceArrayVT And vbArray) <> vbArray Then throw Exps.InvalidArgumentException
    If (targetArrayVT And vbArray) <> vbArray Then
        Call AppendToArray(sourceArray, TargetArray)
        Exit Sub
    End If
    If targetArrayVT <> sourceArrayVT Then throw Exps.InvalidArgumentException("Both arrays must be as the same type.")
    Dim sourceUBound As Long, sourceLBound As Long
    Dim targetUBound As Long, targetLBound As Long
    If IsEmptyArray(TargetArray) Then
        Exit Sub
    Else
        If IsEmptyArray(sourceArray) Then
            sourceArray = TargetArray
        End If
    End If
    sourceLBound = LBound(sourceArray)
    sourceUBound = UBound(sourceArray)
    targetLBound = LBound(TargetArray)
    targetUBound = UBound(TargetArray)
    Dim i As Long
    For i = targetLBound To targetUBound
        sourceUBound = sourceUBound + 1
        ReDim Preserve sourceArray(sourceUBound)
        If VarType(TargetArray(i)) = VBObject Then
            Set sourceArray(sourceUBound) = TargetArray(i)
        Else
                sourceArray(sourceUBound) = TargetArray(i)
        End If
    Next
End Sub

Public Function ExpandVArrayToString(TargetArray, Optional throwExceptionOnUnknownValues As Boolean = True) As String
    Dim targetArrayType As VbVarType, i As Long, RetVal As String
    targetArrayType = VarType(TargetArray)
    For i = LBound(TargetArray) To UBound(TargetArray)
        If (targetArrayType And vbArray) = vbArray Then
            RetVal = RetVal & ExpandVArrayToString(TargetArray(i), throwExceptionOnUnknownValues)
        ElseIf (targetArrayType And VBObject) = VBObject Then
            If throwExceptionOnUnknownValues Then throw Exps.InvalidArgumentException("One Or More Array Items Type Is Object.")
        Else
            RetVal = RetVal & CStr(TargetArray(i))
        End If
    Next
    ExpandVArrayToString = RetVal
End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
Public Function BinaryCompare(A1, A2, Optional LengthToCompare As Long = -1) As CompareResults
    Dim a1Type As VbVarType
    a1Type = VarType(A1)
    If a1Type <> VarType(A2) Then: Exit Function
    If a1Type = vbArray Then
        BinaryCompare = ArrayCompare(A1, A2, LengthToCompare)
    ElseIf a1Type = VBObject Then
        BinaryCompare = (A1 Is A2)
    Else
        BinaryCompare = (A1 = A2)
    End If
End Function


'Public Function GetLocaleString(ByVal lLocaleNum As Long) As String
'    'Generic routine to get the locale string from the Operating system.
'    Dim lBuffSize As String
'    Dim sBuffer As String
'    Dim lRet As Long
'
'    lBuffSize = 256
'    sBuffer = String(lBuffSize, vbNullChar)
'
'    'Get the information from the registry
'    lRet = API_baseMethods_GetLocaleInfo(LOCALE_USER_DEFAULT, lLocaleNum, sBuffer, lBuffSize)
'    'If lRet > 0 then success - lret is the size of the string returned
'    If lRet > 0 Then
'        GetLocaleString = Left$(sBuffer, lRet - 1)
'    End If
'End Function
'Public Sub SetLocaleString(ByVal lLocaleNum As Long, strValue As String)
'    Call API_baseMethods_SetLocaleInfo(LOCALE_USER_DEFAULT, lLocaleNum, strValue)
'End Sub
'
'Public Function LocaleDateFormat() As String
'    ' This function will return the Locale date format for the system. Note that the
'    ' returned Year is always formatted to 'YYYY' regardless, to ensure Y2k compliance.
'    Dim sDateFormat As String
'    On Error GoTo vbErrorHandler
'    sDateFormat = GetLocaleString(LOCALE_SSHORTDATE)
'
'    ' Make sure we always have YYYY format for y2k
'    If InStr(1, sDateFormat, "YYYY", vbTextCompare) = 0 Then
'        Replace sDateFormat, "YY", "YYYY"
'    End If
'    LocaleDateFormat = sDateFormat
'Exit Function
'vbErrorHandler:
'    throw Exps.Exception(Err.Description)
'    'err.Raise err.Number, "LocaleSettings GetDateFormat", err.Description
'End Function
'Public Sub SetLocaleDateFormat(Value As String)
'    Call SetLocaleString(LOCALE_SSHORTDATE, Value)
'End Sub
'Public Function LocaleTimeFormat() As String
'    'This function returns the locale's defined Time Format.
'    LocaleTimeFormat = GetLocaleString(LOCALE_STIMEFORMAT)
'Exit Function
'vbErrorHandler:
'    throw Exps.Exception(Err.Description)
'    'err.Raise err.Number, "LocaleSettings GetTimeFormat", err.Description
'End Function
'Public Sub SetLocaleTimeFormat(Value As String)
'    Call SetLocaleString(LOCALE_STIMEFORMAT, Value)
'End Sub
'Public Function LocaleNumberFormat() As String
'' This function returns the Locales defined Decimal Number format
'    Dim lBuffLen As Long
'    Dim sBuffer As String
'    Dim sDecimal As String
'    Dim sThousand As String
'    Dim LRESULT As Long
'    Dim sNumFormat As String
'
'    On Error GoTo vbErrorHandler
'
'    'Setup a buffer to receive the settings
'    lBuffLen = 128
'    sBuffer = String(lBuffLen, vbNullChar)
'
'    LRESULT = API_baseMethods_GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sBuffer, lBuffLen)
'    If LRESULT <= 0 Then Exit Function
'
'    sDecimal = Left$(sBuffer, LRESULT - 1)
'
'    sBuffer = String(lBuffLen, vbNullChar)
'    LRESULT = API_baseMethods_GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, sBuffer, lBuffLen)
'    If LRESULT <= 0 Then Exit Function
'
'    sThousand = Left$(sBuffer, LRESULT - 1)
'
'    LocaleNumberFormat = ("###" & sThousand & "###" & sDecimal & "######")
'Exit Function
'
'vbErrorHandler:
'    throw Exps.Exception(Err.Description)
'    'err.Raise err.Number, "LocaleSettings GetNumberFormat", err.Description
'End Function
'Public Sub SetLocaleNumberFormat(Value As String)
'    'Call SetLocaleString(LOCALE_SSHORTDATE, Value)
'End Sub
'Public Function LocaleThousandSpecifier() As String
'    'This function returns the correct Thousand Specifier for the system Locale
'    LocaleThousandSpecifier = GetLocaleString(LOCALE_STHOUSAND)
'End Function
'Public Sub SetLocaleThousandSpecifier(Value As String)
'    Call SetLocaleString(LOCALE_STHOUSAND, Value)
'End Sub
'Public Function LocaleDecimalSpecifier() As String
'    'This function returns the correct Decimal Specifier for the system Locale
'    LocaleDecimalSpecifier = GetLocaleString(LOCALE_SDECIMAL)
'End Function
'Public Sub SetLocaleDecimalSpecifier(Value As String)
'    Call SetLocaleString(LOCALE_SDECIMAL, Value)
'End Sub
'Public Function LocaleCurrencySpecifier() As String
'    'This function returns the correct Currency Specifier for the system Locale
'    LocaleCurrencySpecifier = GetLocaleString(LOCALE_SCURRENCY)
'End Function
'Public Sub SetLocaleCurrencySpecifier(Value As String)
'    Call SetLocaleString(LOCALE_SCURRENCY, Value)
'End Sub
'Public Function LocaleSystemLanguageID() As Long
'    'Returns the System Language ID for the machine
'    LocaleSystemLanguageID = API_baseMethods_GetSystemDefaultLangID
'End Function
'Public Function LocaleSystemLanguageName() As String
'    'Returns the System Language Name eg : English (United Kingdom)
'    Dim lLangID As Long
'    Dim sBuffer As String
'    Dim lBuffSize As Long
'    Dim lRet As Long
'
'    On Error GoTo vbErrorHandler
'
'    lLangID = API_baseMethods_GetSystemDefaultLangID
'    'Setup a buffer to receive the settings
'    lBuffSize = 50
'    sBuffer = String(lBuffSize, vbNullChar)
'    lRet = API_baseMethods_VerLanguageName(lLangID, sBuffer, lBuffSize)
'    If lRet > 0 Then
'        LocaleSystemLanguageName = Left$(sBuffer, lRet)
'    End If
'Exit Function
'vbErrorHandler:
'    throw Exps.Exception(Err.Description)
'    'err.Raise err.Number, "LocaleSettings GetSysLanguageName", err.Description
'End Function
'Public Function LocaleShortMonthName(ByVal iMonthNum As Integer) As String
'    'Returns the short-month-name for the specified Month Number
'    'eg 1=Jan, 2=Feb (on English machines)
'    LocaleShortMonthName = GetLocaleString(LOCALE_SABBREVMONTHNAME1 - 1 + iMonthNum)
'End Function
'Public Sub SetLocaleShortMonthName(ByVal iMonthNum As Integer, Value As String)
'    Call SetLocaleString(LOCALE_SABBREVMONTHNAME1 - 1 + iMonthNum, Value)
'End Sub
'Public Function LocaleMonthName(ByVal iMonthNum As Integer) As String
'    'Returns the Full-Month-Name for the specified month number
'    'eg. 1=January, 2=February (on english machines)
'    LocaleMonthName = GetLocaleString(LOCALE_SMONTHNAME1 + iMonthNum - 1)
'End Function
'Public Sub SetLocaleMonthName(ByVal iMonthNum As Integer, Value As String)
'    Call SetLocaleString(LOCALE_SMONTHNAME1 - 1 + iMonthNum, Value)
'End Sub
'Public Function LocaleShortDayName(ByVal iDayNum As Integer) As String
'    'Returns the Short-Day-Name for the specified Day Number
'    'eg. 1=Mon, 2=Tue (on english machines)
'    LocaleShortDayName = GetLocaleString(LOCALE_SABBREVDAYNAME1 + iDayNum - 1)
'End Function
'Public Sub SetLocaleShortDayName(ByVal iDayNum As Integer, Value As String)
'    Call SetLocaleString(LOCALE_SABBREVDAYNAME1 - 1 + iDayNum, Value)
'End Sub
'Public Function LocaleDayName(ByVal iDayNum As Integer) As String
'    'Returns the Full Day Name for the specified Day number
'    'eg. 1=Monday, 2=Tuesday (on english machines)
'    LocaleDayName = GetLocaleString(LOCALE_SDAYNAME1 + iDayNum - 1)
'End Function
'Public Sub SetLocaleDayName(ByVal iDayNum As Integer, Value As String)
'    Call SetLocaleString(LOCALE_SDAYNAME1 - 1 + iDayNum, Value)
'End Sub
'Public Function LocaleCountry() As String
'    'Returns the Country Name eg. 'United Kingdom'
'    LocaleCountry = GetLocaleString(LOCALE_SENGCOUNTRY)
'End Function
'Public Sub SetLocaleCountry(Value As String)
'    Call SetLocaleString(LOCALE_SENGCOUNTRY, Value)
'End Sub
'Public Function LocaleLanguageName() As String
'    'Returns the Native Language Name eg. 'English'
'    LocaleLanguageName = GetLocaleString(LOCALE_SNATIVELANGNAME)
'End Function
'Public Sub SetLocaleLanguageName(Value As String)
'    Call SetLocaleString(LOCALE_SNATIVELANGNAME, Value)
'End Sub
'Public Function LocaleNativeCountryName() As String
'    LocaleNativeCountryName = GetLocaleString(LOCALE_SNATIVECTRYNAME)
'End Function
'Public Sub SetLocaleNativeCountryName(Value As String)
'    Call SetLocaleString(LOCALE_SNATIVECTRYNAME, Value)
'End Sub
'Public Function LocalePositiveSign() As String
'    'Returns the symbol used for the positive sign eg. +
'    LocalePositiveSign = GetLocaleString(LOCALE_SPOSITIVESIGN)
'End Function
'Public Sub SetLocalePositiveSign(Value As String)
'    Call SetLocaleString(LOCALE_SPOSITIVESIGN, Value)
'End Sub
'Public Function LocaleNegativeSign() As String
'' Returns the symbol used for the negative sign eg. -
'    LocaleNegativeSign = GetLocaleString(LOCALE_SNEGATIVESIGN)
'End Function
'Public Sub SetLocaleNegativeSign(Value As String)
'    Call SetLocaleString(LOCALE_SNEGATIVESIGN, Value)
'End Sub
