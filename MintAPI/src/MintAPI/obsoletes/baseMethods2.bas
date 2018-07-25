Attribute VB_Name = "baseMethods2"
''@PROJECT_LICENSE
'
'Option Explicit
'Option Base 0
'Const CLASSID As String = "baseMethods2"
'
'Public Type FILETIME
'    lLowDateTime    As Long
'    lHighDateTime   As Long
'End Type
'Public Type ChooseColor
'    lStructSize As Long
'    hWndOwner As Long
'    hInstance As Long
'    rgbResult As Long
'    lpCustColors As Long
'    Flags As Long
'    lCustData As Long
'    lpfnHook As Long
'    lpTemplateName As String
'End Type
'Public Type ChooseFont
'    lStructSize As Long
'    hWndOwner As Long          '  caller's window handle
'    hdc As Long                '  printer DC/IC or NULL
'    lpLogFont As Long
'    iPointSize As Long         '  10 * size in points of selected font
'    Flags As Long              '  enum. type flags
'    rgbColors As Long          '  returned text color
'    lCustData As Long          '  data passed to hook fn.
'    lpfnHook As Long           '  ptr. to hook function
'    lpTemplateName As String     '  custom template name
'    hInstance As Long          '  instance handle of.EXE that
'                                   '    contains cust. dlg. template
'    lpszStyle As String          '  return the style field here
'                                   '  must be LF_FACESIZE or bigger
'    nFontType As Integer          '  same value reported to the EnumFonts
'                                   '    call back with the extra FONTTYPE_
'                                   '    bits added
'    MISSING_ALIGNMENT As Integer
'    nSizeMin As Long           '  minimum pt size allowed &
'    nSizeMax As Long           '  max pt size allowed if
'                                   '    CF_LIMITSIZE is used
'End Type
'
'
'Private Declare Function API_baseMethods2_GetAllUsersProfileDirectory Lib "userenv" Alias "GetAllUsersProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
'Private Declare Function API_baseMethods2_GetDefaultUserProfileDirectory Lib "userenv" Alias "GetDefaultUserProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
'Private Declare Function API_baseMethods2_GetProfilesDirectory Lib "userenv" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
'Private Declare Function API_baseMethods2_GetUserProfileDirectory Lib "userenv" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
'Private Declare Function API_baseMethods2_GetCurrentProcess Lib "Kernel32" Alias "GetCurrentProcess" () As Long
'Private Declare Function API_baseMethods2_OpenProcessToken Lib "advapi32" Alias "OpenProcessToken" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
'Private Declare Function API_baseMethods2_GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Private Declare Function API_baseMethods2_GetVersion Lib "Kernel32" Alias "GetVersion" () As Long
'
'Private Declare Function API_baseMethods2_ChooseColor Lib "comdlg32" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
'Private Declare Function API_baseMethods2_ChooseFont Lib "comdlg32" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
'
'Private Declare Function API_baseMethods2_GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As Any, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function API_baseMethods2_GetPrivateProfileInt Lib "Kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
'Private Declare Function API_baseMethods2_WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lplFileName As String) As Long
'
'
'
'Public Const API_OFN_ALLOWMULTISELECT = &H200
'Public Const API_OFN_CREATEPROMPT = &H2000
'Public Const API_OFN_ENABLEHOOK = &H20
'Public Const API_OFN_ENABLETEMPLATE = &H40
'Public Const API_OFN_ENABLETEMPLATEHANDLE = &H80
'Public Const API_OFN_ENABLESIZING = &H800000
'Public Const API_OFN_EXPLORER = &H80000
'Public Const API_OFN_FILEMUSTEXIST = &H1000
'Public Const API_OFN_EXTENSIONDIFFERENT = &H400
'Public Const API_OFN_HIDEREADONLY = &H4
'Public Const API_OFN_LONGNAMES = &H200000
'Public Const API_OFN_NOCHANGEDIR = &H8
'Public Const API_OFN_NODEREFERENCELINKS = &H100000
'Public Const API_OFN_NOLONGNAMES = &H40000
'Public Const API_OFN_NOREADONLYRETURN = &H8000&
'Public Const API_OFN_NONETWORKBUTTON = &H20000
'Public Const API_OFN_NOTESTFILECREATE = &H10000
'Public Const API_OFN_NOVALIDATE = &H100
'Public Const API_OFN_OVERWRITEPROMPT = &H2
'Public Const API_OFN_PATHMUSTEXIST = &H800
'Public Const API_OFN_READONLY = &H1
'Public Const API_OFN_SHAREAWARE = &H4000
'Public Const API_OFN_SHAREFALLTHROUGH = 2
'Public Const API_OFN_SHARENOWARN = 1
'Public Const API_OFN_SHAREWARN = 0
'Public Const API_OFN_SHOWHELP = &H10
'
'Public Const REGPATH_SHELLFOLDERS_SYSTEM As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
'Public Const REGPATH_SHELLFOLDERS_LOCAL As String = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
'
''Dim rtn As Long, lBuffer As Long, sBuffer As String
''Dim lBufferSize As Long
''Dim lDataSize As Long
''Dim ByteArray() As Byte
'
''This constant determins wether or not to display error messages to the
''user. I have set the default value to False as an error message can and
''does become irritating after a while. Turn this value to true if you want
''to debug your programming code when reading and writing to your system
''registry, as any errors will be displayed in a message box.
'
'
'Dim inited As Boolean
'
'Public Sub Initialize()
'    If inited Then Exit Sub
'    inited = True
'End Sub
''Public Sub Dispose(Optional ByVal Force As Boolean = False)
''    If Not inited Then Exit Sub
''    inited = False
''End Sub
'
'
''Public Function FileExists(Path As String) As Boolean
''    FileExists = (dir(Path, vbNormal) <> "")
''End Function
''Public Function Exists(Path As String) As Boolean
''    Exists = (Dir(Path) <> "")
''End Function
''Public Sub MakeTreeDirectories(ByVal Path As String, Optional Created As Boolean)
''    If Not CheckPathValidation(Path, True, True) Then throw Exps.InvalidPathException
''    Dim p As String, cPath As String
''    p = Mid(Path, 3) 'remove drive ex: E:\
''    cPath = Left(Path, 3)
''    Dim steps() As String
''    If Left(p, 1) = "\" Then p = Mid(p, 2)
''    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
''    steps = Split(Trim(p), "\") '/ also checked ^
''    Dim leng As Long
''    leng = ArraySize(steps)
''    Dim i As Long
''    Created = False
''    For i = 0 To leng - 1
''        If (steps(i) = "") Or (Not CheckPathValidation(steps(i), False, False)) Then throw Exps.InvalidPathException: Created = True
''    Next
''    For i = 0 To leng - 1
''        cPath = ConcatPath(cPath, steps(i))
''        If Not Directory.Exists(cPath) Then Call MkDir(cPath)
''    Next
''End Sub
'
''---------------------------------------------
''(ByVal lpApplicationName As String, lpKeyName As String, ByVal lpDefault As String, lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
''Public Function ReadPrivateProfileString(ByVal Path As String, ByVal Section As String, ByVal KeyName As String, ByVal Default As String) As String
''    Dim outStr As String
''    outStr = String(SMALLLPSTR, Chr(0))
''    Call rLastError
''    If API_baseMethods2_GetPrivateProfileString(Section, KeyName, Default, outStr, SMALLLPSTR, Path) = 0 Then _
''        throw Exps.SystemCallFailureException
''    outStr = GetLPSTR(outStr)
''    If outStr = "" Then
''        ReadPrivateProfileString = Default
''    Else
''        ReadPrivateProfileString = outStr
''    End If
''End Function
''Public Function ReadPrivateProfileStringEX(Path As String, Section As String, KeyName As String, Default As String, BufferLength As Long) As String
''    Dim outStr As String
''    outStr = String(BufferLength, Chr(0))
''    Call rLastError
''    If API_baseMethods2_GetPrivateProfileString(Section, KeyName, Default, outStr, BufferLength, Path) = 0 Then _
''        throw Exps.SystemCallFailureException
''    outStr = GetLPSTR(outStr)
''    If outStr = "" Then
''        ReadPrivateProfileStringEX = Default
''    Else
''        ReadPrivateProfileStringEX = outStr
''    End If
''End Function
''Public Function ReadPrivateProfileInt(ByVal Path As String, ByVal Section As String, ByVal KeyName As String, ByVal Default As Long) As Long
''    Dim RetVal As Long
''    Call rLastError
''    RetVal = API_baseMethods2_GetPrivateProfileInt(Section, KeyName, Default, Path)
''    throw Exps.IfError
''    ReadPrivateProfileInt = RetVal
''End Function
''Public Function WritePrivateProfileFile(Path As String, Section As String, ByVal KeyName As String, KeyValue As String) As Boolean
''    Dim ret As Long
''    ret = API_baseMethods2_WritePrivateProfileString(Section, KeyName, KeyValue, Path)
''    If ret = 0 Then
''        WritePrivateProfileFile = True
''    Else
''        WritePrivateProfileFile = False
''    End If
''End Function
