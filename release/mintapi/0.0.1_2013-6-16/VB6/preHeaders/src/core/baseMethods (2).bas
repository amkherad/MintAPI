Attribute VB_Name = "baseMethods2"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "baseMethods2"

Public Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type
Public Type ChooseColor
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Type ChooseFont
    lStructSize As Long
    hWndOwner As Long          '  caller's window handle
    hdc As Long                '  printer DC/IC or NULL
    lpLogFont As Long
    iPointSize As Long         '  10 * size in points of selected font
    Flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As String     '  custom template name
    hInstance As Long          '  instance handle of.EXE that
                                   '    contains cust. dlg. template
    lpszStyle As String          '  return the style field here
                                   '  must be LF_FACESIZE or bigger
    nFontType As Integer          '  same value reported to the EnumFonts
                                   '    call back with the extra FONTTYPE_
                                   '    bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
                                   '    CF_LIMITSIZE is used
End Type


Private Declare Function API_baseMethods2_RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function API_baseMethods2_RegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Private Declare Function API_baseMethods2_RegregCreateKey Lib "advapi32" Alias "RegregCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function API_baseMethods2_RegregDeleteKey Lib "advapi32" Alias "RegregDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function API_baseMethods2_RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function API_baseMethods2_RegQueryValueExA Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function API_baseMethods2_RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function API_baseMethods2_RegSetValueExA Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function API_baseMethods2_RegSetValueExB Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Private Declare Function API_baseMethods2_GetAllUsersProfileDirectory Lib "userenv" Alias "GetAllUsersProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function API_baseMethods2_GetDefaultUserProfileDirectory Lib "userenv" Alias "GetDefaultUserProfileDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function API_baseMethods2_GetProfilesDirectory Lib "userenv" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function API_baseMethods2_GetUserProfileDirectory Lib "userenv" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private Declare Function API_baseMethods2_GetCurrentProcess Lib "kernel32" Alias "GetCurrentProcess" () As Long
Private Declare Function API_baseMethods2_OpenProcessToken Lib "advapi32" Alias "OpenProcessToken" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function API_baseMethods2_GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function API_baseMethods2_GetVersion Lib "kernel32" Alias "GetVersion" () As Long

Private Declare Function API_baseMethods2_ChooseColor Lib "comdlg32" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function API_baseMethods2_ChooseFont Lib "comdlg32" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long

Private Declare Function API_baseMethods2_GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As String, ByVal lpDefault As String, lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function API_baseMethods2_WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As String, ByVal lpString As Any, ByVal lplFileName As String) As Long

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1009&
Private Const ERROR_BADKEY = 1010&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_INVALID_PARAMETER = 87&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234&

Private Const REG_NONE = 0&
Private Const REG_SZ = 1&
Private Const REG_EXPAND_SZ = 2&
Private Const REG_BINARY = 3&
Private Const REG_DWORD = 4&
Private Const REG_DWORD_LITTLE_ENDIAN = 4&
Private Const REG_DWORD_BIG_ENDIAN = 5&
Private Const REG_LINK = 6&
Private Const REG_MULTI_SZ = 7&
Private Const REG_RESOURCE_LIST = 8&
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_CREATE_LINK = &H20&
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Private Const KEY_EXECUTE = KEY_READ


Public Const API_OFN_ALLOWMULTISELECT = &H200
Public Const API_OFN_CREATEPROMPT = &H2000
Public Const API_OFN_ENABLEHOOK = &H20
Public Const API_OFN_ENABLETEMPLATE = &H40
Public Const API_OFN_ENABLETEMPLATEHANDLE = &H80
Public Const API_OFN_ENABLESIZING = &H800000
Public Const API_OFN_EXPLORER = &H80000
Public Const API_OFN_FILEMUSTEXIST = &H1000
Public Const API_OFN_EXTENSIONDIFFERENT = &H400
Public Const API_OFN_HIDEREADONLY = &H4
Public Const API_OFN_LONGNAMES = &H200000
Public Const API_OFN_NOCHANGEDIR = &H8
Public Const API_OFN_NODEREFERENCELINKS = &H100000
Public Const API_OFN_NOLONGNAMES = &H40000
Public Const API_OFN_NOREADONLYRETURN = &H8000&
Public Const API_OFN_NONETWORKBUTTON = &H20000
Public Const API_OFN_NOTESTFILECREATE = &H10000
Public Const API_OFN_NOVALIDATE = &H100
Public Const API_OFN_OVERWRITEPROMPT = &H2
Public Const API_OFN_PATHMUSTEXIST = &H800
Public Const API_OFN_READONLY = &H1
Public Const API_OFN_SHAREAWARE = &H4000
Public Const API_OFN_SHAREFALLTHROUGH = 2
Public Const API_OFN_SHARENOWARN = 1
Public Const API_OFN_SHAREWARN = 0
Public Const API_OFN_SHOWHELP = &H10

Public Const REGPATH_SHELLFOLDERS_SYSTEM As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
Public Const REGPATH_SHELLFOLDERS_LOCAL As String = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"

Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte

'This constant determins wether or not to display error messages to the
'user. I have set the default value to False as an error message can and
'does become irritating after a while. Turn this value to true if you want
'to debug your programming code when reading and writing to your system
'registry, as any errors will be displayed in a message box.


Dim inited As Boolean

Public Sub Initialize()
    If inited Then Exit Sub
    Call baseConstants.Initialize
    Call baseExceptions.Initialize
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    inited = False
End Sub


Public Function regSetDWORDValue(subKey As String, Entry As String, Value As Long)

Call ParseKey(subKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, subKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = API_baseMethods2_RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
        throw Exception(regErrorMsg(rtn))        'display the error
      End If
      rtn = API_baseMethods2_RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
        throw Exception(regErrorMsg(rtn))  'display the error
   End If
End If

End Function
Public Function regGetDWORDValue(subKey As String, Entry As String)

Call ParseKey(subKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, subKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = API_baseMethods2_RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = API_baseMethods2_RegCloseKey(hKey)  'close the key
         regGetDWORDValue = lBuffer  'return the value
      Else                        'otherwise, if the value couldnt be retreived
         regGetDWORDValue = ""  'return Error to the user
         throw Exception(regErrorMsg(rtn))        'tell the user what was wrong
      End If
   Else 'otherwise, if the key couldnt be opened
      regGetDWORDValue = ""        'return Error to the user
      throw Exception(regErrorMsg(rtn))        'tell the user what was wrong
   End If
End If

End Function
Public Function regSetBinaryValue(subKey As String, Entry As String, Value As String)

Call ParseKey(subKey, MainKeyHandle)
Dim i As Long
If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, subKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      lDataSize = Len(Value)
      ReDim ByteArray(lDataSize)
      For i = 1 To lDataSize
      ByteArray(i) = Asc(mID$(Value, i, 1))
      Next
      rtn = API_baseMethods2_RegSetValueExB(hKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if the was an error writting the value
         throw Exception(regErrorMsg(rtn))
      End If
      rtn = API_baseMethods2_RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      throw Exception(regErrorMsg(rtn))
   End If
End If

End Function


Public Function regGetBinaryValue(subKey As String, Entry As String)

Call ParseKey(subKey, MainKeyHandle)
Dim i As Long

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, subKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened
      lBufferSize = 1
      rtn = API_baseMethods2_RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0, lBufferSize) 'get the value from the registry
      sBuffer = Space(lBufferSize)
      rtn = API_baseMethods2_RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = API_baseMethods2_RegCloseKey(hKey)  'close the key
         regGetBinaryValue = sBuffer 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         regGetBinaryValue = "" 'return Error to the user
         throw Exception(regErrorMsg(rtn))
      End If
   Else 'otherwise, if the key couldnt be opened
      regGetBinaryValue = "" 'return Error to the user
      throw Exception(regErrorMsg(rtn))
   End If
End If

End Function
Public Function regDeleteKey(KeyName As String)

Call ParseKey(KeyName, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, KeyName, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = API_baseMethods2_RegregDeleteKey(hKey, KeyName) 'delete the key
      rtn = API_baseMethods2_RegCloseKey(hKey)  'close the key
   End If
End If

End Function

Private Function GetMainKeyHandle(MainKeyName As String) As Long

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006

Select Case MainKeyName
    Case "HKEY_CLASSES_ROOT", "HKEY_CLASSESROOT", "CLASSES_ROOT", "CLASSESROOT"
        GetMainKeyHandle = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER", "HKEY_CURRENTUSER", "CURRENT_USER", "CURRENTUSER"
        GetMainKeyHandle = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE", "HKEY_LOCALMACHINE", "LOCAL_MACHINE", "LOCALMACHINE"
        GetMainKeyHandle = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS", "USERS"
        GetMainKeyHandle = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA", "HKEY_PERFORMANCEDATA", "PERFORMANCE_DATA", "PERFORMANCEDATA"
        GetMainKeyHandle = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG", "HKEY_CURRENTCONFIG", "CURRENT_CONFIG", "CURRENTCONFIG"
        GetMainKeyHandle = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA", "HKEY_DYNDATA", "DYN_DATA", "DYNDATA"
        GetMainKeyHandle = HKEY_DYN_DATA
    Case Else
        throw InvalidArgumentValueException("Invalid Registry Hive Key.")
End Select

End Function

Public Function regErrorMsg(lErrorCode As Long) As String

'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
    Case 1009, 1015
        regErrorMsg = "The Registry Database is corrupt!"
    Case 2, 1010
        regErrorMsg = "Bad Key Name"
    Case 1011
        regErrorMsg = "Can't Open Key"
    Case 4, 1012
        regErrorMsg = "Can't Read Key"
    Case 5
        regErrorMsg = "Access to this key is denied"
    Case 1013
        regErrorMsg = "Can't Write Key"
    Case 8, 14
        regErrorMsg = "Out of memory"
    Case 87
        regErrorMsg = "Invalid Parameter"
    Case 234
        regErrorMsg = "There is more data than the buffer has been allocated to hold."
    Case Else
        regErrorMsg = "Undefined Error Code:  " & str$(lErrorCode)
End Select

End Function

Public Function regGetStringValue(subKey As String, Entry As String)
Call ParseKey(subKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, subKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      sBuffer = Space(255)     'make a buffer
      lBufferSize = Len(sBuffer)
      rtn = API_baseMethods2_RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = API_baseMethods2_RegCloseKey(hKey)  'close the key
         regGetStringValue = GetLPSTR(sBuffer)  'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         regGetStringValue = "" 'return Error to the user
         throw Exception(regErrorMsg(rtn))
      End If
   Else 'otherwise, if the key couldnt be opened
      regGetStringValue = ""       'return Error to the user
      throw Exception(regErrorMsg(rtn))
   End If
End If
End Function

Private Sub ParseKey(KeyName As String, Keyhandle As Long)

rtn = InStr(KeyName, "\") 'return if "\" is contained in the Keyname

If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   throw Exception("Incorrect Format:" + Chr(10) + Chr(10) + KeyName)  'display error to the user
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
   Keyhandle = GetMainKeyHandle(KeyName)
   KeyName = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
   Keyhandle = GetMainKeyHandle(Left(KeyName, rtn - 1)) 'seperate the Keyname
   KeyName = Right(KeyName, Len(KeyName) - rtn)
End If

End Sub
Public Function regCreateKey(subKey As String)

Call ParseKey(subKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegregCreateKey(MainKeyHandle, subKey, hKey) 'create the key
   If rtn = ERROR_SUCCESS Then 'if the key was created then
      rtn = API_baseMethods2_RegCloseKey(hKey)  'close the key
   End If
End If

End Function
Public Function regSetStringValue(subKey As String, Entry As String, Value As String)

Call ParseKey(subKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, subKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = API_baseMethods2_RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value)) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         throw Exception(regErrorMsg(rtn))
      End If
      rtn = API_baseMethods2_RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      throw Exception(regErrorMsg(rtn))
   End If
End If
End Function

'====================================================================================
'====================================================================================
'====================================================================================
Public Function DirectoryExists(Path As String) As Boolean
    DirectoryExists = (Dir(Path, vbDirectory) <> "")
End Function
'Public Function FileExists(Path As String) As Boolean
'    FileExists = (dir(Path, vbNormal) <> "")
'End Function
Public Function Exists(Path As String) As Boolean
    Exists = (Dir(Path) <> "")
End Function
Public Sub MakeTreeDirectories(ByVal Path As String, Optional Created As Boolean)
    If Not CheckPathValidation(Path, True, True) Then throw InvalidPathException
    Dim p As String, cPath As String
    p = mID(Path, 3) 'remove drive ex: E:\
    cPath = Left(Path, 3)
    Dim steps() As String
    If Left(p, 1) = "\" Then p = mID(p, 2)
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
    steps = Split(Trim(p), "\") '/ also checked ^
    Dim leng As Long
    leng = ArraySize(steps)
    Dim i As Long
    Created = False
    For i = 0 To leng - 1
        If (steps(i) = "") Or (Not CheckPathValidation(steps(i), False, False)) Then throw InvalidPathException: Created = True
    Next
    For i = 0 To leng - 1
        cPath = ConcatPath(cPath, steps(i))
        If Not DirectoryExists(cPath) Then Call MkDir(cPath)
    Next
End Sub
'====================================================================================
Public Function FileReadAllText(Path As String, Optional NewLineCharacter As String = vbCrLf) As String
    Dim fl As Long, cLine As String, isEof As Boolean
    fl = FreeFile
    Open Path For Input As #fl
    isEof = EOF(fl)
    While Not isEof
        Line Input #fl, cLine
        isEof = EOF(fl)
        If Not isEof Then
            FileReadAllText = FileReadAllText & cLine & NewLineCharacter
        Else
            FileReadAllText = FileReadAllText & cLine
        End If
    Wend
    Close #fl
End Function
Public Function FileReadAllLines(Path As String) As String()
    Dim fl As Long, ln() As String, lnCount As Long, cLine As String
    fl = FreeFile
    Open Path For Input As #fl
    While Not EOF(fl)
        Line Input #fl, cLine
        ReDim Preserve ln(lnCount)
        ln(lnCount) = cLine
        lnCount = lnCount + 1
    Wend
    Close #fl
    FileReadAllLines = ln()
End Function
Public Function FileWriteAllText(Path As String, Buffer As String)
    Dim fl As Long
    fl = FreeFile
    Open Path For Output As #fl
    Print #fl, Buffer
    Close #fl
End Function
Public Function FileWriteAllLines(Path As String, Lines() As String)
    If ArraySize(Lines) <= 0 Then Exit Function
    Dim fl As Long
    fl = FreeFile
    Open Path For Output As #fl
    Dim i As Long
    For i = LBound(Lines) To UBound(Lines)
        Print #fl, Lines(i)
    Next
    Close #fl
End Function
Public Function FileAppendAllText(Path As String, Buffer As String)
    Dim fl As Long
    fl = FreeFile
    Open Path For Append As #fl
    Print #fl, Buffer
    Close #fl
End Function
Public Function FileAppendAllLines(Path As String, Lines() As String)
    If ArraySize(Lines) <= 0 Then Exit Function
    Dim fl As Long
    fl = FreeFile
    Open Path For Append As #fl
    Dim i As Long
    For i = LBound(Lines) To UBound(Lines)
        Print #fl, Lines(i)
    Next
    Close #fl
End Function
'====================================================================================
Public Function GetAppDataPath(Optional ByVal System As Boolean = False) As String
    Dim retVal As String
    If System Then
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_SYSTEM, "AppData")
    Else
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_LOCAL, "Local AppData")
    End If
    GetAppDataPath = retVal
End Function
Public Function GetDesktopPath(Optional ByVal System As Boolean = False) As String
    Dim retVal As String
    If System Then
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_SYSTEM, "Desktop")
    Else
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_LOCAL, "Common Desktop")
    End If
    GetDesktopPath = retVal
End Function
Public Function GetLocalUserPath(Optional ByVal System As Boolean = False) As String
    Dim retVal As String
    If System Then
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_SYSTEM, "Common Documents")
    Else
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_LOCAL, "Personal")
    End If
    GetLocalUserPath = retVal
End Function
Public Function GetApplicationSpecifiedTempPath(Optional ByVal UseRevision As Boolean = True) As String
    Dim retVal As String
    retVal = GetTempPath
    retVal = ConcatPath(retVal, App.CompanyName)                                    'temp/CompanyName
    retVal = ConcatPath(retVal, App.ProductName)                                    'temp/CompanyName/ProductName
    retVal = ConcatPath(retVal, App.Major & "." & App.Minor)                        'temp/CompanyName/ProductName/0.0
    If UseRevision Then _
        retVal = ConcatPath(retVal, App.Major & "." & _
        App.Minor & "." & App.Revision)                                             'temp/CompanyName/ProductName/0.0/0.0.1
    GetApplicationSpecifiedTempPath = retVal
End Function
Public Function GetApplicationDataPath(Optional ByVal System As Boolean = False, Optional ByVal UseRevision As Boolean = True) As String
    Dim retVal As String
    retVal = GetAppDataPath(System)
    retVal = ConcatPath(retVal, App.CompanyName)                                    'user/CompanyName
    retVal = ConcatPath(retVal, App.ProductName)                                    'user/CompanyName/ProductName
    retVal = ConcatPath(retVal, App.Major & "." & App.Minor)                        'user/CompanyName/ProductName/0.0
    If UseRevision Then _
        retVal = ConcatPath(retVal, App.Major & "." & _
        App.Minor & "." & App.Revision)                                             'user/CompanyName/ProductName/0.0/0.0.1
    retVal = ConcatPath(retVal, "local_data")                                       'user/CompanyName/ProductName/0.0/0.0.1/local_data
    GetApplicationDataPath = retVal
End Function
Public Function GetApplicationSpecifiedTempPath_specified(CompanyName As String, ProductName As String, MajorVersion As Long, MinorVersion As Long, Revision As Long, Optional ByVal UseRevision As Boolean = True) As String
    Dim retVal As String
    retVal = GetTempPath
    retVal = ConcatPath(retVal, CompanyName)                                    'temp/CompanyName
    retVal = ConcatPath(retVal, ProductName)                                    'temp/CompanyName/ProductName
    retVal = ConcatPath(retVal, MajorVersion & "." & MinorVersion)              'temp/CompanyName/ProductName/0.0
    If UseRevision Then _
        retVal = ConcatPath(retVal, MajorVersion & "." & _
        MinorVersion & "." & Revision)                                          'temp/CompanyName/ProductName/0.0/0.0.1
    GetApplicationSpecifiedTempPath_specified = retVal
End Function
Public Function GetApplicationDataPath_specified(CompanyName As String, ProductName As String, MajorVersion As Long, MinorVersion As Long, Revision As Long, Optional ByVal System As Boolean = False, Optional ByVal UseRevision As Boolean = True) As String
    Dim retVal As String
    retVal = GetAppDataPath(System)
    retVal = ConcatPath(retVal, CompanyName)                                    'user/CompanyName
    retVal = ConcatPath(retVal, ProductName)                                    'user/CompanyName/ProductName
    retVal = ConcatPath(retVal, MajorVersion & "." & MinorVersion)              'user/CompanyName/ProductName/0.0
    If UseRevision Then _
        retVal = ConcatPath(retVal, MajorVersion & "." & _
        MinorVersion & "." & Revision)                                          'user/CompanyName/ProductName/0.0/0.0.1
    retVal = ConcatPath(retVal, "local_data")                                   'user/CompanyName/ProductName/0.0/0.0.1/local_data
    GetApplicationDataPath_specified = retVal
End Function

'---------------------------------------------

Public Function ReadINIFile(Path As String, Section As String, ByVal KeyName As String, Default As String) As Variant
    Dim str As String
    str = String(LARGELPSTR, Chr(0))
    If API_baseMethods2_GetPrivateProfileString(Section, KeyName, Default, str, LARGELPSTR, Path) = 0 Then _
        throw SystemCallFailureException
    str = GetLPSTR(str)
    If str = "" Then
        ReadINIFile = Default
    Else
        ReadINIFile = str
    End If
End Function
Public Function WriteINIFile(Path As String, Section As String, ByVal KeyName As String, KeyValue As String) As Boolean
    Dim ret As Long
    ret = API_baseMethods2_WritePrivateProfileString(Section, KeyName, KeyValue, Path)
    If ret = 0 Then
        WriteINIFile = True
    Else
        WriteINIFile = False
    End If
End Function
