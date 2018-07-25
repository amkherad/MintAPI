VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "libMethods2"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@PROJECT_LICENSE
Option Explicit
Option Base 0
Const CLASSID As String = "baseMethods2"

Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
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

Private Const REGPATH_SHELLFOLDERS_SYSTEM As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
Private Const REGPATH_SHELLFOLDERS_LOCAL As String = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"

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

Public Function regSetDWORDValue(SubKey As String, Entry As String, Value As Long)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = API_baseMethods2_RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
        throw Exception(ErrorMsg(rtn))        'display the error
      End If
      rtn = API_baseMethods2_RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
        throw Exception(ErrorMsg(rtn))  'display the error
   End If
End If

End Function
Public Function regGetDWORDValue(SubKey As String, Entry As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = API_baseMethods2_RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = API_baseMethods2_RegCloseKey(hKey)  'close the key
         regGetDWORDValue = lBuffer  'return the value
      Else                        'otherwise, if the value couldnt be retreived
         regGetDWORDValue = ""  'return Error to the user
         throw Exception(ErrorMsg(rtn))        'tell the user what was wrong
      End If
   Else 'otherwise, if the key couldnt be opened
      regGetDWORDValue = ""        'return Error to the user
      throw Exception(ErrorMsg(rtn))        'tell the user what was wrong
   End If
End If

End Function
Public Function regSetBinaryValue(SubKey As String, Entry As String, Value As String)

Call ParseKey(SubKey, MainKeyHandle)
Dim i As Long

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      lDataSize = Len(Value)
      ReDim ByteArray(lDataSize)
      For i = 1 To lDataSize
      ByteArray(i) = Asc(Mid$(Value, i, 1))
      Next
      rtn = API_baseMethods2_RegSetValueExB(hKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if the was an error writting the value
         throw Exception(ErrorMsg(rtn))
      End If
      rtn = API_baseMethods2_RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      throw Exception(ErrorMsg(rtn))
   End If
End If

End Function


Public Function regGetBinaryValue(SubKey As String, Entry As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened
      lBufferSize = 1
      rtn = API_baseMethods2_RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0, lBufferSize) 'get the value from the registry
      sBuffer = Space(lBufferSize)
      rtn = API_baseMethods2_RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = API_baseMethods2_RegCloseKey(hKey)   'close the key
         regGetBinaryValue = sBuffer 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         regGetBinaryValue = "" 'return Error to the user
         throw Exception(ErrorMsg(rtn))
      End If
   Else 'otherwise, if the key couldnt be opened
      regGetBinaryValue = "" 'return Error to the user
      throw Exception(ErrorMsg(rtn))
   End If
End If

End Function
Public Function regDeleteKey(Keyname As String)

Call ParseKey(Keyname, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, Keyname, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      rtn = API_baseMethods2_RegregDeleteKey(hKey, Keyname) 'delete the key
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
    Case "HKEY_CLASSES_ROOT"
        GetMainKeyHandle = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        GetMainKeyHandle = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
        GetMainKeyHandle = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
        GetMainKeyHandle = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA"
        GetMainKeyHandle = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG"
        GetMainKeyHandle = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function

Private Function ErrorMsg(lErrorCode As Long) As String
    
'If an error does accurr, and the user wants error messages displayed, then
'display one of the following error messages

Select Case lErrorCode
    Case 1009, 1015
        ErrorMsg = "The Registry Database is corrupt!"
    Case 2, 1010
        ErrorMsg = "Bad Key Name"
    Case 1011
        ErrorMsg = "Can't Open Key"
    Case 4, 1012
        ErrorMsg = "Can't Read Key"
    Case 5
        ErrorMsg = "Access to this key is denied"
    Case 1013
        ErrorMsg = "Can't Write Key"
    Case 8, 14
        ErrorMsg = "Out of memory"
    Case 87
        ErrorMsg = "Invalid Parameter"
    Case 234
        ErrorMsg = "There is more data than the buffer has been allocated to hold."
    Case Else
        ErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
End Select

End Function

Public Function regGetStringValue(SubKey As String, Entry As String)
Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      sBuffer = Space(255)     'make a buffer
      lBufferSize = Len(sBuffer)
      rtn = API_baseMethods2_RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = API_baseMethods2_RegCloseKey(hKey)  'close the key
         regGetStringValue = GetLPSTR(sBuffer)  'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         regGetStringValue = "" 'return Error to the user
         throw Exception(ErrorMsg(rtn))
      End If
   Else 'otherwise, if the key couldnt be opened
      regGetStringValue = ""       'return Error to the user
      throw Exception(ErrorMsg(rtn))
   End If
End If
End Function

Private Sub ParseKey(Keyname As String, Keyhandle As Long)
    
rtn = InStr(Keyname, "\") 'return if "\" is contained in the Keyname

If Left(Keyname, 5) <> "HKEY_" Or Right(Keyname, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   throw Exception("Incorrect Format:" + Chr(10) + Chr(10) + Keyname)  'display error to the user
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
   Keyhandle = GetMainKeyHandle(Keyname)
   Keyname = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
   Keyhandle = GetMainKeyHandle(Left(Keyname, rtn - 1)) 'seperate the Keyname
   Keyname = Right(Keyname, Len(Keyname) - rtn)
End If

End Sub
Public Function regCreateKey(SubKey As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegregCreateKey(MainKeyHandle, SubKey, hKey) 'create the key
   If rtn = ERROR_SUCCESS Then 'if the key was created then
      rtn = API_baseMethods2_RegCloseKey(hKey)  'close the key
   End If
End If

End Function
Public Function regSetStringValue(SubKey As String, Entry As String, Value As String)

Call ParseKey(SubKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = API_baseMethods2_RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value)) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         throw Exception(ErrorMsg(rtn))
      End If
      rtn = API_baseMethods2_RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      throw Exception(ErrorMsg(rtn))
   End If
End If
End Function

Private Function CheckPathValidation(Path As String, MakeTrueForm As Boolean, Optional CheckForDrive As Boolean = True) As Boolean
    If MakeTrueForm Then
        Path = Replace(Path, "/", "\")
    End If
    CheckPathValidation = True
End Function
'====================================================================================
'====================================================================================
'====================================================================================
Public Function DirectoryExists(Path As String) As Boolean
    DirectoryExists = Dir(Path, vbDirectory) <> ""
End Function
Public Function Exists(Path As String) As Boolean
    Exists = Dir(Path) <> ""
End Function
Public Sub MakeTreeDirectories(ByVal Path As String)
    If Not CheckPathValidation(Path, True, True) Then throw InvalidPathException
    Dim p As String, cPath As String
    p = Mid(Path, 3) 'remove drive ex: E:\
    cPath = Left(Path, 3)
    Dim steps() As String
    If Left(p, 1) = "\" Then p = Mid(p, 2)
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
    steps = Split(Trim(p), "\") '/ also checked ^
    Dim leng As Long
    leng = ArraySize(steps)
    Dim i As Long
    For i = 0 To leng - 1
        If (steps(i) = "") Or (Not CheckPathValidation(steps(i), False)) Then throw InvalidPathException
    Next
    For i = 0 To leng - 1
        cPath = ConcatPath(cPath, steps(i))
        If Not DirectoryExists(cPath) Then Call MkDir(cPath)
    Next
End Sub
Public Sub CreateFile(Path As String)
    Call MakeTreeDirectories(GetFilePath(Path))
    Dim fl As Long
    fl = FreeFile
    Open Path For Output As #fl
    Close #fl
End Sub
'====================================================================================
Public Function GetAppDataPath(Optional ByVal System As Boolean = False) As String
    Dim retVal As String
    If System Then
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_LOCAL, "AppData")
    Else
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_LOCAL, "Local AppData")
    End If
    GetAppDataPath = retVal
End Function
Public Function GetDesktopPath(Optional ByVal System As Boolean = False) As String
    Dim retVal As String
    If System Then
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_LOCAL, "Desktop")
    Else
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_SYSTEM, "Common Desktop")
    End If
    GetDesktopPath = retVal
End Function
Public Function GetLocalUserPath(Optional ByVal System As Boolean = False) As String
    Dim retVal As String
    If System Then
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_LOCAL, "AppData")
    Else
        retVal = regGetStringValue(REGPATH_SHELLFOLDERS_LOCAL, "Local AppData")
    End If
    GetLocalUserPath = retVal
End Function
Public Function GetApplicationSpecifiedTempPath(Optional ByVal UseRevision As Boolean = True) As String
    Dim retVal As String
    'retVal = GetTempPath
    retVal = ConcatPath(retVal, app.CompanyName)                                    'temp/CompanyName
    retVal = ConcatPath(retVal, app.ProductName)                                    'temp/CompanyName/ProductName
    retVal = ConcatPath(retVal, app.Major & "." & app.Minor)                        'temp/CompanyName/ProductName/0.0
    If UseRevision Then _
        retVal = ConcatPath(retVal, app.Major & "." & _
        app.Minor & "." & app.Revision)                                             'temp/CompanyName/ProductName/0.0/0.0.1
    GetApplicationSpecifiedTempPath = retVal
End Function
Public Function GetApplicationDataPath(Optional ByVal System As Boolean = False, Optional ByVal UseRevision As Boolean = True) As String
    Dim retVal As String
    retVal = GetAppDataPath(System)
    retVal = ConcatPath(retVal, app.CompanyName)                                    'user/CompanyName
    retVal = ConcatPath(retVal, app.ProductName)                                    'user/CompanyName/ProductName
    retVal = ConcatPath(retVal, app.Major & "." & app.Minor)                        'user/CompanyName/ProductName/0.0
    If UseRevision Then _
        retVal = ConcatPath(retVal, app.Major & "." & _
        app.Minor & "." & app.Revision)                                             'user/CompanyName/ProductName/0.0/0.0.1
    retVal = ConcatPath(retVal, "local_data")                                       'user/CompanyName/ProductName/0.0/0.0.1/local_data
    GetApplicationDataPath = retVal
End Function
