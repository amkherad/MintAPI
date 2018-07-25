VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"


'Ideas from Me , .net Framework , Qt

Option Base 0
Option Explicit
Const CLASSID As String = "Registry"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


'Implements IAliasable


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



Private Declare Function API_baseMethods2_RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function API_baseMethods2_RegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Private Declare Function API_baseMethods2_RegregCreateKey Lib "advapi32" Alias "RegregCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function API_baseMethods2_RegregDeleteKey Lib "advapi32" Alias "RegregDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function API_baseMethods2_RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function API_baseMethods2_RegQueryValueExA Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function API_baseMethods2_RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function API_baseMethods2_RegSetValueExA Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function API_baseMethods2_RegSetValueExB Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long

'    HKEY_CLASSES_ROOT
'    HKEY_CURRENT_USER
'    HKEY_LOCAL_MACHINE
'    HKEY_USERS
'    HKEY_PERFORMANCE_DATA
'    HKEY_CURRENT_CONFIG
'    HKEY_DYN_DATA

Dim i_key As String
Dim lastErrorID As Long

'---------------------------------------
Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
'---------------------------------------

Public Function regSetDWORDValue(subKey As String, Entry As String, Value As Long)

Call ParseKey(subKey, MainKeyHandle)

If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, subKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      rtn = API_baseMethods2_RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
        throw Exps.Exception(regErrorMsg(rtn))        'display the error
      End If
      rtn = API_baseMethods2_RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
        throw Exps.Exception(regErrorMsg(rtn))  'display the error
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
         throw Exps.Exception(regErrorMsg(rtn))        'tell the user what was wrong
      End If
   Else 'otherwise, if the key couldnt be opened
      regGetDWORDValue = ""        'return Error to the user
      throw Exps.Exception(regErrorMsg(rtn))        'tell the user what was wrong
   End If
End If

End Function
Public Function regSetBinaryValue(subKey As String, Entry As String, Value As String)

Call ParseKey(subKey, MainKeyHandle)
Dim i As Long, lDataSize As Long
If MainKeyHandle Then
   rtn = API_baseMethods2_RegOpenKeyEx(MainKeyHandle, subKey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
      lDataSize = Len(Value)
      Dim ByteArray() As Byte
      ReDim ByteArray(lDataSize)
      For i = 1 To lDataSize
      ByteArray(i) = Asc(Mid$(Value, i, 1))
      Next
      rtn = API_baseMethods2_RegSetValueExB(hKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if the was an error writting the value
         throw Exps.Exception(regErrorMsg(rtn))
      End If
      rtn = API_baseMethods2_RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      throw Exps.Exception(regErrorMsg(rtn))
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
         throw Exps.Exception(regErrorMsg(rtn))
      End If
   Else 'otherwise, if the key couldnt be opened
      regGetBinaryValue = "" 'return Error to the user
      throw Exps.Exception(regErrorMsg(rtn))
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
        throw Exps.InvalidArgumentValueException("Invalid Registry Hive Key.")
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
        regErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
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
         throw Exps.Exception(regErrorMsg(rtn))
      End If
   Else 'otherwise, if the key couldnt be opened
      regGetStringValue = ""       'return Error to the user
      throw Exps.Exception(regErrorMsg(rtn))
   End If
End If
End Function

Private Sub ParseKey(KeyName As String, Keyhandle As Long)

rtn = InStr(KeyName, "\") 'return if "\" is contained in the Keyname

If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   throw Exps.Exception("Incorrect Format:" + Chr(10) + Chr(10) + KeyName)  'display error to the user
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
         throw Exps.Exception(regErrorMsg(rtn))
      End If
      rtn = API_baseMethods2_RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      throw Exps.Exception(regErrorMsg(rtn))
   End If
End If
End Function

'========================================
'========================================
'========================================
'========================================
'========================================
'========================================
'========================================

Friend Sub Initialize(Optional ByVal Key As String = "")
    i_key = Key
End Sub

Public Function ClassesRoot() As String
    ClassesRoot = "HKEY_CLASSES_ROOT"
End Function
Public Function CurrentUser() As String
    CurrentUser = "HKEY_CURRENT_USER"
End Function
Public Function LocalMachine() As String
    LocalMachine = "HKEY_LOCAL_MACHINE"
End Function
Public Function Users() As String
    Users = "HKEY_USERS"
End Function
Public Function PerformanceData() As String
    PerformanceData = "HKEY_PERFORMANCE_DATA"
End Function
Public Function CurrentConfig() As String
    CurrentConfig = "HKEY_CURRENT_CONFIG"
End Function
Public Function DynamicData() As String
    DynamicData = "HKEY_DYN_DATA"
End Function

Public Function ConcatRegistry(Root As String, SubKeys As String) As String
    ConcatRegistry = ConcatPath(Root, SubKeys)
End Function

Public Sub SetKey(Key As String)
    i_key = Key
End Sub

Public Sub SetValue(FieldName As String, Value)
    
End Sub
Public Function GetValue(FieldName As String, Optional DefaultValue As String = "") As ByteArray
    
End Function

Public Function cd(Key As String) As Registry
    i_key = ConcatPath(i_key, Key)
    Set cd = Me
End Function
Public Function cdup() As Registry
    If i_key = "" Then Exit Function
    Dim sp As BASEMETHODS_SAFEPATH
    sp = SplitPathToSafePath(i_key)
    If sp.colsCount <= 0 Then Exit Function

    Dim VL As String, cIndex As Long
    cIndex = sp.colsCount - 1
    VL = sp.Cols(cIndex).Value

    If sp.colsCount > 1 Then
        sp.Cols(cIndex).Include = False
    Else
        Select Case VL
            Case "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE", _
                 "HKEY_USERS", "HKEY_PERFORMANCE_DATA", "HKEY_CURRENT_CONFIG", "HKEY_DYN_DATA"
                If cIndex = 0 Then
                    sp.Cols(cIndex).Include = True
                Else
                    sp.Cols(cIndex).Include = False
                End If
            Case Else
                sp.Cols(cIndex).Include = False
        End Select
    End If

    Dim i As Long, cPath As String
    For i = 0 To sp.colsCount - 1
        If sp.Cols(i).Include Then _
            cPath = cPath & sp.Cols(i).Value & "\"
    Next
    
    i_key = cPath
    
    Set cdup = Me
End Function
Public Function Key(strKey As String) As Registry
    Dim R As New Registry
    Call R.Initialize(strKey)
    Set Key = R
End Function
Public Function subKey(strSubKey As String) As Registry
    Dim R As New Registry
    Call R.Initialize(i_key)
    Set subKey = R.cd(strSubKey)
End Function

Public Function Error(lErrorCode As Long) As String
    Error = regErrorMsg(lErrorCode)
End Function
Public Function LastError() As Long
    LastError = lastErrorID
End Function

Public Sub CreateKey(Optional OptionalKey)
    
End Sub
Public Sub DeleteKey(Optional OptionalKey)
    
End Sub
Public Sub DeleteField(FieldName As String, Optional OptionalKey)
    
End Sub

Public Function FieldExists(FieldName As String, Optional OptionalKey) As Boolean
    
End Function
Public Function Exists(Optional OptionalKey) As Boolean
    
End Function


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
