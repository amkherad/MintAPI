Attribute VB_Name = "RegAPI"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "RegAPI"

Private Declare Function API_RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function API_RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function API_RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function API_RegEnumValue Lib "advapi32" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function API_RegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Private Declare Function API_RegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function API_RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function API_RegSetValueExString Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function API_RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function API_RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function API_RegFlushKey Lib "advapi32" Alias "RegFlushKey" (ByVal hKey As Long) As Long
Private Declare Function API_RegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function API_RegDeleteKey Lib "advapi32" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function API_RegDeleteValue Lib "advapi32" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Enum ERR_Type
    ERR_Exception = -999999999
    ERR_Reflection = -999999998
    ERR_Abort = -999999997
End Enum

Public Type RegistryEntry
    EntryName As String
    EntryValue As Variant
End Type

Public Enum RegistryHiveKeys
    rgHiveKEY_CLASSES_ROOT = HKEY_CLASSES_ROOT
    rgHiveKEY_CURRENT_USER = HKEY_CURRENT_USER
    rgHiveKEY_LOCAL_MACHINE = HKEY_LOCAL_MACHINE
    rgHiveKEY_USERS = HKEY_USERS
    rgHiveKEY_PERFORMANCE_DATA = HKEY_PERFORMANCE_DATA
    rgHiveKEY_CURRENT_CONFIG = HKEY_CURRENT_CONFIG
    rgHiveKEY_DYN_DATA = HKEY_DYN_DATA
End Enum

Public Const RegAPI_ERROR_NO_MORE_ITEMS = 259&
Public Const RegAPI_ERROR_SUCCESS = 0&

Public Const RegAPI_REG_SZ = 1
Public Const RegAPI_REG_BINARY = 3
Public Const RegAPI_REG_DWORD = 4

Public Const RegAPI_KEY_READ = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                          ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                          ' SYNCHRONIZE))
Public Const RegAPI_REG_OPENED_EXISTING_KEY = &H2
Public Const RegAPI_KEY_WRITE = &H20006  '((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or
                           ' KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Public Const RegAPI_MAXWin9xLength As Long = 255
'Private vbLT As New LangTool

Dim inited As Boolean

Public Sub Initialize()
    If inited Then Exit Sub
    Call baseConstants.Initialize
    Call Exceptions.Initialize
    Call baseMethods.Initialize
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    Call baseConstants.Dispose(Force)
    inited = False
End Sub

Public Function AppendToPath(ByVal Path, ByVal PathToAdd, Optional Slash As String = "\") As String
    Dim P As String, A As String
    A = Trim(PathToAdd)
    P = Trim(Path)
    If A <> "" Then
        If Left(A, 1) = Slash Then
            A = Mid(A, 2)
        End If
    Else
        AppendToPath = P
        Exit Function
    End If
    If P <> "" Then
        If Right(P, 1) <> Slash Then
            P = P & Slash
        End If
    Else
        AppendToPath = A
        Exit Function
    End If
    AppendToPath = P & A
    If AppendToPath <> "" Then
        If Right(AppendToPath, 1) = Slash Then
            AppendToPath = Mid(AppendToPath, 1, Len(AppendToPath) - 1)
        End If
    End If
End Function

' Create a registry key, then close it
' Returns True if the key already existed, False if it was created

Public Function regCreateRegistryKey(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String) As Boolean
Dim HANDLE As Long, disposition As Long
If API_RegCreateKeyEx(hKey, KeyName, 0, 0, 0, 0, ByVal 0, HANDLE, disposition) Then
    throw SystemCallFailureException
    regCreateRegistryKey = False
Else
    ' Return True if the key already existed.
    regCreateRegistryKey = (disposition = RegAPI_REG_OPENED_EXISTING_KEY)
    ' Close the key.
    Call API_RegCloseKey(HANDLE)
    regCreateRegistryKey = True
End If

End Function
' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---

' Return True if a Registry key exists
Public Function regCheckRegistryKey(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String) As Boolean
Dim HANDLE As Long
Dim Ret As Long
' Try to open the key

Ret = API_RegOpenKeyEx(hKey, KeyName, 0, RegAPI_KEY_READ, HANDLE)
Select Case Ret
    Case 0:
        ' The key exists
        regCheckRegistryKey = True
        ' Close it before exiting
        API_RegCloseKey HANDLE
    Case 5:
        regCheckRegistryKey = False
        throw AccessDeniedException
    Case Else:
        regCheckRegistryKey = False
        throw AccessDeniedException
End Select
End Function
' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Enumerate registry keys under a given key
'
' returns a collection of strings

Public Function regEnumRegistryKeys(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String) As Collection
Dim HANDLE As Long
Dim Length As Long
Dim Index As Long
Dim subkeyName As String

' initialize the result collection
Set regEnumRegistryKeys = New Collection

' Open the key, exit if not found
If Len(KeyName) Then
    If API_RegOpenKeyEx(hKey, KeyName, 0, RegAPI_KEY_READ, HANDLE) Then Exit Function
    ' in all case the subsequent functions use hKey
    hKey = HANDLE
End If

Do
    ' this is the max length for a key name
    Length = 260
    subkeyName = Space$(Length)
    ' get the N-th key, exit the loop if not found
    If API_RegEnumKey(hKey, Index, subkeyName, Length) Then Exit Do
    
    ' add to the result collection
    subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
    Call regEnumRegistryKeys.Add(subkeyName)
    ' prepare to query for next key
    Index = Index + 1
Loop
' Close the key, if it was actually opened
If HANDLE Then Call API_RegCloseKey(HANDLE)
End Function

Public Function regSeekRegistryValue(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, _
    ByVal EntryName As String, ByRef EntryValue) As Boolean
Dim re() As RegistryEntry
Dim reLen As Long
Dim i As Long

re = regEnumRegistryValues(hKey, KeyName)

If Not IsEmptyArrayEntryValues(re) Then
    reLen = ArraySizeEntryValues(re)
    Call ResetVar(EntryValue)
    If reLen > 0 Then
        For i = 0 To reLen - 1
            If re(i).EntryName = EntryName Then
                EntryValue = re(i).EntryValue
                regSeekRegistryValue = True
                Exit Function
            End If
        Next i
    End If
End If

End Function
Private Sub ResetVar(ByRef Var)
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
Private Function IsArray(vArray As Variant) As Boolean

IsArray = (VarType(vArray) And vbArray) = vbArray

End Function

Private Function IsEmptyArray(vArray As Variant) As Boolean
'##BLOCK_DESCRIPTION Returns true if the variant passed is an empty array
IsEmptyArray = (ArraySize(vArray) = 0)
End Function
Private Function IsEmptyArrayEntryValues(reArray() As RegistryEntry) As Boolean
'##BLOCK_DESCRIPTION Returns true if the variant passed is an empty array
    On Error GoTo err
    IsEmptyArrayEntryValues = (UBound(reArray) - LBound(reArray) + 1) > 0
    Exit Function
err: IsEmptyArrayEntryValues = True
End Function
Private Function GetModuleName(ByRef SourceModule) As String

    Select Case VarType(SourceModule)
        Case vbObject:
        GetModuleName = TypeName(SourceModule)
    Case vbString:
        GetModuleName = SourceModule
    Case Else:
        GetModuleName = "<UndefinedModule>"
End Select

End Function
Private Function ArraySizeEntryValues(reArray() As RegistryEntry, Optional ByVal Dimension As Long = 1) As Long
'##PARAMETER_DESCRIPTION Dimension Sets the dimension of the array of to retrieve the size. If omitted _
 1 is assumed.
'##BLOCK_DESCRIPTION Returns the number of elements of the specified dimension of the array.
On Error GoTo ArrayEmpty
    ArraySizeEntryValues = UBound(reArray, Dimension) - LBound(reArray, Dimension) + 1
ArrayEmpty:
End Function
Private Function ArraySize(vArray As Variant, Optional ByVal Dimension As Long = 1) As Long
'##PARAMETER_DESCRIPTION Dimension Sets the dimension of the array of to retrieve the size. If omitted _
 1 is assumed.
'##BLOCK_DESCRIPTION Returns the number of elements of the specified dimension of the array.
On Error GoTo ArrayEmpty
If Not IsArray(vArray) Then throw InvalidArgumentTypeException

ArraySize = 1 + UBound(vArray, Dimension) - LBound(vArray, Dimension)
ArrayEmpty:
End Function
Public Function regEnumRegistryValues(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String) As RegistryEntry()
' ritorna un collezione di coppie (nome,valore) dove
' valore è del tipo relativo a quello rappresentato nel registro
Dim HANDLE As Long
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
Dim i As Long, s As String
Dim retVal() As RegistryEntry, rEntry As RegistryEntry

' Open the key, exit if not found
If Len(KeyName) Then
    If API_RegOpenKeyEx(hKey, KeyName, 0, RegAPI_KEY_READ, HANDLE) Then Exit Function
    ' in all case the subsequent functions use hKey
    hKey = HANDLE
End If

Do
    ValName = Space$(ValLen)
    lenValName = ValLen
    ReDim DataBuffer(DataBufferLen)
    lenDataBuffer = DataBufferLen
    res = API_RegEnumValue(hKey, Index, ValName, lenValName, 0&, ValType, DataBuffer(0), lenDataBuffer)
    If res = RegAPI_ERROR_SUCCESS Then
        
        Call ResetVar(Item)
        Erase byteArrayItem
        stringItem = ""
        longItem = 0
        
        Select Case ValType
            Case RegAPI_REG_BINARY ' 3 ==> ritorna un array di byte
                If lenDataBuffer > 0 Then
                    ReDim byteArrayItem(lenDataBuffer - 1)
                    For i = 0 To lenDataBuffer - 1
                        byteArrayItem(i) = DataBuffer(i)
                    Next i
                End If
                Item = byteArrayItem
                
            Case RegAPI_REG_DWORD ' 4 ==> ritorna un long
                If lenDataBuffer > 0 Then
                    s = Space$(lenDataBuffer)
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
                
        ReDim Preserve retVal(Index)
        rEntry.EntryName = Left(ValName, lenValName)
        rEntry.EntryValue = Item
        retVal(Index) = rEntry
        Index = Index + 1
    End If
Loop While (res = RegAPI_ERROR_SUCCESS)
If res = RegAPI_ERROR_NO_MORE_ITEMS Then 'ok
   regEnumRegistryValues = retVal
End If
    
End Function

Public Function regReadEntryValue(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, _
            ByVal EntryName As String, ByRef EntryValue As Variant) As Boolean
        ' cerca nella chiave specificata na voce; se la trova ritorna true
        ' ed aggiorna il valore di EntryValue
        Dim res() As RegistryEntry
        Dim i As Long
        Dim v As Variant, B() As Byte
        On Error GoTo xERR

    res = regEnumRegistryValues(hKey, KeyName)
    If Not IsEmptyArrayEntryValues(res) Then
            For i = LBound(res) To UBound(res)
            If Trim(LCase(res(i).EntryName)) = Trim(LCase(EntryName)) Then
                v = res(i).EntryValue
                regReadEntryValue = True
                    Select Case VarType(EntryValue)
                    Case vbString:  EntryValue = CStr(v)
                    Case vbBoolean: EntryValue = CBool(v)
                    Case vbLong:    EntryValue = CLng(v)
                    Case vbInteger: EntryValue = CInt(v)
                    Case vbSingle:    EntryValue = CSng(v)
                    Case vbDouble:    EntryValue = CDbl(v)
                    Case vbCurrency:    EntryValue = CCur(v)
                    Case vbArray + vbByte: EntryValue = v
                    Case Else
                    regReadEntryValue = False
                    throw InvalidVarTypeException
                    End Select
                    'EntryValue = v
                    Exit Function
                End If
        Next i
        End If
Exit Function
xERR:
    regReadEntryValue = False
End Function


' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Delete a registry key
'
' Under Windows NT it doesn't work if the key contains subkeys

Public Function regDeleteRegistryKey(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String) As Boolean
    regDeleteRegistryKey = (API_RegDeleteKey(hKey, KeyName) = 0)
End Function


' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Delete a registry value
'
' Return True if successful, False if the value hasn't been found

Public Function regDeleteRegistryValue(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, ByVal ValueName As String) As Boolean
Dim HANDLE As Long
Dim Ret As Long

' Open the key, exit if not found
If API_RegOpenKeyEx(hKey, KeyName, 0, RegAPI_KEY_WRITE, HANDLE) Then Exit Function
Call err.Clear
Ret = API_RegDeleteValue(HANDLE, ValueName)
' Delete the value (returns 0 if success)
'Debug.Print Ret, Err.LastDllError
regDeleteRegistryValue = (Ret = 0)
' Close the handle
API_RegCloseKey HANDLE

End Function



' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
' Write or Create a Registry value
' returns True if successful
'
' Use KeyName = "" for the default value
'
' Value can be an integer value (RegAPI_REG_DWORD), a string (RegAPI_REG_SZ)
' or an array of binary (RegAPI_REG_BINARY). Raises an error otherwise.

Public Function regSetRegistryValue(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, ByVal ValueName As String, Value As Variant) As Boolean
Dim HANDLE As Long
Dim lngValue As Long
Dim StrValue As String
Dim binValue() As Byte
Dim Length As Long
Dim retVal As Long

' Open the key, exit if not found
If API_RegOpenKeyEx(hKey, KeyName, 0, RegAPI_KEY_WRITE, HANDLE) <> 0 Then Exit Function

retVal = -1

' three cases, according to the data type in Value
Select Case VarType(Value)
    Case vbInteger, vbLong
        lngValue = CLng(Value)
        retVal = API_RegSetValueEx(HANDLE, ValueName, 0, RegAPI_REG_DWORD, lngValue, 4)
    Case vbString, vbBoolean
        StrValue = CStr(Value)
        retVal = API_RegSetValueEx(HANDLE, ValueName, 0, RegAPI_REG_SZ, ByVal StrValue, _
            Len(StrValue))
    Case vbArray + vbByte
        binValue = Value
        Length = UBound(binValue) - LBound(binValue) + 1
        retVal = API_RegSetValueEx(HANDLE, ValueName, 0&, RegAPI_REG_BINARY, _
            binValue(LBound(binValue)), Length)
    Case Else
        Call API_RegCloseKey(HANDLE)
        throw InvalidVarTypeException
End Select

' Close the key and signal success
Call API_RegCloseKey(HANDLE)
' signal success if the value was written correctly
regSetRegistryValue = (retVal = 0)

End Function


' --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---

Public Function regEraseRegistryTree(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, ByRef MaxLevelsErased As Long, Optional ByRef TotKeys As Long = 0, Optional ByRef DelKeys As Long = 0) As Boolean
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
        If Not regCheckRegistryKey(hKey, ByVal KeyName) Then
            regEraseRegistryTree = False
            Exit Function
        End If
        notFirstTime = True
        MaxLevelsErased = 0 ' init first time
        TotKeys = 1
        DelKeys = 0
    End If
    
    Set regEntries = regEnumRegistryKeys(hKey, KeyName)
        
    If regEntries.Count = 0 Then ' it's a leaf
        If regCheckRegistryKey(hKey, KeyName) Then
            Ret = regDeleteRegistryKey(hKey, KeyName)
            If Ret Then DelKeys = DelKeys + 1
            regEraseRegistryTree = Ret
        End If
        Exit Function
    Else
        TotKeys = TotKeys + regEntries.Count
        ActLevel = MaxLevelsErased + 1
        MaxLE = ActLevel
        ActLevelDelCount = 0
        For Each rEntry In regEntries
            PassLevel = ActLevel
            Ret = regEraseRegistryTree(hKey, KeyName & "\" & rEntry, PassLevel, TotKeys, DelKeys)
            If Ret Then ActLevelDelCount = ActLevelDelCount + 1
            If (PassLevel > MaxLE) And Ret Then MaxLE = PassLevel
        Next
        MaxLevelsErased = MaxLE
        Ret = regDeleteRegistryKey(hKey, KeyName)
        If Ret Then DelKeys = DelKeys + 1
        If ActLevel = 1 Then
            regEraseRegistryTree = (TotKeys = DelKeys)
            notFirstTime = False
        Else
            regEraseRegistryTree = (ActLevelDelCount = regEntries.Count)
        End If
    End If
End Function

Public Function regSeekRegistryLeafs(ByVal hKey As RegistryHiveKeys, ByVal KeyName As String, _
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
Set EnumKeys = regEnumRegistryKeys(hKey, KeyName)

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
            Set SeekKeys = regSeekRegistryLeafs(hKey, Key)
        Else
            Set SeekKeys = regSeekRegistryLeafs(hKey, Key, MaxEntries)
        End If

        If SeekKeys.Count > 0 Then
            For Each SeekKey In SeekKeys
                Call TotKeys.Add(SeekKey)
            Next
        End If
    Next
End If
Set regSeekRegistryLeafs = TotKeys

Level = Level - 1
If Level = 0 Then KeysFound = 0

End Function

Public Function regFlushRegistryChanges(ByVal hKey As RegistryHiveKeys) As Boolean
    regFlushRegistryChanges = (API_RegFlushKey(hKey) = 0)
End Function
