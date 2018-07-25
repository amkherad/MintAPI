VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="RegistryKey" tags="ascii unicode">
''  <summary>
''    Some class description.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
'Implements IEnumerable
'Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "RegistryKey" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"


Const REG_OPTION_NON_VOLATILE           As Long = 0
Const MAX_KEY_LENGTH                    As Long = 255
Const DEFAULT_VALUE_NAME                As String = ""


Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const Synchronize = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const STANDARD_RIGHTS_ALL = &H1F0000
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not Synchronize))
'Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not Synchronize))

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

Private Declare Function API_RegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long

#If Unicode Then
Private Declare Function API_RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function API_RegregCreateKey Lib "advapi32" Alias "RegregCreateKeyW" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function API_RegregDeleteKey Lib "advapi32" Alias "RegregDeleteKeyW" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function API_RegQueryValueExString Lib "advapi32" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function API_RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function API_RegSetValueExString Lib "advapi32" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function API_RegSetValueExLong Lib "advapi32" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function API_RegSetValueExByte Lib "advapi32" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
#Else
Private Declare Function API_RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function API_RegregCreateKey Lib "advapi32" Alias "RegregCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function API_RegregDeleteKey Lib "advapi32" Alias "RegregDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function API_RegQueryValueExString Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function API_RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function API_RegSetValueExString Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function API_RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function API_RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
#End If


Public Enum RegistryValueKind
    UnknownKind = 0
    StringKind = 1
    DWordKind = 4
    BinaryKind = 3
    MultiStringKind = 7
    ExpandStringKind = 2
    QWordKind = 11
End Enum

Public Enum RegistryValueOptions
    None = 0
    DoNotExpandEnvironmentVariables = 1
End Enum
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables/Events>
'
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Dim p_HKey As Long
Dim p_Name As String
Dim p_Writable As Boolean
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Constructors>
'
''<summary>
''  Basic constructor to use by MintAPI class provider,
''  When creating new instance of this class.
''</summary>
''<params>
''  <param name="Arguments">
''      A list of construction arguments.
''  </param>
''</params>
''<cargs order="0" count="1">
''  <carg type="long" name="ID"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal RH As RegistryHive, ByVal Name As String, ByVal Writable As Boolean)
    'Name = cString.Trim(Name, "\")
    This_Constructed = True
End Sub
'Friend Sub ConstructorCopy(ByVal That As {My Object Type})
'
'    This_Constructed = True
'End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Public Body Members>
'
'*********************************************
'*********************************************


Public Property Get Name() As String: THROW_InvalidHandle
    'Call VerifyHandle
    'Name = mName
End Property

Public Function OpenSubKey(ByVal Name As String, Optional ByVal Writable As Boolean = False) As RegistryKey
THROW_InvalidHandle
    
'    Name = mString.Trim(Name, "\")
'    Call THROW_InvalidKey(Name)
'
'    Dim Access As Long
'    If mWritable Then Access = KEY_ALL_ACCESS Else Access = KEY_READ
'
'    Dim Result As Long
'    Dim NewKey As Long
'    Result = API_RegOpenKeyEx(mHKey, Name, 0, Access, NewKey)
'    If Result <> ERROR_SUCCESS Then Exit Function
'
'    Set OpenSubKey = New RegistryKey
'    Call OpenSubKey.Constructor0(NewKey, Path.Combine(mName, Name), Writable)
End Function

Public Sub CloseKey()
'    If (Not IsSystemKey) And (mHKey <> vbNullPtr) Then
'        Call RegCloseKey(mHKey)
'        mHKey = vbNullPtr
'    End If
End Sub

Public Sub Flush()
'    Call VerifyHandle
'
'    Dim Result As Long
'    Result = RegFlushKey(mHKey)
'    If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey.Flush"
End Sub

Public Function CreateSubKey(ByVal SubKey As String) As RegistryKey
'    Call VerifyHandle
'    Call VerifyWritable
'
'    SubKey = cString.Trim(SubKey, "\")
'    Call VerifyKey(SubKey)
'
'    Dim Result As Long
'    Dim NewKey As Long
'    Result = API.RegCreateKeyEx(mHKey, SubKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, NewKey, 0)
'    If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey"
'
'    Set CreateSubKey = Cor.NewRegistryKey(NewKey, Path.Combine(mName, SubKey), True)
End Function

Public Sub DeleteSubKey(ByVal SubKey As String, Optional ByVal ThrowOnMissingSubKey As Boolean = True)
'    Call VerifyHandle
'    Call VerifyWritable
'
'    Dim rk As RegistryKey
'    Set rk = Me.OpenSubKey(SubKey)
'
'    If rk Is Nothing Then
'        If ThrowOnMissingSubKey Then _
'            throw Cor.NewArgumentException("The specified SubKey was not found.", "SubKey")
'    Else
'        If rk.SubKeyCount > 0 Then _
'            throw Cor.NewInvalidOperationException("Cannot delete. SubKey contains keys.")
'
'        Dim Result As Long
'        Result = API.regDeleteKey(mHKey, SubKey)
'        If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey"
'        rk.CloseKey
'    End If
End Sub

Public Sub DeleteSubKeyTree(ByVal SubKey As String)
'    Call VerifyHandle
'    Call VerifyWritable
'
'    SubKey = cString.Trim(SubKey, "\")
'    Call VerifyKey(SubKey)
'
'    Dim rk As RegistryKey
'    Set rk = Me.OpenSubKey(SubKey, True)
'
'    Dim SubKeys() As String
'    SubKeys = rk.GetSubKeyNames
'    If cArray.GetLength(SubKeys) > 0 Then
'        Dim i As Long
'        For i = 0 To UBound(SubKeys)
'            Call rk.DeleteSubKeyTree(SubKeys(i))
'        Next i
'    End If
'    Call rk.CloseKey
'
'    Dim Result As Long
'    Result = API.regDeleteKey(mHKey, SubKey)
'    If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey.DeleteSubKeyTree"
End Sub

Public Function GetSubKeyNames() As Vector
'    Call VerifyHandle
'
'    Dim Buf As String
'    Buf = String$(MAX_KEY_LENGTH + 1, 0)
'
'    Dim KeyCount As Long
'    KeyCount = Me.SubKeyCount
'
'    Dim Ret() As String
'    Ret = cArray.CreateInstance(ciString, KeyCount)
'
'    Do While KeyCount > 0
'        KeyCount = KeyCount - 1
'
'        Dim BufferSize As Long
'        BufferSize = MAX_KEY_LENGTH + 1
'
'        Dim Result As Long
'        Result = API.RegEnumKeyEx(mHKey, KeyCount, Buf, BufferSize, 0, vbNullString, 0, 0@)
'        If Result <> ERROR_SUCCESS Then IOError Result, "GetSubKeyNames"
'
'        Ret(KeyCount) = Left$(Buf, BufferSize)
'    Loop
'
'    GetSubKeyNames = Ret
End Function

Public Function GetValueNames() As Vector
'    Call VerifyHandle
'
'    Dim Buf As String
'    Buf = String$(Registry.MaxValueNameLength, 0)
'
'    Dim ValueCount As Long
'    ValueCount = Me.ValueCount
'
'    Dim Ret() As String
'    Ret = cArray.CreateInstance(ciString, ValueCount)
'
'    Do While ValueCount > 0
'        ValueCount = ValueCount - 1
'
'        Dim BufferSize As Long
'        BufferSize = Len(Buf)
'
'        Dim Result As Long
'        Result = API.RegEnumValue(mHKey, ValueCount, Buf, BufferSize, 0, 0, 0, 0)
'        If Result <> ERROR_SUCCESS Then IOError Result, "GetValueNames"
'
'        Ret(ValueCount) = Left$(Buf, BufferSize)
'    Loop
'
'    GetValueNames = Ret
End Function

Public Sub SetValue(ByVal Name As String, ByRef Value As Variant, Optional ByVal ValueKind As RegistryValueKind = UnknownKind)
'    Call VerifyWritable
'    Call VerifyHandle
'
'    If Len(Name) > Registry.MaxValueNameLength Then _
'        throw Cor.NewArgumentException("Registry Value name must not exceed " & Registry.MaxValueNameLength & " characters.", "Name")
'
'    Call ClearException
'    On Error GoTo errTrap
'
'    Dim Result As Long
'    If ValueKind = UnknownKind Then
'        Result = SetValueDirect(Name, Value)
'    Else
'        Result = SetValueWithConvert(Name, Value, ValueKind)
'    End If
'    On Error GoTo 0
'
'    If Result <> ERROR_SUCCESS Then IOError Result
'    Exit Sub
'
'errTrap:
'    Dim Ex As Exception
'    If Catch(Ex) Then
'        throw Ex
'    Else
'        throw Cor.NewArgumentException("The value did not match the registry type specified by ValueKind.", "value")
'    End If
End Sub

Public Function GetValue(ByVal Name As String, Optional ByRef Default As Variant, Optional ByVal Options As RegistryValueOptions) As Variant
'    Call VerifyHandle
'
'    If Len(Name) > Registry.MaxValueNameLength Then _
'        throw Cor.NewArgumentException("Registry Value name must not exceed " & Registry.MaxValueNameLength & " characters.", "Name")
'
'    Dim ValueType   As RegistryValueKind
'    Dim ValueSize   As Long
'    Dim Result      As Long
'    Result = API.RegQueryValueEx(mHKey, Name, ValueType, 0, ValueSize)
'
'    If Result = ERROR_FILE_NOT_FOUND Then
'        If Not IsMissing(Default) Then Call VariantCopyInd(GetValue, Default)
'        Exit Function
'    ElseIf Result <> ERROR_SUCCESS Then
'        IOError Result
'    End If
'
'    Select Case ValueType
'        Case StringKind
'            Dim sz As String
'            sz = String$(ValueSize, 0)
'            Result = API.RegQueryValueExStr(mHKey, Name, ValueType, sz)
'            If Result <> ERROR_SUCCESS Then IOError Result
'            GetValue = cString.TrimEnd(sz, vbNullChar)
'            Exit Function
'
'        Case DWordKind
'            Dim i As Long
'            Result = API.RegQueryValueEx(mHKey, Name, ValueType, VarPtr(i), 4)
'            If Result <> ERROR_SUCCESS Then IOError Result
'            GetValue = i
'
'        Case BinaryKind, REG_NONE
'            Dim Bytes() As Byte
'            If ValueSize > 0 Then
'                ReDim Bytes(0 To ValueSize - 1)
'                Result = API.RegQueryValueEx(mHKey, Name, ValueType, VarPtr(Bytes(0)), ValueSize)
'                If Result <> ERROR_SUCCESS Then IOError Result
'            Else
'                Bytes = Cor.NewBytes
'            End If
'
'            GetValue = Bytes
'            Exit Function
'
'        Case MultiStringKind
'            Dim MultiSZ As String
'            MultiSZ = String$(ValueSize, 0)
'            Result = API.RegQueryValueExStr(mHKey, Name, ValueType, MultiSZ)
'            If Result <> ERROR_SUCCESS Then IOError Result
'            GetValue = Split(cString.TrimEnd(MultiSZ, vbNullChar), vbNullChar)
'            Exit Function
'
'        Case ExpandStringKind
'            Dim ExpSZ As String
'            ExpSZ = String$(ValueSize, 0)
'            Result = API.RegQueryValueExStr(mHKey, Name, ValueType, ExpSZ)
'            If Result <> ERROR_SUCCESS Then IOError Result
'            ExpSZ = cString.TrimEnd(ExpSZ, vbNullChar)
'
'            If Options <> DoNotExpandEnvironmentVariables Then
'                ExpSZ = Environment.ExpandEnvironmentVariables(ExpSZ)
'            End If
'
'            GetValue = ExpSZ
'            Exit Function
'
'        Case QWordKind
'            Dim c As Currency
'            Result = API.RegQueryValueEx(mHKey, Name, ValueType, VarPtr(c), 8)
'            If Result <> ERROR_SUCCESS Then IOError Result
'            GetValue = c
'
'        Case Else
'            throw Cor.NewArgumentException("Unsupported Registry type.")
'    End Select
End Function

Public Sub DeleteValue(ByVal Name As String, Optional ByVal ThrowOnMissingValue As Boolean = True)
'    Call VerifyWritable
'    Call VerifyHandle
'
'    Dim Result As Long
'    Result = API.RegDeleteValue(mHKey, Name)
'    If (Result = ERROR_FILE_NOT_FOUND) And ThrowOnMissingValue Then _
'        throw Cor.NewArgumentException("Value was not found.", "Name")
End Sub

Public Property Get SubKeyCount() As Long
'    Call VerifyHandle
'
'    Dim Result As Long
'    Result = API.RegQueryInfoKey(mHKey, vbNullString, 0, 0, VarPtr(SubKeyCount), 0, 0, 0, 0, 0, 0, 0)
'    If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey"
End Property

Public Property Get ValueCount() As Long
'    Call VerifyHandle
'
'    Dim Result As Long
'    Result = API.RegQueryInfoKey(mHKey, vbNullString, 0, 0, 0, 0, 0, VarPtr(ValueCount), 0, 0, 0, 0)
'    If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey"
End Property

Public Function GetValueKind(ByVal Name As String) As RegistryValueKind
'    Call VerifyHandle
'
'    Dim Result As Long
'    Result = API.RegQueryValueEx(mHKey, Name, GetValueKind, 0, 0)
'    If Result <> ERROR_SUCCESS Then IOError Result, "GetValueKind"
End Function

Public Function GetLastWriteTime() 'As cDateTime
'    Call VerifyHandle
'
'    If Environment.IsNT Then
'        Dim Result  As Long
'        Dim Time    As Currency
'
'        Result = API.RegQueryInfoKey(mHKey, vbNullString, 0, 0, 0, 0, 0, 0, 0, 0, 0, VarPtr(Time))
'        If Result <> ERROR_SUCCESS Then IOError Result, "RegistryKey"
'
'        Set GetLastWriteTime = cDateTime.FromFileTime(Time)
'    Else
'        Set GetLastWriteTime = cDateTime.MinValue
'    End If
End Function



'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Basic Members>
'
Public Function Equals(Value As Variant) As Boolean
    Equals = Objects.Equals(Me, Value)
End Function
Public Function ToString() As String
    ToString = CLASSALIAS
End Function
Public Function GetHashCode() As Long
    GetHashCode = Objects.GetHashCode(Me)
End Function
Public Function MetaObject() As MetaObject
    If myMetaObject Is Nothing Then
        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    End If
    Set MetaObject = myMetaObject
End Function

'Public Function Clone() As Object
'
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
'Private Sub Class_Initialize()
'    ...
'End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Private Sub THROW_InvalidHandle() 'VBCorLib
    If p_HKey = vbNullPtr Then _
        throw Exps.ObjectDisposedException("The Registry Key is closed.") '"RegistryKey", "The Registry Key is closed."
    
    If p_HKey = HKEY_DYN_DATA Then _
        If Environment.IsNT Then throw Exps.InvalidOperationException
End Sub

Private Sub THROW_InvalidKey(ByRef Path As String) 'VBCorLib
    Dim Keys() As String
    Keys = Split(Path, "\")
    
    Dim i As Long
    For i = 0 To ArraySize(Keys) - 1
        If Len(Keys(i)) > MAX_KEY_LENGTH Then _
            throw Exps.InvalidArgumentException("Key names cannot exceed 255 characters.")
    Next
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo): Call Constructor(CInfo): End Sub
Private Function IObject_Equals(Value As Variant) As Boolean: IObject_Equals = Equals(Value): End Function
Private Function IObject_GetHashCode() As Long: IObject_GetHashCode = GetHashCode: End Function
Private Function IObject_MetaObject() As MetaObject: Set IObject_MetaObject = MetaObject: End Function
Private Function IObject_ToString() As String: IObject_ToString = ToString: End Function
Private Function IObject_Stub(ByVal ProxyInfo As ObjectProxyInfo) As Boolean
    
    IObject_Stub = False
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
'Private Function ICloneable_Clone() As Object
'    Set ICloneable_Clone = Clone
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
'Private Function IEnumerable_GetEnumerator() As IEnumerator
'
'End Function
'Private Function IEnumerable_NewEnum() As stdole.IUnknown
'
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
