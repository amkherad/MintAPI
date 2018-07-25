VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="Registry" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

''@ Registry system is written by Kelly Ethridge VBCorLib 2004
'' I just changed it's look to keep MintAPI's source and binary structure.

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Registry" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006

Public Enum RegistryHive
    rhClassesRoot = HKEY_CLASSES_ROOT
    rhCurrentConfig = HKEY_CURRENT_CONFIG
    rhCurrentUser = HKEY_CURRENT_USER
    rhDynData = HKEY_DYN_DATA
    rhLocalMachine = HKEY_LOCAL_MACHINE
    rhUsers = HKEY_USERS
    rhPerformanceData = HKEY_PERFORMANCE_DATA
End Enum

'These have no use just reserves some memory.
'User is responsible on extra calls to ClassesRoot and other hive provider methods.
'Private mClassesRoot        As RegistryKey
'Private mCurrentConfig      As RegistryKey
'Private mCurrentUser        As RegistryKey
'Private mDynData            As RegistryKey
'Private mLocalMachine       As RegistryKey
'Private mPerformanceData    As RegistryKey
'Private mUsers              As RegistryKey
Private p_MaxValueNameLength As Long
Private p_MaxValueLength     As Long

'Private p_SuspendExceptions As Boolean
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Private Function GetRegistryKey(ByVal RH As RegistryHive, ByVal Name As String, ByVal Writable As Boolean) As RegistryKey
    Set GetRegistryKey = New RegistryKey
    Call GetRegistryKey.Constructor0(RH, Name, Writable)
End Function


Public Property Get ClassesRoot() As RegistryKey
    Set ClassesRoot = GetRegistryKey(rhClassesRoot, "HKEY_CLASSES_ROOT", True)
End Property

Public Property Get CurrentConfig() As RegistryKey
    Set CurrentConfig = GetRegistryKey(rhCurrentConfig, "HKEY_CURRENT_CONFIG", True)
End Property

Public Property Get CurrentUser() As RegistryKey
    Set CurrentUser = GetRegistryKey(rhCurrentUser, "HKEY_CURRENT_USER", True)
End Property

Public Property Get DynData() As RegistryKey
    Set DynData = GetRegistryKey(rhDynData, "HKEY_DYN_DATA", True)
End Property

Public Property Get LocalMachine() As RegistryKey
    Set LocalMachine = GetRegistryKey(rhLocalMachine, "HKEY_LOCAL_MACHINE", True)
End Property

Public Property Get PerformanceData() As RegistryKey
    Set PerformanceData = GetRegistryKey(rhPerformanceData, "HKEY_PERFORMANCE_DATA", True)
End Property

Public Property Get Users() As RegistryKey
    Set Users = GetRegistryKey(rhUsers, "HKEY_USERS", True)
End Property

Public Property Get MintAPI(Optional ByVal System As Boolean = True) As RegistryKey
    Set MintAPI = GetSubKey(IIf(System, APP_REGISTRYPATH, APP_REGISTRYPATH_USER))
End Property
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

Public Function GetValue(ByVal KeyName As String, ByVal ValueName As String, Optional ByRef DefaultValue As Variant) As Variant
    Dim RegK As RegistryKey
    Set RegK = GetSubKey(KeyName)
    
    If Not RegK Is Nothing Then _
        Call MoveVariant(GetValue, RegK.GetValue(ValueName, DefaultValue))
End Function

Public Sub SetValue(ByVal KeyName As String, ByVal ValueName As String, ByRef Value As Variant)
    Dim RegK As RegistryKey
    Set RegK = GetSubKey(KeyName, True, True)
    
    If RegK Is Nothing Then _
        throw Exps.InvalidArgumentException("Could not save to registry key") ', "KeyName"
    
    Call RegK.SetValue(ValueName, Value)
End Sub

Public Function GetRootKey(ByVal KeyName As String) As RegistryKey
    If Len(KeyName) = 0 Then Exit Function
    KeyName = mString.Trim(KeyName, "\")
    
    Dim Keys() As String
    Keys = Split(KeyName, "\")
    
    'These extra checks is to make more underestanding to user's input.
    Select Case UCase$(Keys(0))
        Case "HKEY_LOCAL_MACHINE", "HKEY_LOCALMACHINE", "LOCAL_MACHINE", "LOCALMACHINE"
            Set GetRootKey = Me.LocalMachine
        Case "HKEY_CLASSES_ROOT", "HKEY_CLASSESROOT", "CLASSES_ROOT", "CLASSESROOT"
            Set GetRootKey = Me.ClassesRoot
        Case "HKEY_CURRENT_USER", "HKEY_CURRENTUSER", "CURRENT_USER", "CURRENTUSER"
            Set GetRootKey = Me.CurrentUser
        Case "HKEY_CURRENT_CONFIG", "HKEY_CURRENTCONFIG", "CURRENT_CONFIG", "CURRENTCONFIG"
            Set GetRootKey = Me.CurrentConfig
        Case "HKEY_DYN_DATA", "HKEY_DYNDATA", "DYN_DATA", "DYNDATA"
            Set GetRootKey = Me.DynData
        Case "HKEY_PERFORMANCE_DATA", "HKEY_PERFORMANCEDATA", "PERFORMANCE_DATA", "PERFORMANCEDATA"
            Set GetRootKey = Me.PerformanceData
        Case "HKEY_USERS", "USERS"
            Set GetRootKey = Me.Users
        Case Else
            'If Not p_SuspendExceptions Then
            throw Exps.InvalidOperationException
    End Select
End Function

Public Function GetSubKey(ByVal KeyName As String, Optional ByVal AutoCreate As Boolean = False, Optional ByVal Writable As Boolean = False) As RegistryKey
    If Len(KeyName) = 0 Then Exit Function
    
    Dim RetVal As RegistryKey
    Set RetVal = GetRootKey(KeyName)
    If RetVal Is Nothing Then Exit Function
    
    Dim i As Long
    i = InStr(KeyName, "\")
    If i = 0 Then Exit Function
        
    If AutoCreate Then
        Set GetSubKey = RetVal.CreateSubKey(Mid$(KeyName, i))
    Else
        Set GetSubKey = RetVal.OpenSubKey(Mid$(KeyName, i), Writable)
    End If
End Function

Public Property Get MaxValueNameLength() As Long
    MaxValueNameLength = p_MaxValueNameLength
End Property

Public Property Get MaxValueLength() As Long
    MaxValueLength = p_MaxValueLength
End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
Private Sub Class_Initialize()
    If Environment.IsNT Then 'VBCorLib
        p_MaxValueNameLength = 16383
        p_MaxValueLength = -1
    Else
        p_MaxValueNameLength = 255
        p_MaxValueLength = 16300
    End If
End Sub
