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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Registry"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


'    HKEY_CLASSES_ROOT
'    HKEY_CURRENT_USER
'    HKEY_LOCAL_MACHINE
'    HKEY_USERS
'    HKEY_PERFORMANCE_DATA
'    HKEY_CURRENT_CONFIG
'    HKEY_DYN_DATA

Dim i_key As String
Dim lastErrorID As Long

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
    VL = sp.cols(cIndex).Value

    If sp.colsCount > 1 Then
        sp.cols(cIndex).Include = False
    Else
        Select Case VL
            Case "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE", _
                 "HKEY_USERS", "HKEY_PERFORMANCE_DATA", "HKEY_CURRENT_CONFIG", "HKEY_DYN_DATA"
                If cIndex = 0 Then
                    sp.cols(cIndex).Include = True
                Else
                    sp.cols(cIndex).Include = False
                End If
            Case Else
                sp.cols(cIndex).Include = False
        End Select
    End If

    Dim i As Long, cPath As String
    For i = 0 To sp.colsCount - 1
        If sp.cols(i).Include Then _
            cPath = cPath & sp.cols(i).Value & "\"
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
