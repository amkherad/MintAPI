VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OperatingSystem1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "OperatingSystem"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING



Public Function osVersion() As osVersion
    'OSVersion = baseMethods.OSVersion
End Function
Public Function Version() As Version
    Dim V As New Version
        Dim vX As API_OSVERSIONINFO
    call v.Constructor0(vX.dwMajorVersion,vX.dwMinorVersion,vX.dwBuildNumber

    vX = baseMethods.GetVersionEx
    Version.PlatformID = vX.dwPlatformId
    Version.ServicePack = Trim(vX.szCSDVersion)
End Function

Public Function ServicePack() As String
    ServicePack = Trim(baseMethods.GetVersionEx.szCSDVersion)
End Function
Public Function GE(CheckForOS As osVersion) As Boolean
    'GE = (baseMethods.OSVersion >= CheckForOS)
End Function
Public Function Gr(CheckForOS As osVersion) As Boolean
    'Gr = (baseMethods.OSVersion > CheckForOS)
End Function
Public Function Eq(CheckForOS As osVersion) As Boolean
    'Eq = (baseMethods.OSVersion = CheckForOS)
End Function
Public Function SupportOS(CheckForOS As osVersion) As Boolean
'    If CheckForOS = OSHigher Then _
'        SupportOS = True: Exit Function
'    If CheckForOS = OSUnknown Then throw InvalidArgumentValueException("Unknown OSVersion.")
'    Dim IsHigher As Boolean
'    IsHigher = ((CheckForOS And OSHigher) = OSHigher)
'    If IsHigher Then
'        Dim lngChk As Long
'        lngChk = CheckForOS
'        lngChk = (lngChk Xor OSHigher)
'        SupportOS = (baseMethods.OSVersion >= lngChk)
'    Else
'        SupportOS = (baseMethods.OSVersion = CheckForOS)
'    End If
End Function


Public Function IsWow64Environment() As Boolean
    
End Function
