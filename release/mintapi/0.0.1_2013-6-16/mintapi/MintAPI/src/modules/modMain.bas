Attribute VB_Name = "modMain"
Option Explicit

'by Ali Mousavi Kherad (LGPL-v3)
'Free to use and distribute but including my name and email as (alimousavikherad@gmail.com)!

Public Const MintAPI_FileCode_Mint As Long = 1953392973 ' Mint

Public Const APPLICATIONDOMAIN As String = "com.MintAPI"
Public Const APPLICATIONID As String = "MintAPI"
Public Const APP_VERSIONTAGS As String = "greenleaf"
Public Const APP_VERSIONSTRING As String = "0.0.1.2013 " & APP_VERSIONTAGS
Public Const APP_WEBSITE As String = "mintapi.com"
Public Const APP_SERVICEWEBSITE As String = "soap." & APP_WEBSITE
Public Const APP_HELPLINK As String = APP_WEBSITE & "help"
Public Const APP_UPDATELINK As String = APP_WEBSITE & "update"
Public Const APP_SUPPORTLINK As String = APP_WEBSITE & "support"

Public Const APP_GUID As String = ""

Public Const APP_PRODUCTCODE As String = "mintapi0000012013greenleafpxAB" ' 30 chars
Public Const APP_PRODUCTCODE50 As String = APP_PRODUCTCODE & "xxxxxxxxxxxxxxxxxxxx"

Public Const APP_REGISTRYPATH As String = "HKEY_LOCAL_MACHINE\SOFTWARE\MintAPI"
Public Const APP_REGISTRYPATH_USER As String = "HKEY_CURRENT_USER\SOFTWARE\MintAPI"


'New7API : determines compiling in win7 environment and remove kernel32 dll.

'=================================================================================================
'=================================================================================================
'=================================================================================================

#If New7API Then
    Public Declare Sub API_CopyMemory Lib "libkernel0.MintAPI.dll" Alias "lk_memcpy" (Destination As Any, Source As Any, ByVal Length As Long)
    Public Declare Sub API_ZeroMemory Lib "libkernel0.MintAPI.dll" Alias "lk_memclr" (Destination As Any, ByVal Length As Long)
#Else
    Public Declare Sub API_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Public Declare Sub API_ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
    Public Declare Sub memzero Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
#End If
    Public Declare Function API_VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
    Public Declare Function API_VarPtr Lib "msvbvm60" Alias "VarPtr" (Ptr As Any) As Long

'Public Declare Function API_Dialogs_Browse Lib "libkernel0.MintAPI.dll" Alias "Dialogs_Browse" (ByVal hWndParent As Long, ByVal strTitle As String, ByVal strPath As String, ByVal CreateNewFolderButton As Boolean, ByVal Flags As Long, Error As Long) As String

'Public Declare Function API_CallFunction0 Lib "libkernel0.MintAPI.dll" Alias "CallFunction0" (Func As Long) As Long
'Public Declare Function API_CallFunction1 Lib "libkernel0.MintAPI.dll" Alias "CallFunction1" (Func As Long, arg0 As Any) As Long
'Public Declare Function API_CallFunction2 Lib "libkernel0.MintAPI.dll" Alias "CallFunction2" (Func As Long, arg0 As Any, arg1 As Any) As Long
'Public Declare Function API_CallFunction3 Lib "libkernel0.MintAPI.dll" Alias "CallFunction3" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any) As Long
'Public Declare Function API_CallFunction4 Lib "libkernel0.MintAPI.dll" Alias "CallFunction4" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any) As Long
'Public Declare Function API_CallFunction5 Lib "libkernel0.MintAPI.dll" Alias "CallFunction5" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any) As Long
'Public Declare Function API_CallFunction6 Lib "libkernel0.MintAPI.dll" Alias "CallFunction6" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any, arg5 As Any) As Long
'Public Declare Function API_CallFunction7 Lib "libkernel0.MintAPI.dll" Alias "CallFunction7" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any, arg5 As Any, arg6 As Any) As Long
'Public Declare Function API_CallFunction8 Lib "libkernel0.MintAPI.dll" Alias "CallFunction8" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any, arg5 As Any, arg6 As Any, arg7 As Any) As Long
'Public Declare Function API_CallFunction9 Lib "libkernel0.MintAPI.dll" Alias "CallFunction9" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any, arg5 As Any, arg6 As Any, arg7 As Any, arg8 As Any) As Long
'Public Declare Function API_CallFunction10 Lib "libkernel0.MintAPI.dll" Alias "CallFunction10" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any, arg5 As Any, arg6 As Any, arg7 As Any, arg8 As Any, arg9 As Any) As Long
'
'Public Declare Function API_CallMethod0 Lib "libkernel0.MintAPI.dll" Alias "CallMethod0" (Func As Long) As Long
'Public Declare Function API_CallMethod1 Lib "libkernel0.MintAPI.dll" Alias "CallMethod1" (Func As Long, arg0 As Any) As Long
'Public Declare Function API_CallMethod2 Lib "libkernel0.MintAPI.dll" Alias "CallMethod2" (Func As Long, arg0 As Any, arg1 As Any) As Long
'Public Declare Function API_CallMethod3 Lib "libkernel0.MintAPI.dll" Alias "CallMethod3" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any) As Long
'Public Declare Function API_CallMethod4 Lib "libkernel0.MintAPI.dll" Alias "CallMethod4" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any) As Long
'Public Declare Function API_CallMethod5 Lib "libkernel0.MintAPI.dll" Alias "CallMethod5" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any) As Long
'Public Declare Function API_CallMethod6 Lib "libkernel0.MintAPI.dll" Alias "CallMethod6" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any, arg5 As Any) As Long
'Public Declare Function API_CallMethod7 Lib "libkernel0.MintAPI.dll" Alias "CallMethod7" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any, arg5 As Any, arg6 As Any) As Long
'Public Declare Function API_CallMethod8 Lib "libkernel0.MintAPI.dll" Alias "CallMethod8" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any, arg5 As Any, arg6 As Any, arg7 As Any) As Long
'Public Declare Function API_CallMethod9 Lib "libkernel0.MintAPI.dll" Alias "CallMethod9" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any, arg5 As Any, arg6 As Any, arg7 As Any, arg8 As Any) As Long
'Public Declare Function API_CallMethod10 Lib "libkernel0.MintAPI.dll" Alias "CallMethod10" (Func As Long, arg0 As Any, arg1 As Any, arg2 As Any, arg3 As Any, arg4 As Any, arg5 As Any, arg6 As Any, arg7 As Any, arg8 As Any, arg9 As Any) As Long

Public Declare Function API_Register_MintAPI_Lib Lib "libkernel0.MintAPI.dll" Alias "registerAPI" () As Boolean
Public Declare Sub API_lpMallocFree Lib "libkernel0.MintAPI.dll" Alias "lpMallocFree" (hndl As Long)

Public modMain_AllTimers As New Collection

Public Sub Main()
    'If Not API_Register_MintAPI_Lib Then throw InvalidCallException

    Call kernelMethods.InitializeCommonControls
    
    Call baseConstants.Initialize
    Call kernelMethods.Initialize
    Call baseMethods.Initialize
    Call baseMethods2.Initialize
    Call bitOperations.Initialize
    Call uiMethods.Initialize
    Call gdiMethods.Initialize
    Call shellMethods.Initialize
    Call baseExceptions.Initialize
    Call baseFiling.Initialize
    Call baseNetwork.Initialize
    
    Call mint_config.Initialize
    Call licensing.Initialize

'    Call Load(frmAbout)
'    frmAbout.unregVis.Enabled = True

    Call CheckIfNotInstalled
    Call DllLoadConfiguration
End Sub
Public Sub SetStabilityState(Optional ByVal State As Boolean = True)
    If State Then
        Call Load(frmStay)
    Else
        Call Unload(frmStay)
    End If
End Sub

Public Function Mtr(Key As String) As String
    Mtr = Key
End Function

Public Sub AboutMintAPI(Optional Modal As Boolean = False)
    Dim A As New frmAbout
    Call A.Show(IIf(Modal, 1, 0))
End Sub

Public Function MintAPIVersion() As Long
    Dim maj_Version As Long
    Dim min_Version As Long
    Dim rev_Version As Long
    maj_Version = tApplication.VersionMajor
    min_Version = tApplication.VersionMinor
    rev_Version = tApplication.VersionRevision
    If maj_Version > 255 Then GoTo outOfRangeExp_Err
    If min_Version > 255 Then GoTo outOfRangeExp_Err
    If rev_Version > 65535 Then GoTo outOfRangeExp_Err
    maj_Version = ShiftLeft(maj_Version And &HFF, 24)
    min_Version = ShiftLeft(min_Version And &HFF, 16)
    rev_Version = rev_Version And &HFFFF
    MintAPIVersion = (maj_Version + min_Version + rev_Version)
    Exit Function
outOfRangeExp_Err:
    MintAPIVersion = 0
End Function

Public Sub CheckIfNotInstalled()
'    If Registry(APP_REGISTRYPATH).Exists And _
'       Registry(APP_REGISTRYPATH).FieldExists("dll_path") Then
'       'exit if dll_path not valid.
'       If File(Registry(APP_REGISTRYPATH).GetValue("dll_path", "").toString()).Exists Then Exit Sub
'    End If



End Sub

Public Sub modMain_Timer_CallBack_Procedure(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
On Error GoTo ErrorHandler
    Const API_WM_TIMER = &H113
    Dim objTimer As Timer
    ' Make sure that the message is WM_TIMER.
    If uMsg = API_WM_TIMER Then
        ' It is a timer event.
        'Debug.Print "Timer: ", hwnd, uMsg, idEvent, dwTime
        
        For Each objTimer In modMain_AllTimers
            ' Execute the callback method in the class.
            Call objTimer.HandleCallBack(idEvent)
        Next objTimer
    End If
    Exit Sub
ErrorHandler:
    throw Exception(Err.Description)
End Sub

Public Function modMain_Thread_CallBack_Procedure(ByVal clsHandle As Long) As Long
    Dim Method As New Method
    Dim Thread As New Thread
    Call Method.Initialize("", clsHandle)
    Call Thread.Initialize(targetFuncHandle:=Method)
    Call Thread.Invoke
End Function

Public Function modMain_Console_CtrlEvent_Handler(ByVal CtrlType As Long) As Long
'Const CTRL_C_EVENT = 0
'Const CTRL_BREAK_EVENT = 1
'Const CTRL_CLOSE_EVENT = 2
''  3 is reserved!
''  4 is reserved!
'Const CTRL_LOGOFF_EVENT = 5
'Const CTRL_SHUTDOWN_EVENT = 6
'    If CtrlType = CTRL_C_EVENT Or CtrlType = CTRL_BREAK_EVENT Then _
'        mint_api_console_is_breaked = True
End Function
