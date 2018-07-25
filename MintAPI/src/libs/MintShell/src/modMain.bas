Attribute VB_Name = "modMain"
Option Explicit

Public Const NRSS As String = "STR_VAL"
Public Const NRS As String = "'" & NRSS & "' is not recognized by MintAPI shell."
Public Const ANESS As String = "STR_ARG"
Public Const ANES As String = "Argument '" & ANESS & "' does not exists."
Public Const IVTSS As String = "STR_VAR"
Public Const IVTS As String = "Invalid type of variable '" & IVTSS & "'"


'commands:
'about                          about
'help                           help about
'install
'uninstall
'startup
'mount
'unmount
Public Declare Function API_ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function API_WaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Const STRINGTABLEID_ABOUT As Long = 100
Public Const STRINGTABLEID_HELP As Long = 101
Public Const STRINGTABLEID_LICENSE As Long = 103

Public stack As New ArgumentList
Public sec_signin As Boolean

Public tApp As tApplication

Public Sub Main()
    Call Application.trigConsole(sdBoth)
    Dim ContinueLoad As Boolean
    Dim appCommandArguments As String
    
    appCommandArguments = VBA.Command$
    
    'break on install and uninstall to prevent load assemblies.
    If appCommandArguments = "install" Then
        
        ContinueLoad = False
    ElseIf appCommandArguments = "uninstall" Then
        ContinueLoad = False
    Else
        ContinueLoad = True
    End If
    
    If ContinueLoad Then
        If Not Instance.RegisterLibraryLicense("", "") Then throw Exceptions.Exception("Unable to register application license.")
        Set tApp = New tApplication
        Call Application.StartApplication(tApp, ApplicationExecutable, App, VBA.Command$)
    End If
End Sub


Public Function GetStringTable(ID As Long) As String
    Dim str As String
    str = LoadResString(ID)
    str = Replace(str, "\n", vbCrLf)
    str = Replace(str, "\t", vbTab)
    GetStringTable = str
End Function

Public Function cSocket(cp As CodeParser) As clsSocket
    Set cSocket = New clsSocket
    Call cSocket.Initialize(cp)
End Function
Public Function cFile(cp As CodeParser) As clsFile
    Set cFile = New clsFile
    Call cFile.Initialize(cp)
End Function


Public Sub errMsg(strOut As String, Optional write_endl As Boolean = True)
    Console.ForeColor = ccRed
    stdout strOut
    Console.ForeColor = ccGrey
    If write_endl Then stdout endl
End Sub
Public Sub outMsg(strOut As String, Optional write_endl As Boolean = True)
    stdout strOut
    If write_endl Then stdout endl
End Sub
Public Sub inMsg(strIn As String)
    stdin strIn
End Sub

Public Sub GetVar(Name As String, outVariantType As Variant)
    If stack.Exists(Name) Then
        If stack.ArgumentType(Name) = vbObject Then
            Set outVariantType = stack(Name)
        Else
                outVariantType = stack(Name)
        End If
    Else
        outVariantType = Name
    End If
End Sub
Public Function GetVarType(Name As String) As String
    If stack.Exists(Name) Then
        Dim vType As VbVarType
        Dim arg0 As String
        Dim FE As Variant
        vType = stack.ArgumentType(Name)
        If vType = vbObject Then
            GetVarType = "object(" & TypeName(stack(Name)) & ")"
        ElseIf vType = vbArray Then
            FE = stack(Name)
            arg0 = TypeName(FE)
            GetVarType = Left(arg0, InStr(1, arg0, "(")) & "[" & ArraySize(FE) & "]"
        Else
            GetVarType = TypeName(stack(Name))
        End If
    Else
        GetVarType = "GetVarStr(N/A)"
    End If
End Function
Public Function GetVarStr(Name As String) As String
    If stack.Exists(Name) Then
        Dim vType As VbVarType
        Dim arg0 As String
        Dim FE As Variant
        vType = stack.ArgumentType(Name)
        If vType = vbObject Then
            GetVarStr = GetVarType(Name)
        ElseIf vType = vbArray Then
            FE = stack(Name)
            arg0 = TypeName(FE)
            GetVarStr = Left(arg0, InStr(1, arg0, "(")) & "[" & ArraySize(FE) & "]"
        Else
            GetVarStr = stack(Name)
        End If
    Else
        GetVarStr = "GetVarStr(N/A)"
    End If
End Function


Public Function CheckVar(varName As String) As Boolean
    CheckVar = stack.Exists(varName)
    If Not CheckVar Then _
        errMsg Replace(ANES, ANESS, varName)
End Function
Public Sub VariableNotExists(varName As String)
    errMsg Replace(ANES, ANESS, varName)
End Sub


Public Function outQuotPath(Path As String) As String ', Optional Comment As Boolean = False
    Dim bkSlashCount As Long, cbkSlashCounter As Long, ln As Long
    Dim i As Long
    outQuotPath = Trim(Path)
    ln = Len(outQuotPath)
    If Left(outQuotPath, 1) = """" Then
        If Right(outQuotPath, 1) = """" Then
            On Error GoTo err
'            If Comment Then
'                For cbkSlashCounter = ln - 1 To 1 Step -1
'                    If Mid(Path, cbkSlashCounter, 1) = "\" Then
'                        bkSlashCount = bkSlashCount + 1
'                    Else
'                        Exit For
'                    End If
'                Next
'            End If
            'If bkSlashCount Mod 2 = 0 Then
                outQuotPath = Mid(outQuotPath, 2, ln - 2)
            'End If
err:
        End If
    End If
End Function
Public Function ValNum(ByVal StrNum As String) As Double
    On Error GoTo errCatch
    If LCase(Left(StrNum, 2)) = "0x" Then
        StrNum = "&H" & Mid(StrNum, 3)
        GoTo errCatch
    End If
    If LCase(Left(StrNum, 1)) = "0" Then
        If InStr(StrNum, ".") = 0 Then
            StrNum = "&O" & Mid(StrNum, 2)
            GoTo errCatch
        End If
    End If
errCatch:
    ValNum = Val(StrNum)
End Function
