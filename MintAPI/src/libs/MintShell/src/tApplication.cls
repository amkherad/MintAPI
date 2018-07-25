VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const PA_WORD As String = "sps23"

Const SLOT_IF As String = "if"
Const SLOT_ENDIF As String = "endif"
Const SLOT_LOOP As String = "loop"
Const SLOT_ENDLOOP As String = "endloop"

Const SLOT_DEF As String = "def"
Const SLOT_UDEF As String = "udef"
Const SLOT_SET As String = "set"
Const SLOT_GET As String = "get"
Const SLOT_ISDEF As String = "isdef"
Const SLOT_TYPEOF As String = "typeof"
Const SLOT_ECHO As String = "echo"
Const SLOT_EXEC As String = "exec"
Const SLOT_OUT As String = "out"
Const SLOT_IN As String = "in"
Const SLOT_CD As String = "cd"
Const SLOT_CHANGEDIR As String = "chdir"
Const SLOT_STARTUP As String = "startup"
Const SLOT_EXIT As String = "exit"
Const SLOT_OPEN As String = "open"
Const SLOT_CLOSE As String = "close"
Const SLOT_SCRIPT As String = "script"
Const SLOT_MKDIR As String = "mkdir"
Const SLOT_RMDIR As String = "rmdir"
Const SLOT_CDIR As String = "cdir"
Const SLOT_DEL As String = "del"
Const SLOT_VB6INSTALL As String = "vb6install"
Const SLOT_SET_POWERSTATE As String = "set_powerstate"
Const SLOT_CANCEL_POWERSTATE As String = "cancel_powerstate"

Const SLOT_PLGINSTL As String = "plginstl"
Const SLOT_INSTALLPLUGIN As String = "installplugin"
Const SLOT_PLGUINSTL As String = "plguinstl"
Const SLOT_UNINSTALLPLUGIN As String = "uninstallplugin"

Const SLOT_ABOUT As String = "about"
Const SLOT_HELP As String = "help"
Const SLOT_GUID As String = "guid"
Const SLOT_PING As String = "ping"
Const SLOT_LICENSE As String = "license"
Const SLOT_WIN_SHELL As String = "win_shell"
Const SLOT_MOUNT As String = "mount"
Const SLOT_UNMOUNT As String = "unmount"
Const SLOT_SHOWVARS As String = "showvars"
Const SLOT_GETSTRING As String = "getstr"
Const SLOT_LIST As String = "list"
Const SLOT_NEW As String = "new"
Const SLOT_DELETE As String = "release"
Const SLOT_REDEFINE As String = "redefine"

Const SLOT_FILE_DEFINE As String = "file"
Const SLOT_FILE_OPEN As String = "file_open"
'Const SLOT_FILE_CLOSE As String = "file_close"
Const SLOT_FILE_WRITE As String = "file_write"
Const SLOT_FILE_READ As String = "file_read"
Const SLOT_FILE_GETLENGTH As String = "file_getlen"
Const SLOT_FILE_SETLENGTH As String = "file_setlen"
Const SLOT_FILE_GETPOSITION As String = "file_getpos"
Const SLOT_FILE_SETPOSITION As String = "file_setpos"

Const SLOT_SOCKET_DEFINE As String = "socket"
'Const SLOT_SOCKET_CLOSE As String = "sock_close"
Const SLOT_SOCKET_CONNECT As String = "sock_connect"
Const SLOT_SOCKET_SEND As String = "sock_send"
Const SLOT_SOCKET_BIND As String = "sock_bind"
Const SLOT_SOCKET_RECIEVE As String = "sock_recv"
Const SLOT_SOCKET_LISTEN As String = "sock_listen"
Const SLOT_SOCKET_SHUTDOWN As String = "sock_shutdown"
Const SLOT_SOCKET_DISCONNECT As String = "sock_disconnect"

Private Type COMMANDRECORDS
    Coms() As String
    ComsCount As Long
End Type
Private Enum AppEnvironment
    aeShell
    aeDeveloper
End Enum

Implements IApplication
Implements ISupportSlot

Dim FE As Variant, i As Long, dlLen As Long, dDouble As Double
Dim arg0 As String, arg1 As String, buf0 As String
Dim chC As Byte, ITS As ITargetStream, vType As VbVarType
Dim objBuf As Object

Dim env_mode As AppEnvironment

Private Sub Class_Initialize()
    env_mode = aeShell
End Sub


Private Sub IApplication_Main(args As Variant)
    Dim CommandString As String
    CommandString = args
    
    Console.Title = "MintAPI shell [version " & App.Major & "." & App.Minor & "." & App.Revision & ".2013]"
    
    Dim cp As New CodeParser
    Call cp.ParseCommand(CommandString)
    
    If CommandString <> "" Then
        Dim CommandName As String
        CommandName = cp.CommandName
        Call ExecuteCommand(cp)
    End If
    
    Call AppendDefaultVariables
    
    Dim Continue As Boolean, strRead As String
    
    outMsg "MintAPI shell [" & "Version " & App.Major & "." & App.Minor & "." & App.Revision & ".2013]"
    outMsg "Copyright (C) by Ali Mousavi Kherad"
    
    Continue = True
    While Continue
        If env_mode = aeShell Then
            outMsg "$MintAPI: " & LCase(CurrentDirectory.AbsolutePath) & ">", False
        Else
            'outMsg "<" & LCase(CurrentDirectory.AbsolutePath) & ">"
            outMsg "Command>", False
        End If
        stdin strRead
        If Trim(strRead) = "" Then GoTo lastofWhile
        outMsg ""
        Select Case strRead
            Case "exit", "quit"
                Continue = False
            Case Else
                Call cp.ParseCommand(strRead)
                If Not ExecuteCommand(cp, Continue) Then
                    errMsg Replace(NRS, NRSS, cp.CommandName)
                End If
        End Select
        If Console.IsKeyPressed Then _
            If Console.ReadKey = 27 Then Continue = False
lastofWhile:
        outMsg ""
    Wend
End Sub

Private Sub AppendDefaultVariables()
    stack.AppendArgument Console, "console"
    stack.AppendArgument Me, "parser"
    stack.AppendArgument endl, "endl"
    stack.AppendArgument sclr, "clr"
End Sub

Public Function ExecuteCommand(cp As CodeParser, Optional Continue As Boolean) As Boolean
    Static IfEntries() As Boolean, IfEntriesCount As Long
    Static Loops() As COMMANDRECORDS, LoopsCount As Long
    Dim stt_exec As Boolean
    On Error GoTo errCatch
    stt_exec = True
    buf0 = cp.CommandName
    
'    arg0 = cp.GetArgString
'    If InStr(1, arg0, ";") > 0 Then
'        Dim splitStrs() As String
'        splitStrs = Split(arg0, ";")
'        ExecuteCommand = True
'        For i = 0 To ArraySize(splitStrs) - 1
'            Set cp = Nothing
'            Set cp = New CodeParser
'            If i = 0 Then
'                Call cp.ParseCommand(buf0 & " " & splitStrs(i))
'            Else
'                Call cp.ParseCommand(splitStrs(i))
'            End If
'            If Not ExecuteCommand(cp, Continue) Then
'                outMsg "Execution breaked..."
'                ExecuteCommand = False
'                Exit For
'            End If
'        Next
'        Exit Function
'    End If
    
    If cp.CountArguments = 1 Then
        If cp.GetArg(0) = "?" Then
            Call ShowHelpOf(buf0)
            GoTo dropCommand
        End If
    End If
    
    If Len(buf0) = 1 Then
        Call ChDrive(buf0)
        GoTo dropCommand
    End If
    
    If Left(buf0, 3) = "end" Then
        Select Case buf0
            Case SLOT_ENDIF
                
            Case SLOT_ENDLOOP
                
        End Select
    End If
    
    Select Case buf0
        Case ""
        
            ExecuteCommand = True
            Exit Function
        '=====================================
        Case SLOT_IF
        
            Dim IfValue As Boolean, IfNot As Boolean, Ifoperation As String
            Call cp.ensurege(1)
            arg0 = cp.GetArg(0)
            Ifoperation = "_"
            For i = 1 To cp.CountArguments - 1
                arg1 = cp.GetArg(i)
                'if _g x 4
                Select Case arg1
                    Case "_"
                        Ifoperation = "_"
                    Case "e"
                        Ifoperation = "e"
                    Case "ne"
                        Ifoperation = "ne"
                    Case "g"
                        Ifoperation = "g"
                    Case "ge"
                        Ifoperation = "ge"
                    Case "l"
                        Ifoperation = "l"
                    Case "le"
                        Ifoperation = "le"
                    Case "not"
                        'if not x
                        IfNot = True
                    Case Else
                        
                End Select
            Next
'            ReDim Preserve IfEntries(IfEntriesCount)
'            IfEntries (IfEntriesCount)
        '=====================================
        Case SLOT_LOOP
        
            
        '=====================================
        Case "exp"
        
            If Not Func_Exp(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case "add"
        
            If Not Func_Add(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case "sub"
        
            If Not Func_Sub(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case "mul"
        
            If Not Func_Mul(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case "div"
        
            If Not Func_Div(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case "pow"
        
            If Not Func_Pow(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case "sqr"
        
            If Not Func_Sqr(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case "not"
        
            If Not Func_Not(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case "or"
        
            If Not Func_Or(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case "xor"
        
            If Not Func_Xor(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_DEF
        
        On Error GoTo def_err_Catch
            For i = 0 To cp.CountArguments - 1
                Call DefineVariable(CStr(outQuotPath(cp.GetArg(i))))
                GoTo Continue_def_err_Catch
def_err_Catch:
                errMsg err.Description
Continue_def_err_Catch:
            Next
        On Error GoTo errCatch
        '=====================================
        Case SLOT_UDEF
        
            For i = 0 To cp.CountArguments - 1
                arg0 = CStr(outQuotPath(cp.GetArg(i)))
                If CheckVar(arg0) Then
                    Call stack.RemoveArgumentbyName(arg0)
                End If
            Next
        On Error GoTo errCatch
        '=====================================
        Case SLOT_SET
        
            If Not Func_Set(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_GET
            
            If Not Func_Get(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_OUT
        
            If Not Func_OUT(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_ECHO
        
            If Not Func_Echo(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_IN
        
            If Not Func_IN(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_ISDEF
        
            If Not Func_IsDef(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_TYPEOF
        
            If Not Func_TypeOf(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_EXEC
        
            If Not Func_Exec(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_NEW
        
            If Not Func_New(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_DELETE
        
            If Not Func_Delete(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_WIN_SHELL
        
            If Not Func_Win_Shell(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_EXIT
        
            Call cp.ensurele(2)
            '/a exit all.
            stt_exec = False
        '=====================================
        Case SLOT_CD, SLOT_CHANGEDIR
        
            If Not Func_CD(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_MOUNT
        
            If Not Func_Mount(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_UNMOUNT
        
            If Not Func_Unmount(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_PING
        
            'ping 192.168.0.1 [/t]
        '=====================================
        Case SLOT_ABOUT
        
            'about {"module name","plugin name","what"} /doc
            outMsg GetStringTable(STRINGTABLEID_ABOUT)
        '=====================================
        Case SLOT_HELP
        
            outMsg GetStringTable(STRINGTABLEID_HELP)
        '=====================================
        Case SLOT_GUID
            'guid {"module name","plugin name","what"}
        '=====================================
        Case SLOT_STARTUP
        
            'startup "License Key"
            '/nomount skip startup mounts
            '/nowrapper
        '=====================================
        Case SLOT_LICENSE
        
            outMsg GetStringTable(STRINGTABLEID_LICENSE)
        '=====================================
        Case SLOT_PLGINSTL, SLOT_INSTALLPLUGIN
        
            
        '=====================================
        Case SLOT_PLGUINSTL, SLOT_UNINSTALLPLUGIN
        
            
        '=====================================
        Case SLOT_VB6INSTALL
        
            
        '=====================================
        Case SLOT_SET_POWERSTATE
        
            
        '=====================================
        Case SLOT_CANCEL_POWERSTATE
        
            
        '=====================================
        Case SLOT_DEL
        
            
        '=====================================
        Case SLOT_CDIR
        
            
        '=====================================
        Case SLOT_MKDIR
        
            
        '=====================================
        Case SLOT_RMDIR
        
            
        '=====================================
        Case SLOT_FILE_DEFINE
        
            If Not cFile(cp).funcDefine(stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_FILE_OPEN
        
            If Not cFile(cp).funcOpen(stt_exec) Then GoTo dropCommand
        '=====================================
        'Case "closefile"
        
        '    If Not cFile(cp).funcClose( stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_SOCKET_DEFINE
        
            If Not cSocket(cp).funcSocket(stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_SOCKET_CONNECT
        
            If Not cSocket(cp).funcConnect(stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_SOCKET_SHUTDOWN
        
            If Not cSocket(cp).funcShutdown(stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_SOCKET_SEND
        
            If Not cSocket(cp).funcSend(stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_SOCKET_RECIEVE
        
            If Not cSocket(cp).funcRecieve(stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_SOCKET_LISTEN
        
            If Not cSocket(cp).funcListen(stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_SOCKET_BIND
        
            If Not cSocket(cp).funcBind(stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_SOCKET_DISCONNECT
        
            If Not cSocket(cp).funcDisconnect(stt_exec) Then GoTo dropCommand
        '=====================================
        'Case "closesock"
        
        '    If Not Func_CloseSocket(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case "fuck_u"
        
            outMsg "kesafat fohsh midi? madare to &%@$*&"
        '=====================================
        Case "how_r_u"
            
            outMsg "afarin hala shod!"
        '=====================================
        Case "yngwie_malmsteen"
        
            outMsg "WOW!"
        '=====================================
        Case "ronnie_james_dio"
            
            outMsg "GOD!"
        '=====================================
        Case "heavy_metal"
            
            outMsg "gonna kick your ass!"
        '=====================================
        Case SLOT_OPEN
            
            If Not Func_Open(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_CLOSE
            
            If Not Func_Close(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_SHOWVARS
            
            If Not Func_ShowVars(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_SCRIPT
            
            If Not Func_Script(cp, stt_exec) Then GoTo dropCommand
            'script about /doc ; help /doc ;
        '=====================================
        Case SLOT_GETSTRING
            
            If Not Func_GetString(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case SLOT_LIST
            
            If Not Func_List(cp, stt_exec) Then GoTo dropCommand
        '=====================================
        Case "developer"
            
            If Not sec_signin Then
                If cp.CountArguments = 3 Then
                    If cp.GetArg(0) = PA_WORD Then
                        sec_signin = True
                    End If
                End If
            End If
            env_mode = aeDeveloper
        '=====================================
        Case "shell"
            
            env_mode = aeShell
        '=====================================
        Case Else
            
            ExecuteCommand = False
            Exit Function
        '=====================================
    End Select
dropCommand:
    Continue = stt_exec
    ExecuteCommand = True
    Exit Function
errCatch:
    errMsg "An error occured:" & err.Description
    ExecuteCommand = True
End Function

Public Sub ShowHelpOf(CommandName As String)
    Select Case CommandName
        Case "about"
            
        Case "help"
            
        Case "exit"
            
        Case Else
            errMsg Replace(NRS, NRSS, CommandName)
    End Select
End Sub

Private Sub DefineVariable(ByVal Name As String)
    Name = Trim(Name)
    If Name = "" Then GoTo throwInvalidVariableName
    If VBA.IsNumeric(Left(Name, 1)) Then GoTo throwInvalidVariableName
    If InStr(1, Name, """") > 0 Then GoTo throwInvalidVariableName
    
    Call stack.AppendArgument(0, Name)
    Exit Sub
throwInvalidVariableName:
    throw Exceptions.Exception("Invalid variable name '" & Name & "'")
End Sub

Private Function Func_Exp(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not stack.Exists(arg0) Then throw Exceptions.ItemNotExistsException
    
    dDouble = 0
    
    For i = 1 To cp.CountArguments - 1
        arg1 = outQuotPath(cp.GetArg(i))
        dDouble = dDouble + ValNum(arg1)
    Next
    
    stack(arg0) = ValNum(stack(arg0)) + IIf(cp.CountArguments = 1, 1, dDouble)
    Func_Exp = True
End Function
Private Function Func_Add(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not stack.Exists(arg0) Then throw Exceptions.ItemNotExistsException
    
    dDouble = 0
    
    For i = 1 To cp.CountArguments - 1
        arg1 = outQuotPath(cp.GetArg(i))
        dDouble = dDouble + ValNum(arg1)
    Next
    
    stack(arg0) = ValNum(stack(arg0)) + IIf(cp.CountArguments = 1, 1, dDouble)
    Func_Add = True
End Function
Private Function Func_Sub(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not stack.Exists(arg0) Then throw Exceptions.ItemNotExistsException
    
    dDouble = 0
    
    For i = 1 To cp.CountArguments - 1
        arg1 = outQuotPath(cp.GetArg(i))
        dDouble = dDouble - ValNum(arg1)
    Next
    
    stack(arg0) = ValNum(stack(arg0)) + IIf(cp.CountArguments = 1, -1, dDouble)
    Func_Sub = True
End Function
Private Function Func_Mul(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not stack.Exists(arg0) Then throw Exceptions.ItemNotExistsException
    
    dDouble = 0
    
    For i = 1 To cp.CountArguments - 1
        arg1 = outQuotPath(cp.GetArg(i))
        dDouble = dDouble + ValNum(arg1)
    Next
    
    stack(arg0) = ValNum(stack(arg0)) * IIf(cp.CountArguments = 1, 2, dDouble)
    Func_Mul = True
End Function
Private Function Func_Div(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not stack.Exists(arg0) Then throw Exceptions.ItemNotExistsException
    
    dDouble = 0
    
    For i = 1 To cp.CountArguments - 1
        arg1 = outQuotPath(cp.GetArg(i))
        dDouble = dDouble + ValNum(arg1)
    Next
    
    stack(arg0) = ValNum(stack(arg0)) / IIf(cp.CountArguments = 1, 2, dDouble)
    Func_Div = True
End Function
Private Function Func_Pow(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not stack.Exists(arg0) Then throw Exceptions.ItemNotExistsException
    
    dDouble = 0
    
    For i = 1 To cp.CountArguments - 1
        arg1 = outQuotPath(cp.GetArg(i))
        dDouble = dDouble + ValNum(arg1)
    Next
    
    stack(arg0) = ValNum(stack(arg0)) ^ IIf(cp.CountArguments = 1, 2, dDouble)
    Func_Pow = True
End Function
Private Function Func_Sqr(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not CheckVar(arg0) Then Exit Function
    
    dDouble = 0
    
    For i = 1 To cp.CountArguments - 1
        arg1 = outQuotPath(cp.GetArg(i))
        dDouble = dDouble + ValNum(arg1)
    Next
    
    stack(arg0) = ValNum(stack(arg0)) ^ (1 / IIf(cp.CountArguments = 1, 2, dDouble))
    Func_Sqr = True
End Function
Private Function Func_Not(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensure(1)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not stack.Exists(arg0) Then throw Exceptions.ItemNotExistsException
    
    stack(arg0) = Not ValNum(stack(arg0))
    Func_Not = True
End Function
Private Function Func_Or(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not stack.Exists(arg0) Then throw Exceptions.ItemNotExistsException
    
    dDouble = ValNum(stack(arg0))
    
    For i = 1 To cp.CountArguments - 1
        arg1 = outQuotPath(cp.GetArg(i))
        dDouble = dDouble Or ValNum(arg1)
    Next
    
    stack(arg0) = dDouble
    Func_Or = True
End Function
Private Function Func_Xor(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not stack.Exists(arg0) Then throw Exceptions.ItemNotExistsException
    
    dDouble = ValNum(stack(arg0))
    
    For i = 1 To cp.CountArguments - 1
        arg1 = outQuotPath(cp.GetArg(i))
        dDouble = dDouble Xor ValNum(arg1)
    Next
    
    stack(arg0) = dDouble
    Func_Xor = True
End Function
Private Function Func_And(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not stack.Exists(arg0) Then throw Exceptions.ItemNotExistsException
    
    dDouble = ValNum(stack(arg0))
    
    For i = 1 To cp.CountArguments - 1
        arg1 = outQuotPath(cp.GetArg(i))
        dDouble = dDouble And ValNum(arg1)
    Next
    
    stack(arg0) = dDouble
    Func_And = True
End Function

Private Function Func_OUT(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(2)
    arg0 = outQuotPath(cp.GetArg(0))
    If stack.Exists(arg0) Then
        If stack.ArgumentType(arg0) <> vbObject Then _
            throw Exceptions.InvalidArgumentTypeException(Replace(IVTS, IVTSS, arg0))
        Set FE = stack(arg0)
        If Not TypeOf FE Is ITargetStream Then _
            throw Exceptions.InvalidArgumentTypeException(Replace(IVTS, IVTSS, arg0))
        Set ITS = FE
        For i = 1 To cp.CountArguments - 1
            If stack.Exists(cp.GetArg(i)) Then
                arg1 = outQuotPath(cp.GetArg(i))
                Call GetVar(arg1, FE)
                out ITS, FE
            Else
                out ITS, outQuotPath(cp.GetArg(i))
            End If
        Next
    Else
        errMsg Replace(ANES, ANESS, arg0)
    End If
    Func_OUT = True
End Function
Private Function Func_IN(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArg(0))
    If stack.Exists(arg0) Then
        If stack.ArgumentType(arg0) <> vbObject Then _
            throw Exceptions.InvalidArgumentTypeException(Replace(IVTS, IVTSS, arg0))
        Set FE = stack(arg0)
        If Not TypeOf FE Is ITargetStream Then _
            throw Exceptions.InvalidArgumentTypeException(Replace(IVTS, IVTSS, arg0))
        Set ITS = FE
        For i = 1 To cp.CountArguments - 1
            arg1 = outQuotPath(cp.GetArg(i))
            If Not stack.Exists(arg1) Then throw Exceptions.InvalidArgumentTypeException(Replace(ANES, ANESS, arg1))
            inp ITS, buf0 'buf0 must be string!
            stack(arg1) = buf0
        Next
    Else
        errMsg Replace(ANES, ANESS, arg0)
    End If
    Func_IN = True
End Function
Private Function Func_Set(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurele(2, 1)
    arg0 = outQuotPath(cp.GetArg(0))
    If cp.CountArguments = 2 Then
        arg1 = cp.GetArg(1)
    Else
        arg1 = 0
    End If
    If Not stack.Exists(arg0) Then _
        Call stack.AppendArgument(0, arg0)
    If stack.Exists(arg1) Then
        If stack.ArgumentType(arg1) = vbObject Then
            Set stack(arg0) = stack(arg1)
        Else
                stack(arg0) = stack(arg1)
        End If
    Else
        If (LCase(Left(arg1, 2)) = "0x") Or (LCase(Left(arg1, 1)) = "0") Then
            stack(arg0) = ValNum(arg1)
        Else
            stack(arg0) = outQuotPath(arg1)
        End If
    End If
    Func_Set = True
End Function
Private Function Func_Get(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensure(1)
    arg0 = outQuotPath(cp.GetArg(0))
    If stack.Exists(arg0) Then
        dlLen = stack.ArgumentType(arg0)
        If (dlLen = vbObject) Then
            outMsg arg0 & ":object(" & TypeName(stack(arg0)) & ")"
        ElseIf dlLen = vbArray Then
            FE = stack(arg0)
            buf0 = TypeName(FE)
            outMsg arg0 & ":" & Left(buf0, InStr(1, buf0, "(")) & "[" & ArraySize(FE) & "]"
        Else
            outMsg arg0 & ":" & stack(arg0)
        End If
    Else
        errMsg Replace(ANES, ANESS, arg0)
    End If
    Func_Get = True
End Function
Private Function Func_IsDef(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensure(1)
    arg0 = outQuotPath(cp.GetArg(0))
    If stack.Exists(arg0) Then
        vType = stack.ArgumentType(arg0)
        If vType = vbObject Or vType = vbArray Then
            outMsg arg0 & " exists containing:" & GetVarType(arg0)
        Else
            outMsg arg0 & " exists containing:" & stack(arg0)
        End If
    Else
        errMsg Replace(ANES, ANESS, arg0)
    End If
    Func_IsDef = True
End Function
Private Function Func_TypeOf(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensure(1)
    arg0 = outQuotPath(cp.GetArg(0))
    If stack.Exists(arg0) Then
        outMsg arg0 & " exists, type:" & GetVarType(arg0)
    Else
        errMsg Replace(ANES, ANESS, arg0)
    End If
    Func_TypeOf = True
End Function
Private Function Func_New(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensure(2)
    arg0 = outQuotPath(cp.GetArg(0))
    arg1 = outQuotPath(cp.GetArg(1))
    If stack.Exists(arg0) Then
        Dim objBuf As Object
        Set objBuf = CreateObject("arg1")
        'call InitializeCP.
        stack(arg0) = objBuf
    Else
        errMsg Replace(ANES, ANESS, arg0)
    End If
    Func_New = True
End Function
Private Function Func_Delete(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    
    If cp.CountArguments = 0 Then
        Call stack.Clear
        Call AppendDefaultVariables
        Func_Delete = True
        Exit Function
    End If
    
    arg0 = outQuotPath(cp.GetArg(0))
    arg1 = outQuotPath(cp.GetArg(1))
    If stack.Exists(arg0) Then
        stack(arg0) = endl
    Else
        errMsg Replace(ANES, ANESS, arg0)
    End If
    Func_Delete = True
End Function

Private Function Func_Echo(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1)
    For i = 0 To cp.CountArguments - 1
        arg1 = outQuotPath(cp.GetArg(i))
        Call GetVar(arg1, FE)
        out Console, FE
    Next
    Func_Echo = True
End Function
Private Function Func_Win_Shell(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Dim shellProccessHandle As Long
    Call cp.ensurege(1)
    arg0 = outQuotPath(cp.GetArgString)
    Call GetVar(arg0, FE)
    arg0 = CStr(FE)
    shellProccessHandle = Shell(arg0)
    'outMsg API_WaitForSingleObject(shellProccessHandle, tINFINITE)
    Func_Win_Shell = True
End Function
Private Function Func_CD(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurele(1)
    If cp.CountArguments = 0 Then
        Call outMsg(CurrentDirectory.AbsolutePath, False)
    Else
        arg0 = outQuotPath(cp.GetArg(0))
        Call GetVar(arg0, FE)
        arg0 = CStr(FE)
        Call ChDir(arg0)
    End If
    Func_CD = True
End Function
Private Function Func_Exec(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    
    Func_Exec = True
End Function

Private Function Func_Mount(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(2, 3)
    arg0 = outQuotPath(cp.GetArg(0))
    arg1 = outQuotPath(cp.GetArg(1))
    
    Call GetVar(arg0, FE)
    arg0 = CStr(FE)
    Call GetVar(arg1, FE)
    arg1 = CStr(FE)
    
    buf0 = ""
    If cp.CountArguments = 3 Then buf0 = cp.GetArg(2)
    dlLen = Len(arg0)
    If dlLen = 3 Then
        arg0 = Replace(arg0, "\", "/")
        arg0 = Replace(arg0, "|", "/")
        If Not arg0 Like "?:/" Then throw Exceptions.InvalidPathException("Invalid drive letter.")
    ElseIf dlLen = 2 Then
        If Not arg0 Like "?:" Then throw Exceptions.InvalidPathException("Invalid drive letter.")
    ElseIf dlLen = 1 Then
        If Not arg0 Like "?" Then throw Exceptions.InvalidPathException("Invalid drive letter.")
    Else
        throw Exceptions.InvalidPathException("Invalid drive letter.")
    End If
    chC = Asc(Left(arg0, 1))
    If (chC < 65) Or (chC > 90 And chC < 97) Or (chC > 122) Then
        throw Exceptions.InvalidPathException("Invalid drive letter.")
    End If
    If Not Directory.Exists(arg1) Then throw Exceptions.PathNotFoundException
    'Call DiskDrive.MountFolder(arg0, arg1)
    If buf0 = "/r" Or buf0 = "/remember" Then
        'Complete later.
    End If
    'mount A: "C:\Folder1" [{/r,/remember}]
    '/r register mount to MintAPI startup mount.
    Func_Mount = True
End Function
Private Function Func_Unmount(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(2, 3)
    arg0 = outQuotPath(cp.GetArg(0))
    arg1 = outQuotPath(cp.GetArg(1))
    
    Call GetVar(arg0, FE)
    arg0 = CStr(FE)
    Call GetVar(arg1, FE)
    arg1 = CStr(FE)
    
    buf0 = ""
    If cp.CountArguments = 3 Then buf0 = cp.GetArg(2)
    dlLen = Len(arg0)
    If dlLen = 3 Then
        arg0 = Replace(arg0, "\", "/")
        arg0 = Replace(arg0, "|", "/")
        If Not arg0 Like "?:/" Then throw Exceptions.InvalidPathException("Invalid drive letter.")
    ElseIf dlLen = 2 Then
        If Not arg0 Like "?:" Then throw Exceptions.InvalidPathException("Invalid drive letter.")
    ElseIf dlLen = 1 Then
        If Not arg0 Like "?" Then throw Exceptions.InvalidPathException("Invalid drive letter.")
    Else
        throw Exceptions.InvalidPathException("Invalid drive letter.")
    End If
    chC = Asc(Left(arg0, 1))
    If (chC < 65) Or (chC > 90 And chC < 97) Or (chC > 122) Then
        throw Exceptions.InvalidPathException("Invalid drive letter.")
    End If
    If Not Directory.Exists(arg1) Then throw Exceptions.PathNotFoundException
    'Call DiskDrive.UnmountFolder(arg0, arg1)
    If buf0 = "/t" Or buf0 = "/temp" Then
        'Complete later.
    End If
    'unmount A: "C:\Folder1" [{/t,/temp}]
    'default remove it from MintAPI startup mount but use /t for temporary unmount
    Func_Unmount = True
End Function
Private Function Func_ShowVars(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    For i = 0 To stack.Count - 1
        dlLen = stack.ArgumentType(i)
        outMsg StringParser(CStr(i)).PadLeft(4, "0") & " " & StringParser(stack.NameOf(i)).PadRight(20, " ") & " - " & GetVarStr(stack.NameOf(i))
'        If dlLen = vbArray Then
'            FE = stack(i)
'            buf0 = TypeName(FE)
'            outMsg StringParser(CStr(i)).PadLeft(4, "0") & " " & stack.NameOf(i) & " - " & Left(buf0, InStr(1, buf0, "(")) & "[" & ArraySize(FE) & "]"
'        ElseIf dlLen = vbObject Then
'            outMsg StringParser(CStr(i)).PadLeft(4, "0") & " " & stack.NameOf(i) & " - object(" & TypeName(stack(i)) & ")"
'        Else
'            outMsg StringParser(CStr(i)).PadLeft(4, "0") & " " & stack.NameOf(i) & " - " & CStr(stack(i))
'        End If
    Next
    Func_ShowVars = True
End Function
Private Function Func_Script(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Dim isFileContent As Boolean
    Call cp.ensurege(1)
    arg0 = cp.GetArg(0)
    If Directory(arg0).Exists Then
        'Load file contents.
        isFileContent = True
    Else
        buf0 = cp.GetArgString
        isFileContent = False
    End If
    
    Dim ContinueExecute As Boolean, Coms() As String
    Coms = Split(buf0, ";")
    Dim newCP As New CodeParser
    
    For i = 0 To ArraySize(Coms) - 1
        Call newCP.ParseCommand(Coms(i))
        ContinueExecute = True
        Call ExecuteCommand(newCP, ContinueExecute)
        If Not ContinueExecute Then Exit For
    Next
    If Not isFileContent Then
        Execute = ContinueExecute
    End If
    Func_Script = True
End Function
Private Function Func_GetString(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(2)
    arg0 = outQuotPath(cp.GetArg(0))
    
    For i = 1 To cp.CountArguments - 1
        buf0 = InputBox(Replace(arg0, "INDEX", i))
        arg1 = outQuotPath(cp.GetArg(i))
        stack(arg1) = buf0
    Next
    Func_GetString = True
End Function
Private Function Func_List(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(0)
    If cp.CountArguments > 0 Then
        arg0 = outQuotPath(cp.GetArg(0))
        If Not Directory(arg0).Exists Then
            arg0 = CurrentDirectory.AbsolutePath
        End If
    Else
        arg0 = CurrentDirectory.AbsolutePath
    End If
    
    Dim Attr As VbFileAttribute
    Dim StepByStep As Boolean, ShowAttrs As Boolean
    Attr = vbNormal
    
    For i = 0 To cp.CountArguments - 1
        buf0 = cp.GetArg(i)
        If buf0 = "/r" Or buf0 = "/readonly" Then Attr = Attr Or vbReadOnly
        If buf0 = "/h" Or buf0 = "/hidden" Then Attr = Attr Or vbHidden
        If buf0 = "/s" Or buf0 = "/system" Then Attr = Attr Or vbSystem
        If buf0 = "/a" Or buf0 = "/archive" Then Attr = Attr Or vbArchive
        If buf0 = "/alias" Then Attr = Attr Or vbAlias
        If buf0 = "/sbs" Then StepByStep = True
        If buf0 = "/attrs" Then ShowAttrs = True
    Next
    
    Dim buf100_ As String
    Dim Files() As String, Directories() As String, ScreenHeight As Long
    Dim d As Directory, DirectoriesCount As Long, LastStepCount As Long
    Set d = Directory(arg0)
    Files = d.FileNames(Attr, False)
    Directories = d.SubDirectories(Attr)
    
    ScreenHeight = Console.WindowSize.Height
    
    outMsg "Directory listing(" & arg0 & ")..."
    DirectoriesCount = ArraySize(Directories)
    For i = 0 To DirectoriesCount - 1
        arg1 = Directory.ConcatPath(arg0, Directories(i))
        If ShowAttrs Then
            dlLen = FileSystem.GetAttr(arg1)
            buf0 = GetAttributeStringByFileAttribute(dlLen)
        Else
            buf0 = "    "
        End If
        outMsg StringParser(CStr(i)).PadLeft(4, "0") & " " & _
            Format(FileSystem.FileDateTime(arg1), "DD/MM/YYYY HH:mm") & _
            "  DIR " & buf0 & "  " & Directories(i)
        If StepByStep Then
            Console.Title = ScreenHeight
            LastStepCount = LastStepCount + 1
            If LastStepCount >= ScreenHeight - 2 Then
                outMsg "Press enter to continue..."
                inMsg buf100_
                LastStepCount = 0
                ScreenHeight = Console.WindowSize.Height
            End If
        End If
    Next
    For i = 0 To ArraySize(Files) - 1
        arg1 = Directory.ConcatPath(arg0, Files(i))
        If ShowAttrs Then
            dlLen = FileSystem.GetAttr(arg1)
            buf0 = GetAttributeStringByFileAttribute(dlLen)
        Else
            buf0 = "    "
        End If
        outMsg StringParser(CStr(DirectoriesCount + i)).PadLeft(4, "0") & " " & _
            Format(FileSystem.FileDateTime(arg1), "DD/MM/YYYY HH:mm") & _
            "      " & buf0 & "  " & Files(i)
        If StepByStep Then
            Console.Title = ScreenHeight
            LastStepCount = LastStepCount + 1
            If LastStepCount >= ScreenHeight - 2 Then
                outMsg "Press enter to continue..."
                inMsg buf100_
                LastStepCount = 0
                ScreenHeight = Console.WindowSize.Height
            End If
        End If
    Next
    outMsg ""
    outMsg "   Total Files:        " & ArraySize(Files)
    outMsg "   Total Directories:  " & DirectoriesCount
    
    Func_List = True
End Function

Public Function GetAttributeStringByFileAttribute(FA As VbFileAttribute) As String
    If (FA And vbReadOnly) = vbReadOnly Then _
        GetAttributeStringByFileAttribute = _
        GetAttributeStringByFileAttribute & "R"
        
    If (FA And vbHidden) = vbHidden Then _
        GetAttributeStringByFileAttribute = _
        GetAttributeStringByFileAttribute & "H"
        
    If (FA And vbSystem) = vbSystem Then _
        GetAttributeStringByFileAttribute = _
        GetAttributeStringByFileAttribute & "S"
        
    If (FA And vbArchive) = vbArchive Then _
        GetAttributeStringByFileAttribute = "A"
End Function


Public Function Func_Open(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1, 2)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If cp.CountArguments > 1 Then
        arg1 = outQuotPath(cp.GetArg(1))
    Else
        arg1 = ""
    End If
    
    If Not CheckVar(arg0) Then Exit Function
    
    vType = stack.ArgumentType(arg0)
    
    If vType = vbObject Then
        Set objBuf = stack(arg0)
        If TypeOf objBuf Is ITargetStream Then
            Select Case LCase(arg1)
                Case "both", "sdboth", StreamDirection.sdBoth, ""
                    dlLen = StreamDirection.sdBoth
                Case "in", "sdin", "sdinstream", StreamDirection.sdInStream
                    dlLen = StreamDirection.sdInStream
                Case "out", "sdout", "sdoutstream", StreamDirection.sdOutStream
                    dlLen = StreamDirection.sdOutStream
            End Select
            Set ITS = objBuf
            Call ITS.OpenStream(dlLen)
        Else
            throw Exceptions.InvalidArgumentTypeException
        End If
    Else
        throw Exceptions.InvalidArgumentTypeException
    End If
    
    Func_Open = True
End Function

'var,[stream direction]
'sd : sdboth, sdinstream, sdoutstream, sdin, sdout, 1, 2, 3
Public Function Func_Close(cp As CodeParser, Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensurege(1, 2)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If cp.CountArguments > 1 Then
        arg1 = outQuotPath(cp.GetArg(1))
    Else
        arg1 = ""
    End If
    
    If Not stack.Exists(arg0) Then
        'CLOSE HANDLE.
        Exit Function
    End If
    
    vType = stack.ArgumentType(arg0)
    
    If vType = vbObject Then
        Set objBuf = stack(arg0)
        If TypeOf objBuf Is ITargetStream Then
            Select Case LCase(arg1)
                Case "both", "sdboth", StreamDirection.sdBoth, ""
                    dlLen = StreamDirection.sdBoth
                Case "in", "sdin", "sdinstream", StreamDirection.sdInStream
                    dlLen = StreamDirection.sdInStream
                Case "out", "sdout", "sdoutstream", StreamDirection.sdOutStream
                    dlLen = StreamDirection.sdOutStream
            End Select
            Set ITS = objBuf
            Call ITS.CloseStream(dlLen)
        Else
            throw Exceptions.InvalidArgumentTypeException
        End If
    Else
        'CLOSE HANDLE.
    End If
    
    Func_Close = True
End Function





Private Function IApplication_AppDomain() As String
    IApplication_AppDomain = "com.MintAPI.shell"
End Function
Private Function IApplication_ApplicationBinPath() As String
    
End Function
Private Sub IApplication_ApplicationConfigurationChanged(ByVal E As MintAPI.EventArgs)
    
End Sub
Private Function IApplication_ApplicationConfigurationPath() As String
    
End Function
Private Function IApplication_ApplicationConfigurationRegistryPath(Optional ByVal System As Boolean = True) As String
    
End Function
Private Function IApplication_ApplicationDataPath() As String
    IApplication_ApplicationDataPath = ""
End Function
Private Sub IApplication_ApplicationInputMethodChanged(ByVal E As MintAPI.EventArgs)
    
End Sub
Private Sub IApplication_ApplicationLanguageChanged(ByVal E As MintAPI.EventArgs)
    
End Sub
Private Function IApplication_ApplicationLogFile() As String
    IApplication_ApplicationLogFile = ""
End Function
Private Function IApplication_ApplicationName() As String
    IApplication_ApplicationName = ""
End Function
Private Sub IApplication_ApplicationNetworkStateChanged(ByVal E As MintAPI.EventArgs)
    
End Sub
Private Function IApplication_ApplicationPath() As String
    IApplication_ApplicationPath = App.Path
End Function
Private Function IApplication_ApplicationPluginsPath() As String
    IApplication_ApplicationPluginsPath = Directory.ConcatPath(App.Path, "plugins")
End Function
Private Function IApplication_ApplicationRegistryPath(Optional ByVal System As Boolean = True) As String
    
End Function
Private Function IApplication_ApplicationTempPath() As String
    IApplication_ApplicationTempPath = Directory.GetApplicationSpecifiedTempPath(True, True)
End Function
Private Function IApplication_Author() As String
    IApplication_Author = ""
End Function
Private Function IApplication_CompanyName() As String
    IApplication_CompanyName = ""
End Function
Private Sub IApplication_Debugger(ByVal E As MintAPI.ExceptionOccuredEventArgs)
    
End Sub
Private Sub IApplication_Error(ByVal E As MintAPI.ExceptionOccuredEventArgs)
    
End Sub
Private Function IApplication_ExecutionPath() As String
    IApplication_ExecutionPath = ""
End Function
Private Function IApplication_GUID() As String
    IApplication_GUID = ""
End Function
Private Function IApplication_HelpLink() As String
    IApplication_HelpLink = ""
End Function
Private Function IApplication_LegalCopyright() As String
    IApplication_LegalCopyright = ""
End Function
Private Sub IApplication_Message(ByVal E As MintAPI.ApplicationMessageEventArgs)
    
End Sub
Private Function IApplication_ProductCode() As String
    IApplication_ProductCode = ""
End Function
Private Function IApplication_ProductCode50chars() As String
    IApplication_ProductCode50chars = ""
End Function
Private Function IApplication_ProductName() As String
    IApplication_ProductName = ""
End Function
Private Sub IApplication_Restart(ByVal E As MintAPI.EventArgs)
    
End Sub
Private Function IApplication_ServiceWebsite() As String
    
End Function
Private Sub IApplication_Started(ByVal E As MintAPI.EventArgs)
    
End Sub
Private Sub IApplication_Starting(ByVal E As MintAPI.EventArgs)
    
End Sub
Private Function IApplication_SupportLink() As String
    
End Function
Private Sub IApplication_Terminated(ByVal E As MintAPI.EventArgs)
    
End Sub
Private Sub IApplication_Terminating(ByVal E As MintAPI.CancelEventArgs)
    
End Sub
Private Function IApplication_UpdateLink() As String
    IApplication_UpdateLink = ""
End Function
Private Function IApplication_VersionMajor() As Integer
    IApplication_VersionMajor = App.Major
End Function
Private Function IApplication_VersionMinor() As Long
    IApplication_VersionMinor = App.Minor
End Function
Private Function IApplication_VersionRevision() As Long
    IApplication_VersionRevision = App.Revision
End Function
Private Function IApplication_VersionString() As String
    IApplication_VersionString = ""
End Function
Private Function IApplication_VersionTags() As String
    IApplication_VersionTags = ""
End Function
Private Function IApplication_Website() As String
    IApplication_Website = ""
End Function




Private Property Get ISupportSlot_AutoManage() As Boolean
    ISupportSlot_AutoManage = False
End Property
Private Sub ISupportSlot_ManualManageCall(SlotName As String)
    Dim Continue As Boolean, cp As New CodeParser
    Continue = True
    Call ExecuteCommand(cp, Continue)
End Sub
Private Function ISupportSlot_Methods() As String
    ISupportSlot_Methods = _
        "0" & SLOT_IF & " 0" & SLOT_ENDIF & " 0" & SLOT_LOOP & _
        " 0" & SLOT_ENDLOOP & " 0" & SLOT_NEW & " 0" & SLOT_DELETE & _
        " 0" & SLOT_OPEN & " 0" & SLOT_CLOSE & _
        " 0" & SLOT_DEF & " 0" & SLOT_UDEF & " 0" & SLOT_SET & _
        " 0" & SLOT_GET & " 0" & SLOT_ISDEF & " 0" & SLOT_OUT & _
        " 0" & SLOT_EXEC & " 0" & SLOT_ECHO & " 0" & " 0" & SLOT_TYPEOF & _
        " 0" & SLOT_IN & " 0" & SLOT_CD & " 0" & SLOT_CHANGEDIR & _
        " 0" & SLOT_STARTUP & " 0" & SLOT_EXIT & " 0" & _
        " 0" & SLOT_SCRIPT & " 0" & SLOT_MKDIR & " 0" & SLOT_RMDIR & _
        " 0" & SLOT_CDIR & " 0" & SLOT_DEL & " 0" & SLOT_VB6INSTALL & _
        " 0" & SLOT_SET_POWERSTATE & " 0" & SLOT_CANCEL_POWERSTATE & _
        " 0" & SLOT_PLGINSTL & " 0" & SLOT_INSTALLPLUGIN & _
        " 0" & SLOT_PLGUINSTL & " 0" & SLOT_UNINSTALLPLUGIN & _
        " 0" & SLOT_ABOUT & " 0" & SLOT_HELP & " 0" & SLOT_GUID & _
        " 0" & SLOT_PING & " 0" & SLOT_LICENSE & " 0" & SLOT_WIN_SHELL & _
        " 0" & SLOT_MOUNT & " 0" & SLOT_UNMOUNT & " 0" & SLOT_SHOWVARS & _
        " 0" & SLOT_GETSTRING & " 0" & SLOT_LIST & " 0" & SLOT_FILE_DEFINE & _
        " 0" & SLOT_FILE_OPEN
End Function
