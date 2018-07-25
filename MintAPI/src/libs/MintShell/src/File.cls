VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim cp As CodeParser

Friend Sub Initialize(cp1 As CodeParser)
    Set cp = cp1
End Sub

'fileVar,[path]
Public Function funcDefine(Execute As Boolean)
    If Not Execute Then Exit Function
    Dim arg0 As String, arg1 As String
    Call cp.ensurege(1, 2)
    arg0 = outQuotPath(cp.GetArg(0))
    
    Dim fl As File
    
    If cp.CountArguments > 1 Then
        arg1 = outQuotPath(cp.GetArg(1))
        Set fl = File(arg1)
    Else
        arg1 = ""
        Set fl = File()
    End If
    
    If Not CheckVar(arg0) Then Exit Function
        
    Set stack(arg0) = fl
    
    funcDefine = True
End Function
'fileVar,[OpenMode],[Access],[FileShareMode],[FileCreateAttributes],[path]
Public Function funcOpen(Execute As Boolean)
    If Not Execute Then Exit Function
    Dim arg0 As String, arg1 As String, objBuf As Object, dlLen As Long
    Call cp.ensurege(1, 6)
    arg0 = outQuotPath(cp.GetArg(0))
    
    If Not CheckVar(arg0) Then Exit Function
    If stack.ArgumentType(arg0) <> vbObject Then
        errMsg Replace(IVTS, IVTSS, arg0)
        Exit Function
    End If
    
    
    Dim fl As File, FOM As FileMode, FSH As FileShare
    Dim FAT As FileAttributes, FAC As FileAccess
    
    FOM = fmCreate
    FAT = fNormal
    FSH = fshNone
    FAC = faReadWrite
    
    Set objBuf = stack(arg0)
    If TypeOf objBuf Is File Then
        Set fl = objBuf
    Else
        errMsg Replace(IVTS, IVTSS, arg0) & " " & GetVarType(arg0)
        Exit Function
    End If
    
    dlLen = cp.CountArguments
    
    If dlLen > 1 Then
        arg1 = outQuotPath(cp.GetArg(1))
        Select Case LCase(arg1)
            Case "a", "append", "w+", FileMode.fmAppend
                FOM = fmAppend
            Case "c", "create", FileMode.fmCreate
                FOM = fmCreate
            Case "cn", "createnew", FileMode.fmCreateNew
                FOM = fmCreateNew
            Case "o", "open", FileMode.fmOpen
                FOM = fmOpen
            Case "oc", "ooc", "openorcreate", FileMode.fmOpenOrCreate
                FOM = fmOpenOrCreate
            Case "t", "truncate", FileMode.fmTruncate
                FOM = fmTruncate
            Case Else
                FOM = fmOpen
        End Select
    End If
    If dlLen > 2 Then
        arg1 = outQuotPath(cp.GetArg(2))
        Select Case LCase(arg1)
            Case "a", "all", FileAccess.faAll
                FAC = faAll
            Case "r", "read", FileAccess.faRead
                FAC = faRead
            Case "w", "write", FileAccess.faWrite
                FAC = faWrite
            Case "wr", "writeread", "rw", "readwrite", FileAccess.faReadWrite
                FAC = faReadWrite
            Case "exec", "execute", FileAccess.faExecute
                FAC = faExecute
            Case Else
                FAC = faReadWrite
        End Select
    End If
    If dlLen > 3 Then
        arg1 = outQuotPath(cp.GetArg(2))
        Select Case LCase(arg1)
            Case "n", "none", FileShare.fshNone
                FSH = fshNone
            Case "r", "read", FileShare.fshRead
                FSH = fshRead
            Case "w", "write", FileShare.fshWrite
                FSH = fshWrite
            Case "wr", "writeread", "rw", "readwrite", FileShare.fshReadWrite
                FSH = fshReadWrite
            Case "del", "delete", FileShare.fshDelete
                FSH = fshDelete
            Case "inh", "inheritable", FileShare.fshInheritable
                FSH = fshInheritable
            Case Else
                FSH = fshNone
        End Select
    End If
    If dlLen > 4 Then
        arg1 = outQuotPath(cp.GetArg(2))
        FAT = ValNum(arg1)
    End If
    If dlLen > 5 Then
        arg1 = outQuotPath(cp.GetArg(5))
    Else
        arg1 = ""
    End If
    
    
    If arg1 = "" Then
        Call fl.Create(FOM, FAT, FAC, FSH)
    Else
        Call fl.Create(FOM, FAT, FAC, FSH, arg1)
    End If
    
    funcOpen = True
End Function
Public Function funcClose(Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensure(1)
    
    funcClose = True
End Function

