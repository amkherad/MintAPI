VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodeParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type IndexedStr
    Index As Long
    str As String
End Type

Dim comName As String
Dim argItSelf As String
Dim args() As String
Dim ps() As IndexedStr
Dim fs() As IndexedStr


Public Sub ParseCommand(StrCommand As String)
    Dim fNCI As Long 'first non character index.
    Dim ln As Long, i As Long, chS As String * 1, chC As Byte
    Call Dispose
    ln = Len(StrCommand)
    For i = 1 To ln
        chS = Mid(StrCommand, i, 1)
        chC = Asc(chS)
        If ((chC < 65) Or ((chC > 90) And (chC < 97)) Or (chC > 122)) Then
            If chS <> "_" Then _
                Exit For
        End If
        fNCI = i
    Next
    If fNCI = 0 Then Exit Sub
    comName = Left(StrCommand, fNCI)
    Dim extractionLength As Long, extractStr As String
    extractionLength = ln - fNCI
    If extractionLength > 0 Then
        If Mid(StrCommand, fNCI + 1, 1) = " " Then
            extractStr = Right(StrCommand, extractionLength - 1)
        Else
            extractStr = Right(StrCommand, extractionLength)
        End If
        args = SeparateCommands(extractStr)
        ps = GetPaths(args)
        fs = GetFlags(args)
        argItSelf = extractStr
    End If
End Sub

Private Function SeparateCommands(str As String) As String()
    Dim retS() As String, retSCount As Long, retBuff As String
    Dim IsInStr As Boolean, chS As String * 1
    Dim i As Long, ln As Long, bkSlashCount As Long, cbkSlashCounter As Long
    ln = Len(str)
    For i = 1 To ln
        chS = Mid(str, i, 1)
        If chS = """" Then
            If i > 1 Then
                cbkSlashCounter = 0
                For cbkSlashCounter = i - 1 To 1 Step -1
                    If Mid(str, cbkSlashCounter, 1) = "\" Then
                        bkSlashCount = bkSlashCount + 1
                    Else
                        Exit For
                    End If
                Next
                If bkSlashCount Mod 2 <> 0 Then GoTo dropQuot
            End If
            IsInStr = Not IsInStr
dropQuot:
            retBuff = retBuff & chS
'        ElseIf chS = "\" Then
'            If i > 1 Then
'                cbkSlashCounter = 0
'                For cbkSlashCounter = i - 1 To 1 Step -1
'                    If Mid(str, cbkSlashCounter, 1) = "\" Then
'                        bkSlashCount = bkSlashCount + 1
'                    Else
'                        Exit For
'                    End If
'                Next
'                If bkSlashCount Mod 2 = 0 Then GoTo dropbkSlash
'            End If
'            retBuff = retBuff & chS
'            GoTo continue_bkSlash
'dropbkSlash:
'            chS = Mid(str, i + 1, 1)
'            If Not (chS = "\" Or chS = """") Then _
'                retBuff = retBuff & "\"
'continue_bkSlash:
        Else
            If Not IsInStr Then
                If chS = " " Then
                    If retBuff <> "" Then
                        ReDim Preserve retS(retSCount)
                        retS(retSCount) = retBuff
                        retSCount = retSCount + 1
                        retBuff = ""
                    End If
                Else
                    retBuff = retBuff & chS
                End If
            Else
                retBuff = retBuff & chS
            End If
        End If
    Next
    If retBuff <> "" Then
        ReDim Preserve retS(retSCount)
        retS(retSCount) = retBuff
        retSCount = retSCount + 1
    End If
    SeparateCommands = retS
End Function
Private Function GetPaths(Coms() As String) As IndexedStr()
    Dim retS() As IndexedStr, retSCount As Long
    Dim C As Long, i As Long, IsInStr As Boolean
    For C = 0 To ArraySize(Coms) - 1
        If Left(Coms(C), 1) = """" Then
            If Right(Coms(C), 1) = """" Then
                For i = 1 To Len(Coms(C))
                    If i > 1 Then _
                        If Mid(Coms(C), i - 1, 1) = "\" Then GoTo dropQuot
                    If Mid(Coms(C), i, 1) = """" Then IsInStr = Not IsInStr
dropQuot:
                Next
                If Not IsInStr Then
                    ReDim Preserve retS(retSCount)
                    retS(retSCount).str = Coms(C)
                    retS(retSCount).Index = C
                    retSCount = retSCount + 1
                End If
            End If
        End If
    Next
    GetPaths = retS
End Function
Private Function GetFlags(Coms() As String) As IndexedStr()
    'only args with / start.
    Dim retS() As IndexedStr, retSCount As Long
    Dim C As Long, i As Long, IsInStr As Boolean
    For C = 0 To ArraySize(Coms) - 1
        If Left(Coms(C), 1) = "/" Then
            For i = 1 To Len(Coms(C))
                If i > 1 Then _
                    If Mid(Coms(C), i - 1, 1) = "\" Then GoTo dropQuot
                If Mid(Coms(C), i, 1) = """" Then IsInStr = Not IsInStr
dropQuot:
            Next
            If Not IsInStr Then
                ReDim Preserve retS(retSCount)
                retS(retSCount).str = Coms(C)
                retS(retSCount).Index = C
                retSCount = retSCount + 1
            End If
        End If
    Next
    GetFlags = retS
End Function

Public Sub Dispose()
    comName = ""
    Erase args
    Erase ps
    Erase fs
End Sub

Public Property Get CountArguments() As Long
    CountArguments = ArraySize(args)
End Property

Public Sub ensure(HowMany As Long)
    If ArraySize(args) <> HowMany Then throw Exceptions.InvalidCallException("Invalid number of arguments passed.")
End Sub
Public Sub ensurel(HowMany As Long, Optional Less As Long = 0)
    Dim S As Long
    S = ArraySize(args)
    If S >= HowMany Then throw Exceptions.TooManyArgumentsException
    If S < Less Then throw Exceptions.NotEnoughArgumentsException
End Sub
Public Sub ensurele(HowMany As Long, Optional Less As Long = 0)
    Dim S As Long
    S = ArraySize(args)
    If S > HowMany Then throw Exceptions.TooManyArgumentsException
    If S < Less Then throw Exceptions.NotEnoughArgumentsException
End Sub
Public Sub ensureg(HowMany As Long, Optional High As Long = 100)
    Dim S As Long
    S = ArraySize(args)
    If S <= HowMany Then throw Exceptions.NotEnoughArgumentsException
    If S > High Then throw Exceptions.TooManyArgumentsException
End Sub
Public Sub ensurege(HowMany As Long, Optional High As Long = 100)
    Dim S As Long
    S = ArraySize(args)
    If S < HowMany Then throw Exceptions.NotEnoughArgumentsException
    If S > High Then throw Exceptions.TooManyArgumentsException
End Sub

Public Function CommandName() As String
    CommandName = comName
End Function

Public Function GetArgString() As String
    GetArgString = argItSelf
End Function

Public Function Arguments() As String()
    Arguments = args
End Function
Public Function GetArg(Index As Long) As String
    GetArg = args(Index)
End Function
Public Function Paths() As String()
    Dim retVal() As String, i As Long, C As Long
    On Error GoTo zeroLength
    C = UBound(ps) - LBound(ps) + 1
zeroLength:
    ReDim retVal(C - 1)
    For i = 0 To C - 1
        retVal(i) = ps(i).str
    Next
    Paths = retVal
End Function
Public Function Flags() As String()
    Dim retVal() As String, i As Long, C As Long
    On Error GoTo zeroLength
    C = UBound(ps) - LBound(ps) + 1
zeroLength:
    ReDim retVal(C - 1)
    For i = 0 To C - 1
        retVal(i) = fs(i).str
    Next
    Flags = retVal
End Function
