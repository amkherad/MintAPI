Attribute VB_Name = "FileAPI"
Option Explicit

Public Type IN_FILE_H
    ID As Long
    fOSHandle As fOSFILE
    Path As String
    Key As String
End Type

Dim fs() As IN_FILE_H
Dim fscount As Long

Private Function IndexOf(what) As Long
    Dim i As Long
    Dim vt As VbVarType
    vt = VarType(what)
    If vt = vbString Then 'vt = vbLong Or vt = vbInteger Or vt = vbByte Or vt = vbDecimal Or vt = vbSingle Or vt = vbDouble
        For i = 0 To fscount - 1
            If fs(i).Path = what Or fs(i).Key = what Then
                IndexOf = i
                Exit Function
            End If
        Next
    Else
        For i = 0 To fscount - 1
            If fs(i).ID = what Then
                IndexOf = i
                Exit Function
            End If
        Next
    End If
    IndexOf = -1
End Function
Private Function GetFreeID() As Long
    Dim X As Long, Y As Long
    X = 1001
    While True
        For Y = 0 To fscount - 1
            If fs(Y).ID = X Then
                X = X + 1
                GoTo continue
            End If
        Next
        GoTo break
continue:
    Wend
break:
    GetFreeID = X
End Function
Public Function OpenFile_Internal(Path As String, Optional Key As String) As Long
    Dim fOSHndl As fOSFILE, IsOpened As Boolean
    On Error GoTo ErrHandler
    fOSHndl = baseFiling.OpenFile(Path)
    IsOpened = (fOSHndl.fHandle <> 0)
    
    ReDim Preserve fs(fscount)
    fs(fscount).ID = GetFreeID
    fs(fscount).Path = Path
    fs(fscount).Key = Key
    fs(fscount).fOSHandle = fOSHndl
    
    OpenFile_Internal = fs(fscount).ID
    fscount = fscount + 1
    Exit Function
ErrHandler:
    If IsOpened Then _
        Call baseFiling.CloseFile(fOSHndl)

    OpenFile_Internal = 0
    throw OpenFileException(err.Description)
End Function
Private Sub CloseFile_Internal_byIndex(Index As Long)
    Call baseFiling.CloseFile(fs(Index).fOSHandle)
    
    If fscount <= 1 Then
        Erase fs
        fscount = 0
    Else
        Dim i As Long
        For i = Index To fscount - 2
            fs(i) = fs(i + 1)
        Next
        fscount = fscount - 1
        ReDim Preserve fs(fscount)
    End If
End Sub
Public Sub CloseFile_Internal_byPath(Path As String)
    Dim i As Long, Index As Long
    For i = 0 To fscount - 1
        If fs(i).Path = Path Or fs(i).Key = Path Then
            Index = i
            GoTo continue
        End If
    Next
    throw ItemNotExistsException
continue:
    Call CloseFile_Internal_byIndex(Index)
End Sub
Public Sub CloseFile_Internal(ID As Long)
    Dim i As Long, Index As Long
    For i = 0 To fscount - 1
        If fs(i).ID = ID Then
            Index = i
            GoTo continue
        End If
    Next
    throw ItemNotExistsException
continue:
    Call CloseFile_Internal_byIndex(Index)
End Sub
