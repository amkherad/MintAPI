VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gateway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event DataArrived(ByRef Data As Variant, ByVal Synchronous As Boolean)
Public Event AsynchronousError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
Public Event LinkInitialised()
Public Event LinkTerminated()

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CoCreateGuid Lib "ole32" (ID As Any) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const PROCESS_VM_READ        As Long = &H10

Private Const PROP_COMMAND      As String = "Command"
Private Const PROP_PARAMS       As String = "Parameters"

Private Const CMD_DATA          As String = "Data"
Private Const CMD_ERROR         As String = "Error"


Private Type mtypStruct
    MemPtr  As Long
    Length  As Long
End Type

Private mcolStack       As Collection

Private mbytReply()     As Byte
Private mbytData()      As Byte

Private mudtStruct      As mtypStruct

Private mlnghWnd        As Long
Private mlngChildhWnd   As Long

Private mblnClient      As Boolean
Private mblnSynchronous As Boolean

Private mblnHandleAsyncErrors   As Boolean

Private mlngOldWndProc  As Long

Private mdtePing        As Date

Private Sub Class_Initialize()
    Set mcolStack = New Collection
    CreateWindow
End Sub

Private Sub Class_Terminate()
    Destroy
    Set mcolStack = Nothing
End Sub

Friend Property Let OldWndProc(ByVal Value As Long)
    mlngOldWndProc = Value
End Property

Friend Property Get OldWndProc() As Long
    OldWndProc = mlngOldWndProc
End Property

Public Property Get hWnd() As Long
    hWnd = mlnghWnd
End Property

Public Property Get Linked() As Boolean
    Linked = (mlngChildhWnd > 0)
End Property

Public Property Let HandleAsyncErrors(ByVal Value As Boolean)
    mblnHandleAsyncErrors = Value
End Property

Public Property Get HandleAsyncErrors() As Boolean
    HandleAsyncErrors = mblnHandleAsyncErrors
End Property

Public Property Get StackCount() As Long
Dim lngCount    As Long
    lngCount = mcolStack.Count
    If SafeArrayGetDim(mbytData) > 0 Then
        lngCount = lngCount + 1
    End If
    StackCount = lngCount
End Property

Friend Property Let ChildhWnd(ByVal Value As Long)
    mlngChildhWnd = Value
    RaiseEvent LinkInitialised
End Property

Public Sub StartLink(ByVal hWnd As Long)
Dim lngRet  As Long
    If Linked Then
        StopLink
    End If
    lngRet = SendMessage(hWnd, MSG_LINKED, mlnghWnd, 0)
    If lngRet = SUCCESS Then
        ChildhWnd = hWnd
        mblnClient = True
    Else
        Err.Raise vbObjectError, App.EXEName, "Gateway not found."
    End If
End Sub

Public Sub StopLink()
    If Linked Then
        PostMessage mlngChildhWnd, MSG_LINK_TERMINATED, 0, 0
        mlngChildhWnd = 0
        mblnClient = False
        RaiseEvent LinkTerminated
        ClearAllData
    End If
End Sub

Private Sub CreateWindow()
    If mlnghWnd = 0 Then
        mlnghWnd = CreateWindowEx(0, "STATIC", CreateUID, 0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0&)
        HookWindow Me
    End If
End Sub

Friend Sub Destroy()
    If mlnghWnd > 0 Then
        StopLink
        UnhookWindow Me
        Call DestroyWindow(mlnghWnd)
        mlnghWnd = 0
        ClearAllData
    End If
End Sub

Private Sub ClearAllData()
    Set mcolStack = New Collection
    Erase mbytReply
    Erase mbytData
End Sub

Private Function CreateUID() As String
Dim bytID(1 To 16)      As Byte
Dim intIndex            As Integer
Dim strUID              As String
    Call CoCreateGuid(bytID(1))
    For intIndex = 1 To 16
        If bytID(intIndex) < CByte(16) Then
            strUID = strUID & "0"
        End If
        strUID = strUID & Hex$(bytID(intIndex))
        Select Case intIndex
            Case 4, 6, 8, 10
                strUID = strUID & "-"
        End Select
    Next intIndex
    CreateUID = "{" & strUID & "}"
End Function

Friend Function NewMessage(ByVal plngMemPtr As Long, ByVal pblnSynchronous As Boolean) As Long
Dim strCommand  As String
Dim varParams   As Variant
Dim objPB       As PropertyBag
On Error GoTo ErrHandler
    mblnSynchronous = pblnSynchronous
    ExtractData ReadMemory(mlngChildhWnd, plngMemPtr), strCommand, varParams
    PostMessage mlngChildhWnd, MSG_CLEAR_DATA, 0, 0
    Select Case strCommand
        Case CMD_DATA
            RaiseEvent DataArrived(varParams, pblnSynchronous)
        Case CMD_ERROR
            Set objPB = New PropertyBag
            With objPB
                .Contents = varParams
                RaiseEvent AsynchronousError(.ReadProperty("Number"), .ReadProperty("Source"), .ReadProperty("Description"))
            End With
            Set objPB = Nothing
    End Select
    If SafeArrayGetDim(mbytReply) > 0 Then
        With mudtStruct
            .Length = UBound(mbytReply) + 1
            .MemPtr = VarPtr(mbytReply(0))
        End With
        NewMessage = VarPtr(mudtStruct)
    End If
    mblnSynchronous = False
    Exit Function
ErrHandler:
    mblnSynchronous = False
    ClearReply
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SendData(ByRef Data As Variant) As Variant
    SendData = SendFormattedData(CreateData(CMD_DATA, Data))
End Function

Private Function SendFormattedData(ByRef pbytData() As Byte) As Variant
Dim strCommand  As String
Dim varParams   As Variant
Dim lngRet      As Long
Dim objPB       As PropertyBag
On Error GoTo ErrHandler
    If Linked Then
        If StackCount = 0 Then
            mbytData = pbytData
            With mudtStruct
                .Length = UBound(mbytData) + 1
                .MemPtr = VarPtr(mbytData(0))
            End With
            lngRet = SendMessage(mlngChildhWnd, MSG_DATA, VarPtr(mudtStruct), 0)
            Erase mbytData
            If lngRet > 0 Then
                ExtractData ReadMemory(mlngChildhWnd, lngRet), strCommand, varParams
                PostMessage mlngChildhWnd, MSG_CLEAR_REPLY, 0, 0
                Select Case strCommand
                    Case CMD_DATA
                        SendFormattedData = varParams
                    Case CMD_ERROR
                        Set objPB = New PropertyBag
                        With objPB
                            .Contents = varParams
                            Err.Raise .ReadProperty("Number"), .ReadProperty("Source"), .ReadProperty("Description")
                        End With
                        Set objPB = Nothing
                End Select
            End If
        Else
            Err.Raise vbObjectError, App.EXEName, "Outstanding data in stack, operation not allowed."
        End If
    Else
        Err.Raise vbObjectError, App.EXEName, "Gateway is not linked."
    End If
    Exit Function
ErrHandler:
    Erase mbytData
    If Linked Then
        PostMessage mlngChildhWnd, MSG_CLEAR_REPLY, 0, 0
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub PostData(ByRef Data As Variant)
    PostFormattedData CreateData(CMD_DATA, Data)
End Sub

Private Sub PostFormattedData(ByRef pbytData() As Byte)
    If Linked Then
        mcolStack.Add pbytData
        PostMessage mlngChildhWnd, MSG_POSTED_DATA, 0, 0
    Else
        Err.Raise vbObjectError, App.EXEName, "Gateway is not linked."
    End If
End Sub

Friend Function PushStack() As Long
    mbytData = mcolStack.Item(1)
    mcolStack.Remove 1
    With mudtStruct
        .Length = UBound(mbytData) + 1
        .MemPtr = VarPtr(mbytData(0))
    End With
    PushStack = VarPtr(mudtStruct)
End Function

Friend Function PushChildStack() As Long
    PushChildStack = SendMessage(mlngChildhWnd, MSG_PUSH_STACK, 0, 0)
End Function

Private Function ReadMemory(ByVal plnghWnd As Long, ByVal plngMemPtr As Long) As Byte()
Dim lngProcessID        As Long
Dim lngProcesshWnd      As Long
Dim lngBytesRead        As Long
Dim bytData()           As Byte
Dim udtStruct           As mtypStruct
On Error GoTo ErrHandler
    GetWindowThreadProcessId plnghWnd, lngProcessID
    lngProcesshWnd = OpenProcess(PROCESS_VM_READ, False, lngProcessID)
    If lngProcesshWnd > 0 Then
        Call ReadProcessMemory(lngProcesshWnd, ByVal plngMemPtr, ByVal VarPtr(udtStruct), LenB(udtStruct), lngBytesRead)
        ReDim bytData(0 To udtStruct.Length - 1) As Byte
        Call ReadProcessMemory(lngProcesshWnd, ByVal udtStruct.MemPtr, ByVal VarPtr(bytData(0)), udtStruct.Length, lngBytesRead)
        Call CloseHandle(lngProcesshWnd)
        ReadMemory = bytData
    End If
    Exit Function
ErrHandler:
    If lngProcesshWnd > 0 Then
        CloseHandle (lngProcesshWnd)
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Friend Sub ClearReply()
    Erase mbytReply
End Sub

Friend Sub ClearData()
    Erase mbytData
End Sub

Public Sub Ping()
    PingChild False
End Sub

Private Sub PingChild(ByVal pblnIsReply As Boolean)
    PostMessage mlngChildhWnd, MSG_PING, pblnIsReply, 0
End Sub

Friend Sub HasBeenPinged(ByVal pblnRequiresReply As Boolean)
    mdtePing = Now
    If pblnRequiresReply Then
        PingChild True
    End If
End Sub

Public Property Let Reply(ByRef Value As Variant)
    mbytReply = CreateData(CMD_DATA, Value)
End Property

Public Sub RaiseError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
Dim bytError()  As Byte
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty "Number", Number
        .WriteProperty "Source", Source
        .WriteProperty "Description", Description
        bytError = CreateData(CMD_ERROR, .Contents)
    End With
    Set objPB = Nothing
    If mblnSynchronous Then
        mbytReply = bytError
    Else
        If mblnHandleAsyncErrors Then
            RaiseEvent AsynchronousError(Number, Source, Description)
        Else
            PostFormattedData bytError
        End If
    End If
End Sub

Private Function CreateData(ByVal pstrCommand As String, ByVal pvarData As Variant) As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PROP_COMMAND, pstrCommand
        .WriteProperty PROP_PARAMS, pvarData
        CreateData = .Contents
    End With
    Set objPB = Nothing
End Function

Private Sub ExtractData(ByRef pbytContents() As Byte, ByRef pstrCommand As String, ByRef pvarParams As Variant)
Dim objPB       As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = pbytContents
        pstrCommand = .ReadProperty(PROP_COMMAND)
        pvarParams = .ReadProperty(PROP_PARAMS)
    End With
    Set objPB = Nothing
End Sub
