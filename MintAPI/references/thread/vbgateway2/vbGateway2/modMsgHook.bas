Attribute VB_Name = "modMsgHook"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Const GWL_WNDPROC As Long = -4
Private Const WM_DESTROY As Long = &H2

Public Const SUCCESS As Long = 1

Private mcolItems       As Collection

Private mlnghWnd        As Long
Private mlngGatewayPtr  As Long

Public Sub HookWindow(ByRef pobjGateway As Gateway)
Dim lngOldWndProc       As Long
    lngOldWndProc = SetWindowLong(pobjGateway.hWnd, GWL_WNDPROC, AddressOf WndProc)
    If lngOldWndProc > 0 Then
        pobjGateway.OldWndProc = lngOldWndProc
        If mcolItems Is Nothing Then
            Set mcolItems = New Collection
        End If
        mcolItems.Add ObjPtr(pobjGateway), CreateKey(pobjGateway.hWnd)
    Else
        Err.Raise vbObjectError, App.EXEName, "Unable to hook window."
    End If
End Sub

Public Sub UnhookWindow(ByRef pobjGateway As Gateway)
    mcolItems.Remove CreateKey(pobjGateway.hWnd)
    Call SetWindowLong(pobjGateway.hWnd, GWL_WNDPROC, pobjGateway.OldWndProc)
    If mcolItems.Count = 0 Then
        Set mcolItems = Nothing
    End If
End Sub

Private Function CreateKey(ByVal plnghWnd As Long) As String
    CreateKey = CStr(plnghWnd) & "K"
End Function

Private Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lngRet      As Long
Dim objGateway  As Gateway
On Error GoTo ErrHandler
    Set objGateway = PtrObj(mcolItems.Item(CreateKey(hWnd)))
    Select Case uMsg
        Case MSG_DATA
            WndProc = objGateway.NewMessage(wParam, True)
        Case MSG_POSTED_DATA
            lngRet = objGateway.PushChildStack
            If lngRet > 0 Then
                objGateway.NewMessage lngRet, False
            End If
        Case MSG_PUSH_STACK
            WndProc = objGateway.PushStack
        Case MSG_CLEAR_REPLY
            objGateway.ClearReply
        Case MSG_CLEAR_DATA
            objGateway.ClearData
        Case MSG_LINK_TERMINATED
            objGateway.StopLink
        Case MSG_PING
            objGateway.HasBeenPinged (wParam = 0)
        Case MSG_LINKED
            objGateway.ChildhWnd = wParam
            WndProc = SUCCESS
        Case Else
            WndProc = CallWindowProc(objGateway.OldWndProc, hWnd, uMsg, wParam, lParam)
    End Select
    Set objGateway = Nothing
    Exit Function
ErrHandler:
    If Not (objGateway Is Nothing) Then
        Set objGateway = Nothing
    End If
End Function

Private Function PtrObj(ByVal Pointer As Long) As Object
Dim objObject   As Object
    CopyMemory objObject, Pointer, 4&
    Set PtrObj = objObject
    CopyMemory objObject, 0&, 4&
End Function

Public Property Get MSG_DATA() As Long
Static lngMsgID     As Long
    If lngMsgID = 0 Then
        lngMsgID = RegisterWindowMessage("MSG_DATA")
    End If
    MSG_DATA = lngMsgID
End Property

Public Property Get MSG_CLEAR_REPLY() As Long
Static lngMsgID     As Long
    If lngMsgID = 0 Then
        lngMsgID = RegisterWindowMessage("MSG_CLEAR_REPLY")
    End If
    MSG_CLEAR_REPLY = lngMsgID
End Property

Public Property Get MSG_CLEAR_DATA() As Long
Static lngMsgID     As Long
    If lngMsgID = 0 Then
        lngMsgID = RegisterWindowMessage("MSG_CLEAR_DATA")
    End If
    MSG_CLEAR_DATA = lngMsgID
End Property

Public Property Get MSG_LINK_TERMINATED() As Long
Static lngMsgID     As Long
    If lngMsgID = 0 Then
        lngMsgID = RegisterWindowMessage("MSG_LINK_TERMINATED")
    End If
    MSG_LINK_TERMINATED = lngMsgID
End Property

Public Property Get MSG_PING() As Long
Static lngMsgID     As Long
    If lngMsgID = 0 Then
        lngMsgID = RegisterWindowMessage("MSG_PING")
    End If
    MSG_PING = lngMsgID
End Property

Public Property Get MSG_LINKED() As Long
Static lngMsgID     As Long
    If lngMsgID = 0 Then
        lngMsgID = RegisterWindowMessage("MSG_LINKED")
    End If
    MSG_LINKED = lngMsgID
End Property

Public Property Get MSG_POSTED_DATA() As Long
Static lngMsgID     As Long
    If lngMsgID = 0 Then
        lngMsgID = RegisterWindowMessage("MSG_POSTED_DATA")
    End If
    MSG_POSTED_DATA = lngMsgID
End Property

Public Property Get MSG_PUSH_STACK() As Long
Static lngMsgID     As Long
    If lngMsgID = 0 Then
        lngMsgID = RegisterWindowMessage("MSG_PUSH_STACK")
    End If
    MSG_PUSH_STACK = lngMsgID
End Property

