Attribute VB_Name = "uiMethods"
'@PROJECT_LICENSE

Option Explicit

Private Const API_uiMethods_SW_HIDE = 0
Private Const API_uiMethods_SW_PARENTCLOSING = 1
Private Const API_uiMethods_SW_NORMAL = 1
Private Const API_uiMethods_SW_SHOWNORMAL = 1
Private Const API_uiMethods_SW_SCROLLCHILDREN = &H1
Private Const API_uiMethods_SW_OTHERZOOM = 2
Private Const API_uiMethods_SW_SHOWMINIMIZED = 2
Private Const API_uiMethods_SW_INVALIDATE = &H2
Private Const API_uiMethods_SW_PARENTOPENING = 3
Private Const API_uiMethods_SW_SHOWMAXIMIZED = 3
Private Const API_uiMethods_SW_MAXIMIZE = 3
Private Const API_uiMethods_SW_SHOWNOACTIVATE = 4
Private Const API_uiMethods_SW_OTHERUNZOOM = 4
Private Const API_uiMethods_SW_ERASE = &H4
Private Const API_uiMethods_SW_SHOW = 5
Private Const API_uiMethods_SW_MINIMIZE = 6
Private Const API_uiMethods_SW_SHOWMINNOACTIVE = 7
Private Const API_uiMethods_SW_SHOWNA = 8
Private Const API_uiMethods_SW_RESTORE = 9
Private Const API_uiMethods_SW_MAX = 10
Private Const API_uiMethods_SW_SHOWDEFAULT = 10

Private Const API_uiMethods_SB_THUMBPOSITION = 4

Private Const API_uiMethods_SBS_HORZ = &H0&
Private Const API_uiMethods_SBS_VERT = &H1&

Private Const API_uiMethods_WM_HSCROLL = &H114
Private Const API_uiMethods_WM_VSCROLL = &H115
    
Private Const API_uiMethods_SCROLLACTION_JUMP = 0
Private Const API_uiMethods_SCROLLACTION_RELATIVE = 1

Public Enum ScrollDirection
    sdHorizontal = API_uiMethods_SBS_HORZ
    sdVertical = API_uiMethods_SBS_VERT
End Enum
Public Enum ScrollActionMode
    samJump = API_uiMethods_SCROLLACTION_JUMP
    samRelative = API_uiMethods_SCROLLACTION_RELATIVE
End Enum

Private Declare Function API_uiMethods_ShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal Action As Long) As Long
Private Declare Function API_uiMethods_GetScrollPos Lib "user32" Alias "GetScrollPos" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Private Declare Function API_uiMethods_PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function API_uiMethods_SetScrollPos Lib "user32" Alias "SetScrollPos" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function API_uiMethods_FlashWindow Lib "user32" Alias "FlashWindow" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Private Declare Function API_uiMethods_SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function API_uiMethods_CreateRoundRectRgn Lib "gdi32" Alias "CreateRoundRectRgn" (ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal x3 As Integer, ByVal y3 As Integer) As Long
Private Declare Function API_uiMethods_DeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function API_uiMethods_SetWindowRgn Lib "user32" Alias "SetWindowRgn" (ByVal hWnd As Long, ByVal hRGN As Long, ByVal bRedraw As Boolean) As Long

Dim inited As Boolean

Public Sub Initialize()
    If inited Then Exit Sub
    Call baseConstants.Initialize
    Call Exceptions.Initialize
    'Call baseMethods.Initialize
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    Call baseConstants.Dispose(Force)
    inited = False
End Sub

Public Sub RoundWindowEdges(hWnd As Long, fWidth As Long, fHeight As Long, rX As Long, rY As Long)
    Dim hRGN As Long
    hRGN = API_uiMethods_CreateRoundRectRgn(0, 0, fWidth, fHeight, rX, rY)
    Call API_uiMethods_SetWindowRgn(hWnd, hRGN, True)
    Call API_uiMethods_DeleteObject(hRGN)
End Sub
Public Sub ShowWindow(hWnd As Long)
    If API_uiMethods_ShowWindow(hWnd, API_uiMethods_SW_SHOW) <> SUCCESS Then throw SystemCallFailureException
End Sub
Public Sub HideWindow(hWnd As Long)
    If API_uiMethods_ShowWindow(hWnd, API_uiMethods_SW_HIDE) <> SUCCESS Then throw SystemCallFailureException
End Sub
Public Sub FlashWindow(hWnd As Long)
    If API_uiMethods_FlashWindow(hWnd, 1) <> SUCCESS Then throw SystemCallFailureException
End Sub

Public Sub ScrollWindow(ByVal hWnd As Long, ByVal Direction As ScrollDirection, ByVal ActionMode As ScrollActionMode, ByVal Amount As Long)
    Dim Position As Long
    ' What direction are we going
    If Direction = API_uiMethods_SBS_HORZ Then
        ' What action are we taking (Jumping or Relative)
        If ActionMode = API_uiMethods_SCROLLACTION_RELATIVE Then
            Position = API_uiMethods_GetScrollPos(hWnd, API_uiMethods_SBS_HORZ) + Amount
        Else
            Position = Amount
        End If
        ' Make it so
        If (API_uiMethods_SetScrollPos(hWnd, API_uiMethods_SBS_HORZ, Position, True) <> -1) Then
            Call API_uiMethods_PostMessage(hWnd, API_uiMethods_WM_HSCROLL, API_uiMethods_SB_THUMBPOSITION + &H10000 * Position, 0)
        Else
            Call throw(SystemCallFailureException)
        End If
    Else
        ' What action are we taking (Jumping or Relative)
        If ActionMode = API_uiMethods_SCROLLACTION_RELATIVE Then
            Position = API_uiMethods_GetScrollPos(hWnd, API_uiMethods_SBS_VERT) + Amount
        Else
            Position = Amount
        End If
        ' Make it so
        If (API_uiMethods_SetScrollPos(hWnd, API_uiMethods_SBS_VERT, Position, True) <> -1) Then
            Call API_uiMethods_PostMessage(hWnd, API_uiMethods_WM_VSCROLL, API_uiMethods_SB_THUMBPOSITION + &H10000 * Position, 0)
        Else
            Call throw(SystemCallFailureException)
        End If
    End If
End Sub
