VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="Console" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IClassStream

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Console" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_CLEAR As String = "clear"
Const SIGNAL_ALLOCED As String = "alloced"
Const SIGNAL_WRITE As String = "write"
Const SIGNAL_READ As String = "read"
Const SIGNAL_TITLECHANGED As String = "titlechanged"
Const SIGNAL_FORECOLORCHANGED As String = "forecolorchanged"
Const SIGNAL_BACKCOLORCHANGED As String = "backcolorchanged"
Const SIGNAL_CURSORPOSITIONCHANGED As String = "cursorpositionchanged"
Const SIGNAL_CURSORSIZECHANGED As String = "cursorsizechanged"
Const SIGNAL_CURSORVISIBLECHANGED As String = "cursorvisiblechanged"
Const SIGNAL_BUFFERSIZECHANGED As String = "buffersizechanged"
Const SIGNAL_WINDOWSIZECHANGED As String = "windowsizechanged"
Const SIGNAL_WINDOWPOSITIONCHANGED As String = "windowpositionchanged"
Const SIGNAL_BEEP As String = "beep"
Const SIGNAL_RESETCOLORS As String = "resetcolors"
Const SIGNAL_INSERTMODECHANGED As String = "insertmodechanged"
Const SIGNAL_CONTROLC As String = "ctrl_c"
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

Private Const INVALID_HANDLE_VALUE = -1
Private Const STD_OUTPUT_HANDLE = -11
Private Const STD_ERROR_HANDLE = -12&
Private Const STD_INPUT_HANDLE = -10&
Private Const CREATE_NEW_CONSOLE = &H10
Private Const MAX_DEFAULTCHAR = 2
Private Const MAX_LEADBYTES = 12  '  5 ranges, 2 bytes ea., 0 term.

Private Const BACKSPACECHAR As Long = 8
Private Const DELETECHAR As Long = 8
Private Const ENTERCHAR As Long = 13
Private Const ESCAPECHAR As Long = 27

Private Const FOREGROUND_BLUE = &H1     '  text color contains blue.
Private Const FOREGROUND_GREEN = &H2     '  text color contains green.
Private Const FOREGROUND_RED = &H4     '  text color contains red.
Private Const FOREGROUND_INTENSITY = &H8     '  text color is intensified.
Private Const BACKGROUND_BLUE = &H10    '  background color contains blue.
Private Const BACKGROUND_GREEN = &H20    '  background color contains green.
Private Const BACKGROUND_RED = &H40    '  background color contains red.
Private Const BACKGROUND_INTENSITY = &H80    '  background color is intensified.

Private Const KEY_EVENT = &H1     '  Event contains key event record
Private Const MOUSE_EVENTC = &H2     '  Event contains mouse event record
Private Const WINDOW_BUFFER_SIZE_EVENT = &H4     '  Event contains window change event record
Private Const MENU_EVENT = &H8     '  Event contains menu event record
Private Const FOCUS_EVENT = &H10    '  event contains focus change

Private Const FROM_LEFT_1ST_BUTTON_PRESSED = &H1
Private Const RIGHTMOST_BUTTON_PRESSED = &H2
Private Const FROM_LEFT_2ND_BUTTON_PRESSED = &H4
Private Const FROM_LEFT_3RD_BUTTON_PRESSED = &H8
Private Const FROM_LEFT_4TH_BUTTON_PRESSED = &H10

Private Const CTRL_C_EVENT = 0
Private Const CTRL_BREAK_EVENT = 1
Private Const CTRL_CLOSE_EVENT = 2
'  3 is reserved!
'  4 is reserved!
Private Const CTRL_LOGOFF_EVENT = 5
Private Const CTRL_SHUTDOWN_EVENT = 6
' Input Mode flags:
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_WINDOW_INPUT = &H8
Private Const ENABLE_MOUSE_INPUT = &H10
' Output Mode flags:
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2
'  ControlKeyState flags
Private Const RIGHT_ALT_PRESSED = &H1     '  the right alt key is pressed.
Private Const LEFT_ALT_PRESSED = &H2     '  the left alt key is pressed.
Private Const RIGHT_CTRL_PRESSED = &H4     '  the right ctrl key is pressed.
Private Const LEFT_CTRL_PRESSED = &H8     '  the left ctrl key is pressed.
Private Const SHIFT_PRESSED = &H10    '  the shift key is pressed.
Private Const NUMLOCK_ON = &H20    '  the numlock light is on.
Private Const SCROLLLOCK_ON = &H40    '  the scrolllock light is on.
Private Const CAPSLOCK_ON = &H80    '  the capslock light is on.
Private Const ENHANCED_KEY = &H100   '  the key is enhanced.
'  EventFlags
Private Const MOUSE_MOVED = &H1
Private Const DOUBLE_CLICK = &H2

Private Type CPINFO
    MaxCharSize As Long                    '  max length (Byte) of a char
    DefaultChar(MAX_DEFAULTCHAR) As Byte   '  default character
    LeadByte(MAX_LEADBYTES) As Byte        '  lead byte ranges
End Type
Private Type COORD
    X As Integer
    Y As Integer
End Type
Private Type SMALL_RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type
Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD
    dwCursorPosition As COORD
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type
Private Type CONSOLE_CURSOR_INFO
    dwSize As Long
    bVisible As Long
End Type
Private Type CHAR_INFO
    Char As Integer
    Attributes As Integer
End Type

Private Type KEY_EVENT_RECORD
    bKeyDown As Long
    wRepeatCount As Integer
    wVirtualKeyCode As Integer
    wVirtualScanCode As Integer
    uChar As Byte
    dwControlKeyState As Long
End Type
Private Type MOUSE_EVENT_RECORD
    dwMousePosition As COORD
    dwButtonState As Long
    dwControlKeyState As Long
    dwEventFlags As Long
End Type
Private Type WINDOW_BUFFER_SIZE_RECORD
    dwSize As COORD
End Type
Private Type MENU_EVENT_RECORD
    dwCommandId As Long
End Type
Private Type FOCUS_EVENT_RECORD
    bSetFocus As Long
End Type
Private Type CONSOLE_INPUT_RECORD_EVENT
    KeyEvent As KEY_EVENT_RECORD
    MouseEvent As MOUSE_EVENT_RECORD
    WindowBufferSizeEvent As WINDOW_BUFFER_SIZE_RECORD
    MenuEvent As MENU_EVENT_RECORD
    FocusEvent As FOCUS_EVENT_RECORD
End Type
Private Type INPUT_RECORD
    EventType As Byte
    Event As CONSOLE_INPUT_RECORD_EVENT
End Type
Public Type Console_Input_Available_Record
    LPart As Byte
    UPart As Byte
End Type

Public Enum ConsoleColors
    ccDefault = -1
    ccBlack = 0
    ccDarkBlue = 1
    ccDarkGreen = 2
    ccDarkAqua = 3
    ccDarkRed = 4
    ccPurple = 5
    ccBrown = 6
    ccGrey = 7
    ccDarkWhite = ccGrey
    ccDarkGrey = 8
    ccBlue = 9
    ccGreen = 10
    ccAqua = 11
    ccRed = 12
    ccMagenta = 13
    ccPink = ccMagenta
    ccYellow = 14
    ccWhite = 15
End Enum
Public Enum ConsoleCaretSize
    ccsThin
    ccsNormal = ccsThin
    ccsThick
    ccsFull
End Enum

Const TITLELENGTH = 255


Private Declare Function API_AllocConsole Lib "kernel32" Alias "AllocConsole" () As Long
Private Declare Function API_FreeConsole Lib "kernel32" Alias "FreeConsole" () As Long

Private Declare Function API_SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function API_SetConsoleTitleUnicode Lib "kernel32" Alias "SetConsoleTitleW" (ByVal lpConsoleTitle As String) As Long
Private Declare Function API_GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
Private Declare Function API_GetConsoleTitleUnicode Lib "kernel32" Alias "GetConsoleTitleW" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
Private Declare Function API_SetConsoleCursorPosition Lib "kernel32" Alias "SetConsoleCursorPosition" (ByVal hConsoleOutput As Long, ByVal dwCursorPosition As Long) As Long
Private Declare Function API_GetConsoleScreenBufferInfo Lib "kernel32" Alias "GetConsoleScreenBufferInfo" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function API_SetConsoleTextAttribute Lib "kernel32" Alias "SetConsoleTextAttribute" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function API_FillConsoleOutputCharacter Lib "kernel32" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, ByVal dwWriteCoord As Long, lpNumberOfCharsWritten As Long) As Long
Private Declare Function API_FillConsoleOutputAttribute Lib "kernel32" Alias "FillConsoleOutputAttribute" (ByVal hConsoleOutput As Long, ByVal wAttribute As Long, ByVal nLength As Long, ByVal dwWriteCoord As Long, lpNumberOfAttrsWritten As Long) As Long
Private Declare Function API_SetConsoleScreenBufferSize Lib "kernel32" Alias "SetConsoleScreenBufferSize" (ByVal hConsoleOutput As Long, ByVal dwSize As Long) As Long
Private Declare Function API_SetConsoleCursorInfo Lib "kernel32" Alias "SetConsoleCursorInfo" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare Function API_GetConsoleCursorInfo Lib "kernel32" Alias "GetConsoleCursorInfo" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare Function API_SetConsoleWindowInfo Lib "kernel32" Alias "SetConsoleWindowInfo" (ByVal hConsoleOutput As Long, ByVal bAbsolute As Long, lpConsoleWindow As SMALL_RECT) As Long

Private Declare Function API_WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function API_WriteConsoleUnicode Lib "kernel32" Alias "WriteConsoleW" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long

Private Declare Function API_FlushConsoleInputBuffer Lib "kernel32" Alias "FlushConsoleInputBuffer" (ByVal hConsoleInput As Long) As Long
Private Declare Function API_ReadConsoleInput Lib "kernel32" Alias "ReadConsoleInputA" (ByVal hConsoleInput As Long, ByRef lpBuffer As INPUT_RECORD, ByVal nLength As Long, ByRef lpNumberOfEventsRead As Long) As Long

Private Declare Function API_GetStdHandle Lib "kernel32" Alias "GetStdHandle" (ByVal nStdHandle As Long) As Integer
Private Declare Function API_CloseConsoleHandle Lib "kernel32" Alias "CloseConsoleHandle" (Handle As Long) As Long

Private Declare Function API_SetConsoleMode Lib "kernel32" Alias "SetConsoleMode" (ByVal hConsoleHandle As Long, ByVal dwMode As Long) As Long
Private Declare Function API_SetConsoleCtrlHandler Lib "kernel32" Alias "SetConsoleCtrlHandler" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long

Private Declare Function API_GetLargestConsoleWindowSize Lib "kernel32" Alias "GetLargestConsoleWindowSize" (ByVal hConsoleOutput As Long) As COORD


Private out_hndl As Long
Private in_hndl As Long
Private err_hndl As Long

Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
'Public Function Create(ByVal Args As Long) As Object
'
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

Public Property Get CursorPosition() As Position
    Dim CSBINFO As CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) = 0 Then _
        throw Exps.SystemCallFailureException("Unable to get console screen buffer info.")

    CursorPosition.Left = CSBINFO.dwCursorPosition.X
    CursorPosition.Top = CSBINFO.dwCursorPosition.Y
End Property
Public Property Let CursorPosition(Value As Position)
    Dim cPos As Long, cTop As Long, cLeft As Long, tWidth As Long
    With GetBufferInfo.dwSize
        If Value.Left < 0 Or Value.Left >= .X Then _
            throw Exps.OutOfRangeException
        If Value.Top < 0 Or Value.Top >= .Y Then _
            throw Exps.OutOfRangeException
        tWidth = .X
    End With
    If Value.Left > &H7FFF Then throw Exps.OutOfRangeException
    If Value.Top > &H7FFF Then throw Exps.OutOfRangeException
    
    cLeft = IIf(Value.Left > 0, Value.Left, 0)
    cTop = ((IIf(Value.Top > 0, Value.Top, 0)) + (cLeft \ tWidth))
    cLeft = cLeft Mod tWidth
    cPos = (cTop * &H10000) + cLeft

    'Call API_SetConsoleCursorPosition(out_hndl, cPos)
    If API_SetConsoleCursorPosition(out_hndl, cPos) = 0 Then
        throw Exps.SystemCallFailureException("Unable to set console cursor position.")
    End If
    On Error GoTo 0
    'Call mySignalEmitter.Emit(SIGNAL_CURSORPOSITIONCHANGED)
End Property

Public Property Get CursorLeft() As Long
    CursorLeft = CursorPosition.Left
End Property
Public Property Let CursorLeft(Value As Long)
    CursorPosition = mint_constructor.Position(Value, CursorPosition.Top)
End Property
Public Property Get CursorTop() As Long
    CursorTop = CursorPosition.Top
End Property
Public Property Let CursorTop(Value As Long)
    CursorPosition = mint_constructor.Position(CursorPosition.Left, Value)
End Property

Public Property Get CursorVisible() As Boolean
    Dim CCI As CONSOLE_CURSOR_INFO

    'Call API_GetConsoleCursorInfo(out_hndl, CCI)
    If API_GetConsoleCursorInfo(out_hndl, CCI) = 0 Then _
        throw Exps.SystemCallFailureException("Unable to get cursor info.")

    CursorVisible = IIf(CCI.bVisible = 0, False, True)
End Property
Public Property Let CursorVisible(Value As Boolean)
    Dim CCI As CONSOLE_CURSOR_INFO
    CCI.bVisible = IIf(Value, 1, 0)
    CCI.dwSize = CursorSize

    'Call API_SetConsoleCursorInfo(out_hndl, CCI)
    If API_SetConsoleCursorInfo(out_hndl, CCI) = 0 Then _
        throw Exps.SystemCallFailureException("Unable to change cursor info.")
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CURSORVISIBLECHANGED)
End Property
Public Property Get CursorSize() As Long 'ConsoleCaretSize
    Dim CCI As CONSOLE_CURSOR_INFO

    'Call API_GetConsoleCursorInfo(out_hndl, CCI)
    If API_GetConsoleCursorInfo(out_hndl, CCI) = 0 Then _
        throw Exps.SystemCallFailureException("Unable to get cursor info.")

    CursorSize = CCI.dwSize
End Property
Public Property Let CursorSize(Value As Long) 'ConsoleCaretSize
    If Value < 0 Or Value > 100 Then throw Exps.OutOfRangeException("CursorSize allowed 0~100.")
    Dim CCI As CONSOLE_CURSOR_INFO
    CCI.dwSize = Value
    CCI.bVisible = CursorVisible

    'Call API_SetConsoleCursorInfo(out_hndl, CCI)
    If API_SetConsoleCursorInfo(out_hndl, CCI) = 0 Then _
        throw Exps.SystemCallFailureException("Unable to change cursor info.")
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CURSORSIZECHANGED)
End Property

Public Property Get BackColor() As ConsoleColors
    Dim CSBINFO As CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) = 0 Then _
        throw Exps.SystemCallFailureException("Unable to get console screen buffer info.")

    BackColor = CLng(CSBINFO.wAttributes And &HF0) \ &H10
End Property
Public Property Let BackColor(Value As ConsoleColors)
    Dim bkColor As Long
    bkColor = (Value * &H10) + (ForeColor And &HF)
    'Current_Console_backColor = Value
    Call API_SetConsoleTextAttribute(out_hndl, bkColor)
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_BACKCOLORCHANGED)
End Property
Public Property Get ForeColor() As ConsoleColors
    Dim CSBINFO As CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) = 0 Then _
        throw Exps.SystemCallFailureException("Unable to get console screen buffer info.")

    ForeColor = (CSBINFO.wAttributes And &HF)
End Property
Public Property Let ForeColor(Value As ConsoleColors)
    Dim frColor As Long
    frColor = (BackColor * &H10) + (Value And &HF)
    'Current_Console_foreColor = Value
    Call API_SetConsoleTextAttribute(out_hndl, frColor)
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_FORECOLORCHANGED)
End Property

Public Property Get GetLargestSize() As Size
    Dim Coords As COORD
    Coords = API_GetLargestConsoleWindowSize(out_hndl)
    If ((Coords.X = 0) Or (Coords.Y = 0)) Then throw Exps.SystemCallFailureException("Can't get largest window width.")
    GetLargestSize.Width = Coords.X
    GetLargestSize.Height = Coords.Y
End Property
Public Property Get GetLargestWidth() As Long
    Dim Coords As COORD
    Coords = API_GetLargestConsoleWindowSize(out_hndl)
    If Coords.X = 0 Then throw Exps.SystemCallFailureException("Can't get largest window width.")
    GetLargestWidth = Coords.X
End Property
Public Property Get GetLargestHeight() As Long
    Dim Coords As COORD
    Coords = API_GetLargestConsoleWindowSize(out_hndl)
    If Coords.Y = 0 Then throw Exps.SystemCallFailureException("Can't get largest window width.")
    GetLargestHeight = Coords.Y
End Property

Public Function ReadChar() As Byte
    Dim BytesTransferred As Long
    Dim Char As INPUT_RECORD
    Dim SUCCESS As Long

    Call API_FlushConsoleInputBuffer(in_hndl)
    Operation(OPERATION_CONSOLE_READSTATE).State = False
looper:
    SUCCESS = API_ReadConsoleInput(in_hndl, Char, 1, BytesTransferred)
    If SUCCESS = 0 Then throw Exps.SystemCallFailureException("Unable to read console input.")
    If Char.EventType = KEY_EVENT Then
        ReadChar = Char.Event.KeyEvent.uChar
        If (ReadChar <> 0) And _
            CBool(Char.Event.KeyEvent.bKeyDown) Then Exit Function
    End If
    Call API_Sleep(10)

If Not Operation(OPERATION_CONSOLE_READSTATE).State Then _
    GoTo looper
End Function
Public Function ReadKey() As Integer
    Dim BytesTransferred As Long
    Dim Char As INPUT_RECORD
    Dim SUCCESS As Long

    Call API_FlushConsoleInputBuffer(in_hndl)
    Operation(OPERATION_CONSOLE_READSTATE).State = False
looper:

    SUCCESS = API_ReadConsoleInput(in_hndl, Char, 1, BytesTransferred)
    If SUCCESS = 0 Then throw Exps.SystemCallFailureException("Unable to read console input.")
    If Char.EventType = KEY_EVENT Then
        ReadKey = Char.Event.KeyEvent.wVirtualKeyCode
        If (ReadKey <> 0) And _
            CBool(Char.Event.KeyEvent.bKeyDown) Then Exit Function
    End If
    Call API_Sleep(10)
If Not Operation(OPERATION_CONSOLE_READSTATE).State Then _
    GoTo looper
End Function

Public Sub FillLine(StartPosition As Position, ByVal Count As Long, ByVal Character As Byte, Optional ByVal BackColor As ConsoleColors = ccDefault, Optional ByVal ForeColor As ConsoleColors = ccDefault)
    Dim dwI As Long  ' Characters written in buffer, not used, but necessary output for later function
    Dim CAttr As Long
    Dim dwSize As Long ' Characters to clear ("area" of console, so to speak)
    Dim CSBINFO As CONSOLE_SCREEN_BUFFER_INFO
    Dim bfLeft As Long, bfTop As Long
    Dim pos As Long

    If Count <= 0 Then Exit Sub

    dwSize = Count

    If ForeColor = ccDefault Then CAttr = Me.ForeColor Else CAttr = ForeColor
    If BackColor = ccDefault Then CAttr = CAttr Or Me.BackColor Else CAttr = (BackColor * &H10) + CAttr

    bfLeft = StartPosition.Left
    bfTop = StartPosition.Top
    If bfLeft < 0 Then bfLeft = 0
    If bfTop < 0 Then bfTop = 0
    pos = (bfTop * &H10000) + bfLeft
    Dim MaxRect_Bounds As Long
    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then
        MaxRect_Bounds = (CLng(CSBINFO.dwSize.X) * CLng(CSBINFO.dwSize.Y))
        If (bfTop * bfLeft) >= MaxRect_Bounds Then Exit Sub
        If bfLeft + dwSize > CLng(CSBINFO.dwSize.X) Then dwSize = CLng(CSBINFO.dwSize.X) - bfLeft
        If Character <> -1 Then _
            Call API_FillConsoleOutputCharacter(out_hndl, Character Mod 255, dwSize, pos, dwI) ' Write char to buffer as many times as we tell it, i.e. dwSize
        Call API_FillConsoleOutputAttribute(out_hndl, CAttr, dwSize, pos, dwI)
        'CursorPosition = mint_constructor.Position(0, 0) ' Set console cursor to top left coord
    Else
        throw Exps.SystemCallFailureException("Can't fill console screen buffer.")
    End If
End Sub
Public Sub FillRect(StartPosition As Position, EndPosition As Position, ByVal Character As Long, Optional ByVal BackColor As ConsoleColors = ccDefault, Optional ByVal ForeColor As ConsoleColors = ccDefault)
    Dim dwI As Long  ' Characters written in buffer, not used, but necessary output for later function
    Dim CAttr As Long
    Dim dwSize As Long ' Characters to clear ("area" of console, so to speak)
    Dim CSBINFO As CONSOLE_SCREEN_BUFFER_INFO
    Dim bfLeft As Long, bfLinesCount As Long
    Dim pos As Long

    If ForeColor = ccDefault Then CAttr = Me.ForeColor Else CAttr = ForeColor
    If BackColor = ccDefault Then CAttr = CAttr Or Me.BackColor Else CAttr = (BackColor * &H10) + CAttr

    bfLeft = StartPosition.Left
    'bfTop = StartPosition.Top

    If bfLeft < 0 Then Exit Sub
    If StartPosition.Top < 0 Then Exit Sub
    dwSize = EndPosition.Left - StartPosition.Left + 1
    If dwSize <= 0 Then Exit Sub
    If EndPosition.Top < StartPosition.Top Then Exit Sub
    bfLinesCount = EndPosition.Top - StartPosition.Top + 1

    Dim i As Long
    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then
        If bfLeft + dwSize > CLng(CSBINFO.dwSize.X) Then dwSize = CLng(CSBINFO.dwSize.X) - bfLeft
        For i = 0 To bfLinesCount - 1
            pos = ((StartPosition.Top + i) * &H10000) + bfLeft
            If Character <> -1 Then _
                Call API_FillConsoleOutputCharacter(out_hndl, Character Mod 255, dwSize, pos, dwI) ' Write char to buffer as many times as we tell it, i.e. dwSize
            Call API_FillConsoleOutputAttribute(out_hndl, CAttr, dwSize, pos, dwI)
            'CursorPosition = mint_constructor.Position(0, 0) ' Set console cursor to top left coord
        Next
    Else
        throw Exps.SystemCallFailureException("Can't fill console screen buffer.")
    End If
End Sub

Public Sub Clear()
    Dim dwI As Long  ' Characters written in buffer, not used, but necessary output for later function
    Dim dwSize As Long ' Characters to clear ("area" of console, so to speak)
    Dim CSBINFO As CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then
        dwSize = CLng(CSBINFO.dwSize.X) * CLng(CSBINFO.dwSize.Y)
        Call API_FillConsoleOutputCharacter(out_hndl, 0, dwSize, 0, dwI) ' Write char to buffer as many times as we tell it, i.e. dwSize
        Call API_FillConsoleOutputAttribute(out_hndl, 0, dwSize, 0, dwI)
        Me.CursorPosition = mint_constructor.Position(0, 0) ' Set console cursor to top left coord
        
    Else
        throw Exps.SystemCallFailureException("Can't clear console screen buffer.")
    End If
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CLEAR)
End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private Helpers>
'
Friend Sub Dispose()
    
End Sub

Friend Property Get OutputHandle() As Long
    OutputHandle = out_hndl
End Property
Friend Property Get InputHandle() As Long
    InputHandle = in_hndl
End Property
Friend Property Get StdErrorHandle() As Long
    StdErrorHandle = err_hndl
End Property

Private Function GetBufferInfo() As CONSOLE_SCREEN_BUFFER_INFO
    If API_GetConsoleScreenBufferInfo(out_hndl, GetBufferInfo) = 0 Then _
        throw Exps.GetIONumberException(Err.LastDllError, sdBoth)
End Function

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IClassStream implementation>
'
Private Function IClassStream_GetState(ByVal Dir As StreamDirection) As Boolean
    Dim Val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        Val = inState
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        Val = Val And outState
    End If
    ITargetStream_getState = Val
End Function
Private Sub IClassStream_InStream(ByVal Buffer As ClassStreamBuffer, Optional ByVal Length As Long = -1&)
    If Not inState Then throw Exps.InvalidStatusException
    
End Sub
Private Sub IClassStream_OpenStream(Optional ByVal Direction As StreamDirection = 3&)
    
End Sub
Private Sub IClassStream_OutStream(ByVal Buffer As ClassStreamBuffer, Optional ByVal Length As Long = -1&)
    If Not outState Then throw Exps.InvalidStatusException
    
End Sub
Private Function IClassStream_ReadBytes(ByVal Length As Long, Optional ByVal StartAt As Long = 0&) As Byte()
    
End Function
Private Sub IClassStream_SetInState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    
End Sub
Private Sub IClassStream_SetOutState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    
End Sub
Private Function IClassStream_Available(ByVal Direction As StreamDirection) As Long
    
End Function
Private Sub IClassStream_CloseStream(Optional ByVal Direction As StreamDirection = 3&)
    
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
