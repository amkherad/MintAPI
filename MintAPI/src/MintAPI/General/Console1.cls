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


'Ideas from Me , Qt , .net Framework , [VBCorLib streaming In_/Out methods]

Option Base 0
Option Explicit
Const CLASSID As String = "Console"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


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

Const SLOT_CLEAR As String = "clear"
Const SLOT_BEEP As String = "beep"
Const SLOT_RESETCOLORS As String = "resetcolors"

Const OPERATION_CONSOLE_READSTATE As String = "mint_console_readstate"

'Implements ITargetStream
Implements IClassStream
'Implements ISupportSignal
'Implements ISupportSlot
'Implements IAliasable

Const DEFAULT_BUFFER_X As Long = 80
Const DEFAULT_BUFFER_Y As Long = 1500


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


Dim out_hndl As Long
Dim in_hndl As Long
Dim err_hndl As Long

'Dim Current_Console_foreColor As ConsoleColors
'Dim Current_Console_backColor As ConsoleColors

Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String

Dim supportUIEventProcessing As Boolean
Dim mySignalEmitter As New SignalEmitter

Private Sub Class_Initialize()
    'Call mySignalEmitter.Initialize(Me)
    If API_AllocConsole Then
        out_hndl = API_GetStdHandle(STD_OUTPUT_HANDLE)
        in_hndl = API_GetStdHandle(STD_INPUT_HANDLE)
        err_hndl = API_GetStdHandle(STD_ERROR_HANDLE)
        If out_hndl = INVALID_HANDLE_VALUE Then throw Exps.InvalidOperationException("Unable to get stdoutput handle.")
        Call API_SetConsoleMode(out_hndl, ENABLE_PROCESSED_OUTPUT Or ENABLE_WRAP_AT_EOL_OUTPUT)
    Else
        out_hndl = API_GetStdHandle(STD_OUTPUT_HANDLE)
        in_hndl = API_GetStdHandle(STD_INPUT_HANDLE)
        err_hndl = API_GetStdHandle(STD_ERROR_HANDLE)
        If (API_SetConsoleMode(out_hndl, ENABLE_PROCESSED_OUTPUT Or ENABLE_WRAP_AT_EOL_OUTPUT) = 0) Or (out_hndl = INVALID_HANDLE_VALUE) Then _
            throw Exps.SystemCallFailureException("Unable to initialize process console.")
    End If

    'Call API_SetConsoleCtrlHandler(AddressOf modMain.modMain_Console_CtrlEvent_Handler, 0)

    Call API_SetConsoleScreenBufferSize(out_hndl, (DEFAULT_BUFFER_Y * &H10000) + DEFAULT_BUFFER_X)

    'Current_Console_backColor = ccBlack
    'Current_Console_foreColor = ccDarkWhite
End Sub
Private Sub Class_Terminate()
    Call Dispose
End Sub
Friend Sub Dispose()
    Call API_CloseConsoleHandle(out_hndl)
    Call API_CloseConsoleHandle(in_hndl)
    Call API_CloseConsoleHandle(err_hndl)

    Call API_FreeConsole
End Sub
Friend Sub Construct(Optional Arguments)

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

Public Sub Clear()
    Dim dwI As Long  ' Characters written in buffer, not used, but necessary output for later function
    Dim dwSize As Long ' Characters to clear ("area" of console, so to speak)
    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If Not API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then _
        throw Exps.IfError
    dwSize = CLng(CSBINFO.dwSize.X) * CLng(CSBINFO.dwSize.Y)
    Call API_FillConsoleOutputCharacter(out_hndl, 0, dwSize, 0, dwI) ' Write char to buffer as many times as we tell it, i.e. dwSize
    Call API_FillConsoleOutputAttribute(out_hndl, 0, dwSize, 0, dwI)
    CursorPosition = mint_constructor.Position(0, 0) ' Set console cursor to top left coord
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CLEAR)
End Sub

Public Sub ResetColors(Optional ByVal CharacterBackColor As ConsoleColors = ConsoleColors.ccBlack, Optional ByVal CharacterForeColor As ConsoleColors = ConsoleColors.ccDarkWhite)
    Dim chAttr As Long

    If CharacterBackColor = ccDefault Then CharacterBackColor = ccBlack
    If CharacterForeColor = ccDefault Then CharacterForeColor = ccGrey

    chAttr = (CharacterBackColor * &H10) + CharacterForeColor

    'Current_Console_backColor = CharacterBackColor
    'Current_Console_foreColor = CharacterForeColor

    Call API_SetConsoleTextAttribute(out_hndl, chAttr)
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_RESETCOLORS)
End Sub

Public Sub Beep(Optional ByVal Frequence As Long = 1000, Optional ByVal Duration As Long = 500, Optional ByVal Count As Integer = 1)
    Do While Count > 0
        Call API_Beep(Frequence, Duration)
        Count = Count - 1
        If Count > 0 Then Call API_Sleep(5)
    Loop
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_BEEP)
End Sub

Public Property Get CursorPosition() As Position
    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) = 0 Then _
        throw Exps.SystemCallFailureException("Unable to get console screen buffer info.")

    CursorPosition.Left = CSBINFO.dwCursorPosition.X
    CursorPosition.Top = CSBINFO.dwCursorPosition.Y
End Property
Public Property Let CursorPosition(Value As Position)
    If Value.Left > &H7FFF Then throw Exps.IndexOutOfRangeException
    If Value.Top > &H7FFF Then throw Exps.IndexOutOfRangeException
    Dim cPos As Long, cTop As Long, cLeft As Long, tWidth As Long
    tWidth = BufferSize.Width
    cLeft = IIf(Value.Left > 0, Value.Left, 0)
    cTop = ((IIf(Value.Top > 0, Value.Top, 0)) + (cLeft \ tWidth))
    cLeft = cLeft Mod tWidth
    cPos = (cTop * &H10000) + cLeft

    'Call API_SetConsoleCursorPosition(out_hndl, cPos)
    If Not API_SetConsoleCursorPosition(out_hndl, cPos) Then _
        throw Exps.IfError
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CURSORPOSITIONCHANGED)
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
    Dim CCI As API_CONSOLE_CURSOR_INFO

    If Not API_GetConsoleCursorInfo(out_hndl, CCI) Then _
        throw Exps.IfError

    CursorVisible = IIf(CCI.bVisible = 0, False, True)
End Property
Public Property Let CursorVisible(Value As Boolean)
    Dim CCI As API_CONSOLE_CURSOR_INFO
    CCI.bVisible = IIf(Value, 1, 0)
    CCI.dwSize = CursorSize

    If Not API_SetConsoleCursorInfo(out_hndl, CCI) Then _
        throw Exps.IfError
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CURSORVISIBLECHANGED)
End Property
Public Property Get CursorSize() As Long 'ConsoleCaretSize
    Dim CCI As API_CONSOLE_CURSOR_INFO

    'Call API_GetConsoleCursorInfo(out_hndl, CCI)
    If API_GetConsoleCursorInfo(out_hndl, CCI) = 0 Then _
        throw Exps.SystemCallFailureException("Unable to get cursor info.")

    CursorSize = CCI.dwSize
End Property
Public Property Let CursorSize(Value As Long) 'ConsoleCaretSize
    If Value < 0 Or Value > 100 Then throw Exps.IndexOutOfRangeException("CursorSize allowed 0~100.")
    Dim CCI As API_CONSOLE_CURSOR_INFO
    CCI.dwSize = Value
    CCI.bVisible = CursorVisible

    'Call API_SetConsoleCursorInfo(out_hndl, CCI)
    If Not API_SetConsoleCursorInfo(out_hndl, CCI) Then _
        throw Exps.IfError
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CURSORSIZECHANGED)
End Property


Public Property Get WindowPosition() As Position
    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If Not API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then _
        throw Exps.IfError

    WindowPosition.Left = CSBINFO.srWindow.Left
    WindowPosition.Top = CSBINFO.srWindow.Top
End Property
Public Property Let WindowPosition(Value As Position)
    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If Not API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then _
        throw Exps.IfError
    CSBINFO.srWindow.Top = Value.Top
    CSBINFO.srWindow.Left = Value.Left
    If Not API_SetConsoleWindowInfo(out_hndl, (Value.Top * &H10000) + Value.Left, CSBINFO.srWindow) Then _
        throw Exps.IfError
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_WINDOWPOSITIONCHANGED)
End Property
Public Property Get WindowSize() As Size
    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If Not API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then _
        throw Exps.IfError

    WindowSize.Width = CSBINFO.srWindow.Right
    WindowSize.Height = CSBINFO.srWindow.Bottom
End Property
Public Property Let WindowSize(Value As Size)

    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_WINDOWSIZECHANGED)
End Property

Public Property Get BufferSize() As Size
    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If Not API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then _
        throw Exps.IfError
    BufferSize.Width = CSBINFO.dwSize.X
    BufferSize.Height = CSBINFO.dwSize.Y
End Property
Public Property Let BufferSize(Value As Size)

    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_BUFFERSIZECHANGED)
End Property
Public Property Get LinearBufferSize() As Long
    Dim bSize As Size
    bSize = BufferSize
    LinearBufferSize = bSize.Width * bSize.Height
End Property

Public Property Get Title() As String
    If out_hndl = INVALID_HANDLE_VALUE Then _
        throw Exps.InvalidOperationException

    Dim strBuf As String
    strBuf = String(TITLELENGTH, Chr(0))

    'Call API_GetConsoleTitle(strBuf, TITLELENGTH)
    If API_GetConsoleTitle(strBuf, TITLELENGTH) <> 0 Then
        Title = strBuf
    End If
End Property
Public Property Let Title(Value As String)
    If out_hndl = INVALID_HANDLE_VALUE Then _
        throw Exps.InvalidOperationException

    'Call API_SetConsoleTitle(Value)
    If Not API_SetConsoleTitle(Value) Then _
        throw Exps.IfError
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_TITLECHANGED)
End Property

Public Property Get BackColor() As ConsoleColors
    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If Not API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then _
        throw Exps.IfError

    BackColor = ShiftRight(CLng(CSBINFO.wAttributes And &HF0), 4)
End Property
Public Property Let BackColor(Value As ConsoleColors)
    Dim bkColor As Long
    bkColor = (Value * &H10) + (ForeColor And &HF)
    'Current_Console_backColor = Value
    If Not API_SetConsoleTextAttribute(out_hndl, bkColor) Then _
        throw Exps.IfError
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_BACKCOLORCHANGED)
End Property
Public Property Get ForeColor() As ConsoleColors
    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If Not API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then _
        throw Exps.IfError

    ForeColor = (CSBINFO.wAttributes And &HF)
End Property
Public Property Let ForeColor(Value As ConsoleColors)
    Dim frColor As Long
    frColor = (BackColor * &H10) + (Value And &HF)
    'Current_Console_foreColor = Value
    If Not API_SetConsoleTextAttribute(out_hndl, frColor) Then _
        throw Exps.IfError
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_FORECOLORCHANGED)
End Property

Public Property Get Attributes() As Long
    Attributes = (ForeColor * &H10) + BackColor
End Property
Public Property Let Attributes(Value As Long)
    ForeColor = Value And &HF0
    BackColor = Value And &HF
End Property

Public Property Get GetLargestSize() As Size
    Dim Coords As API_COORD
    Coords = API_GetLargestConsoleWindowSize(out_hndl)
    If ((Coords.X = 0) Or (Coords.Y = 0)) Then throw Exps.IfError
    GetLargestSize.Width = Coords.X
    GetLargestSize.Height = Coords.Y
End Property
Public Property Get GetLargestWidth() As Long
    Dim Coords As API_COORD
    Coords = API_GetLargestConsoleWindowSize(out_hndl)
    If Coords.X = 0 Then throw Exps.IfError
    GetLargestWidth = Coords.X
End Property
Public Property Get GetLargestHeight() As Long
    Dim Coords As API_COORD
    Coords = API_GetLargestConsoleWindowSize(out_hndl)
    If Coords.Y = 0 Then throw Exps.IfError
    GetLargestHeight = Coords.Y
End Property

Public Sub SetUIEventProcessingState(Optional State As Boolean = True)
    supportUIEventProcessing = State
End Sub

Public Sub PrintF(Template As String, ParamArray Args())
    Dim cArgs() As Variant
    cArgs = Args
    Call WriteStr(funcwArgs(Template, cArgs))
End Sub
Public Sub ScanF(Template As String, ParamArray Args())
    Dim cArgs() As Variant
    cArgs = Args
    Call funcrArgs(Template, Me, cArgs)
End Sub
Public Sub WriteVar(Var)
    Dim strOut As String, numofCharsWritten As Long
    'Dim bPos As Position, cPos As Position
    'bPos = CursorPosition
    strOut = CStr(Var)
    'Call API_WriteConsole(out_hndl, ByVal strOut, Len(strOut), numofCharsWritten, 1)
    If Not API_WriteConsole(out_hndl, ByVal strOut, Len(strOut), numofCharsWritten, 1) Then _
        throw Exps.IfError
    'cPos = CursorPosition
    'Call FillRectAttribute(bPos, mint_constructor.Position(cPos.Left - 1, cPos.Top), -1, Attributes)
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_WRITE)
End Sub
Public Sub Cout(ParamArray Vars() As Variant)
    Dim strOut As String, numofCharsWritten As Long
    Dim i As Long, Count As Long
    On Error GoTo zeroLength
    Count = UBound(Vars) - LBound(Vars) + 1
zeroLength:
    For i = 0 To Count - 1
        Call WriteVar(Vars(i))
    Next
End Sub
Public Sub WriteStr(ByVal lineStr As String)
    Dim Strs_() As String
    Dim numofCharsWritten As Long
    Dim i As Long
    'Dim bPos As Position, cPos As Position
    'bPos = CursorPosition

    Strs_ = Split(lineStr, vbCrLf)
    lineStr = vbNullString
    For i = 0 To ArraySize(Strs_) - 1
        Strs_(i) = Replace(Strs_(i), vbCr, vbCrLf)
        Strs_(i) = Replace(Strs_(i), vbLf, vbCrLf)
        lineStr = lineStr & vbCrLf & Strs_(i)
    Next

    lineStr = Mid(lineStr, 3)

    'Call API_WriteConsole(out_hndl, ByVal lineStr, Len(lineStr), numofCharsWritten, 1)
    If Not API_WriteConsoleUnicode(out_hndl, ByVal lineStr, Len(lineStr), numofCharsWritten, 1) Then _
        throw Exps.IfError
    'cPos = CursorPosition
    'Call FillRectAttribute(bPos, mint_constructor.Position(cPos.Left - 1, cPos.Top), -1, Attributes)
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_WRITE)
End Sub
Public Sub WriteStr_SkipCrLf(ByVal lineStr As String)
    Dim Strs_() As String
    Dim numofCharsWritten As Long
    Dim i As Long
    'Dim bPos As Position, cPos As Position
    'bPos = CursorPosition

    Strs_ = Split(lineStr, vbCrLf)
    lineStr = ""
    For i = 0 To ArraySize(Strs_) - 1
        Strs_(i) = Replace(Strs_(i), vbCr, "")
        Strs_(i) = Replace(Strs_(i), vbLf, "")
        lineStr = lineStr & Strs_(i)
    Next

    'Call API_WriteConsole(out_hndl, ByVal lineStr, Len(lineStr), numofCharsWritten, 1)
    If Not API_WriteConsole(out_hndl, ByVal lineStr, Len(lineStr), numofCharsWritten, 1) Then _
        throw Exps.IfError
    'cPos = CursorPosition
    'Call FillRectAttribute(bPos, mint_constructor.Position(cPos.Left - 1, cPos.Top), -1, Attributes)
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_WRITE)
End Sub
Public Sub WriteLine(ByVal lineStr As String)
    Call WriteStr(lineStr & vbCrLf)
End Sub
Public Sub WriteByteArray(ByVal bArray As ByteArray)
    If bArray Is Nothing Then throw Exps.ArgumentNullException("bArray")
    If bArray.IsEmpty Then Exit Sub
    Dim numofCharsWritten As Long, bytesArray() As Byte
    'Dim bPos As Position, cPos As Position
    'bPos = CursorPosition

    bytesArray = bArray.ToConstData
    'Call API_WriteConsole(out_hndl, bytesArray(0), bArray.Length, numofCharsWritten, 1)
    If Not API_WriteConsole(out_hndl, bytesArray(0), bArray.Length, numofCharsWritten, 1) Then _
        throw Exps.IfError
    'cPos = CursorPosition
    'Call FillRectAttribute(bPos, mint_constructor.Position(cPos.Left - 1, cPos.Top), -1, Attributes)
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_WRITE)
End Sub
Private Sub internal_WriteByteArray(bArray() As Byte)
    Dim arrLen As Long
    arrLen = ArraySize(bArray)
    If arrLen = 0 Then Exit Sub
    Dim numofCharsWritten As Long
    'Dim bPos As Position, cPos As Position
    'bPos = CursorPosition

    'Call API_WriteConsole(out_hndl, bArray(0), arrLen, numofCharsWritten, 1)
    If Not API_WriteConsole(out_hndl, bArray(0), arrLen, numofCharsWritten, 1) Then _
        throw Exps.IfError
    'cPos = CursorPosition
    'Call FillRectAttribute(bPos, mint_constructor.Position(cPos.Left - 1, cPos.Top), -1, Attributes)
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_WRITE)
End Sub

Public Sub FillLine(StartPosition As Position, Count As Long, Character As Byte, Optional ByVal BackColor As ConsoleColors = ccDefault, Optional ByVal ForeColor As ConsoleColors = ccDefault)
    Dim dwI As Long  ' Characters written in buffer, not used, but necessary output for later function
    Dim CAttr As Long
    Dim dwSize As Long ' Characters to clear ("area" of console, so to speak)
    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO
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
    If Not API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then _
        throw Exps.IfError
    MaxRect_Bounds = (CLng(CSBINFO.dwSize.X) * CLng(CSBINFO.dwSize.Y))
    If (bfTop * bfLeft) >= MaxRect_Bounds Then Exit Sub
    If bfLeft + dwSize > CLng(CSBINFO.dwSize.X) Then dwSize = CLng(CSBINFO.dwSize.X) - bfLeft
    If Character <> -1 Then _
        If Not API_FillConsoleOutputCharacter(out_hndl, Character Mod 255, dwSize, pos, dwI) Then _
            throw Exps.IfError ' Write char to buffer as many times as we tell it, i.e. dwSize
    If Not API_FillConsoleOutputAttribute(out_hndl, CAttr, dwSize, pos, dwI) Then _
        throw Exps.IfError
    'CursorPosition = mint_constructor.Position(0, 0) ' Set console cursor to top left coord
End Sub
Public Sub FillRect(StartPosition As Position, EndPosition As Position, Character As Long, Optional ByVal BackColor As ConsoleColors = ccDefault, Optional ByVal ForeColor As ConsoleColors = ccDefault)
    Dim dwI As Long  ' Characters written in buffer, not used, but necessary output for later function
    Dim CAttr As Long
    Dim dwSize As Long ' Characters to clear ("area" of console, so to speak)
    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO
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
    If Not API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then _
        throw Exps.IfError
    If bfLeft + dwSize > CLng(CSBINFO.dwSize.X) Then dwSize = CLng(CSBINFO.dwSize.X) - bfLeft
    For i = 0 To bfLinesCount - 1
        pos = ((StartPosition.Top + i) * &H10000) + bfLeft
        If Character <> -1 Then _
            Call API_FillConsoleOutputCharacter(out_hndl, Character Mod 255, dwSize, pos, dwI) ' Write char to buffer as many times as we tell it, i.e. dwSize
        Call API_FillConsoleOutputAttribute(out_hndl, CAttr, dwSize, pos, dwI)
        'CursorPosition = mint_constructor.Position(0, 0) ' Set console cursor to top left coord
    Next
End Sub
Public Sub FillRectAttribute(StartPosition As Position, EndPosition As Position, Character As Long, ByVal Attributes As Long)
    Dim dwI As Long  ' Characters written in buffer, not used, but necessary output for later function
    Dim dwSize As Long ' Characters to clear ("area" of console, so to speak)
    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO
    Dim bfLeft As Long, bfLinesCount As Long
    Dim pos As Long

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
    If Not API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then _
        throw Exps.IfError
    If bfLeft + dwSize > CLng(CSBINFO.dwSize.X) Then dwSize = CLng(CSBINFO.dwSize.X) - bfLeft
    For i = 0 To bfLinesCount - 1
        pos = ((StartPosition.Top + i) * &H10000) + bfLeft
        If Character <> -1 Then _
            Call API_FillConsoleOutputCharacter(out_hndl, Character Mod 255, dwSize, pos, dwI) ' Write char to buffer as many times as we tell it, i.e. dwSize
        Call API_FillConsoleOutputAttribute(out_hndl, Attributes, dwSize, pos, dwI)
        'CursorPosition = mint_constructor.Position(0, 0) ' Set console cursor to top left coord
    Next
End Sub
Public Sub SetChar(CharPos As Position, ByVal Char As Byte)
    If Len(Char) <> 1 Then throw Exps.InvalidArgumentException("Character length must be 1.")
    Call FillLine(CharPos, 1, Char)
End Sub

Private Function CreateRange(fromI As Byte, toI As Byte) As Console_Input_Available_Record
    CreateRange.LPart = fromI
    CreateRange.UPart = toI
End Function
Private Function CreateRange1(Character As Byte) As Console_Input_Available_Record
    CreateRange1.LPart = Character
    CreateRange1.UPart = Character
End Function
Private Function Array_Range(ParamArray Bytes() As Variant) As Console_Input_Available_Record()
    Dim i As Long, btSize As Long, RetVal() As Console_Input_Available_Record
    On Error GoTo zeroLength
    btSize = UBound(Bytes) - LBound(Bytes) + 1
zeroLength:
    If btSize > 0 Then
        ReDim RetVal(btSize - 1)
        For i = 0 To btSize - 1
            RetVal(i) = Bytes(i)
        Next
    End If
    Array_Range = RetVal()
End Function

Public Function ReadChar() As Byte
    Dim BytesTransferred As Long
    Dim Char As API_INPUT_RECORD

    If Not API_FlushConsoleInputBuffer(in_hndl) Then _
        throw Exps.IfError
    'Operation(OPERATION_CONSOLE_READSTATE).State = False
looper:
'    If Not API_ReadConsoleInput(in_hndl, Char, 1, BytesTransferred) Then _
'        throw Exps.IfError
    If Char.EventType = KEY_EVENT Then
        ReadChar = Char.Event.KeyEvent.UCHAR
        If (ReadChar <> 0) And _
            CBool(Char.Event.KeyEvent.bKeyDown) Then Exit Function
    End If
End Function
Public Function ReadKey() As Integer
    Dim BytesTransferred As Long
    Dim Char As API_INPUT_RECORD

    If Not API_FlushConsoleInputBuffer(in_hndl) Then _
        throw Exps.IfError
'    Operation(OPERATION_CONSOLE_READSTATE).State = False
looper:

'    If Not API_ReadConsoleInput(in_hndl, Char, 1, BytesTransferred) Then _
'        throw Exps.IfError
    If Char.EventType = KEY_EVENT Then
        ReadKey = Char.Event.KeyEvent.wVirtualKeyCode
        If (ReadKey <> 0) And _
            CBool(Char.Event.KeyEvent.bKeyDown) Then Exit Function
    End If
    Call API_Sleep(10)
'If Not Operation(OPERATION_CONSOLE_READSTATE).State Then _
'    GoTo looper
End Function
Private Function CalculateLineRange(Str As String, ByVal StartFrom As Long, outLineStart As Long, outLineEnd As Long)
    'CrLf
    Const BREAKEDWITHCR As Long = 1
    Const BREAKEDWITHLF As Long = 2
    Dim i As Long, ln As Long, cChar As String * 1
    ln = Len(Str)
    If StartFrom < 0 And StartFrom > ln Then
        outLineStart = 0
        outLineEnd = 0
        Exit Function
    End If
    For i = StartFrom To ln
        cChar = Mid(Str, i, 1)
        If cChar = vbCr Then
            Exit For
        ElseIf cChar = vbLf Then
            If Mid(Str, i - 1, 1) = vbLf Then
                i = i - 1
                Exit For
            End If
        End If
    Next
    outLineEnd = IIf(i = StartFrom, StartFrom, i - 1)
    For i = StartFrom To 1 Step -1
        cChar = Mid(Str, i, 1)
        If cChar = vbCr Or cChar = vbLf Then _
            Exit For
    Next
    If i > outLineEnd Then
        outLineStart = StartFrom
        outLineEnd = StartFrom
    Else
        outLineStart = i + 1
    End If
End Function
'-------------------------------------------------
'==================================================
'-------------------------------------------------
' Reads a dynamic string from console input.
' reserved keys:
' Home,End,Arrow Up,Arrow Down,Arrow Left,Arrow Right
' this also can be written by FileRead() Win32Api
' but i wanted to write my own code for educational purposes !! :)
Public Function ReadString( _
        AvailableInputs() As Console_Input_Available_Record, _
        Length As Long, BreakAtCharacters() As Long, _
        BackClearerChar As Byte, FrontClearerChar As Byte, _
        outBreakChar As Byte, outinStartAt As Position, outEndAt As Position, _
        NumberAcceptedRange_Lower, NumberAcceptedRange_Upper, _
        Optional ByVal BackColor As ConsoleColors = ccDefault, _
        Optional ByVal ForeColor As ConsoleColors = ccDefault) As String

    If Length <= 0 Then Exit Function
    Dim i As Long, breakerLength As Long, availableInputsLength As Long
    Dim BytesTransferred As Long
    Dim Char As API_INPUT_RECORD
    Dim RetVal As String
    Dim retVal_Len As Long
    Dim retVal_CIndex As Long
    Dim chChar As String
    Dim AvailableInputs_() As Console_Input_Available_Record

    Dim haveIntegeralCheck As Boolean
    Dim insertState As Boolean

    Dim start_X As Long, start_Y As Long
    Dim maxX As Long, maxY As Long
    Dim currentRelativeX As Long, currentRelativeY As Long

    Dim CSBINFO As API_CONSOLE_SCREEN_BUFFER_INFO

    Dim buffSize As Size
    buffSize = BufferSize
    maxX = buffSize.Width
    maxY = buffSize.Height

    breakerLength = ArraySize(BreakAtCharacters)
    If breakerLength <= 0 Then throw Exps.InvalidOperationException("BreakAtCharacters is empty ,this may cause in infinite loop.")
    On Error GoTo AvailableInputsZeroLength
    availableInputsLength = (UBound(AvailableInputs) - LBound(AvailableInputs) + 1)
AvailableInputsZeroLength:
    On Error GoTo 0
    If availableInputsLength > 0 Then
        AvailableInputs_ = AvailableInputs
    Else
        ReDim AvailableInputs_(0)
        AvailableInputs_(0).LPart = 0
        AvailableInputs_(0).UPart = 255
        availableInputsLength = 1
    End If

    If NumberAcceptedRange_Lower = 0 And _
        NumberAcceptedRange_Upper = 0 Then
        haveIntegeralCheck = False
    Else
        haveIntegeralCheck = True
    End If

    'Call API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO)
    If Not API_GetConsoleScreenBufferInfo(out_hndl, CSBINFO) Then _
        throw Exps.IfError

    start_X = CSBINFO.dwCursorPosition.X
    start_Y = CSBINFO.dwCursorPosition.Y

    outBreakChar = 0
    insertState = Me.InsertMode

    Dim must_modify As Boolean
    If BackColor <> ccDefault Then BackColor = CSBINFO.wAttributes And &HF: must_modify = True
    If ForeColor <> ccDefault Then ForeColor = CSBINFO.wAttributes And &HF0: must_modify = True

    If must_modify Then
        If Not API_SetConsoleTextAttribute(out_hndl, (ForeColor * &H10) + BackColor) Then _
            throw Exps.IfError
    End If

    If (outinStartAt.Left >= 0 And outinStartAt.Top >= 0) And _
        (outEndAt.Left > 0 And outEndAt.Top > 0) Then
        CursorPosition = outinStartAt
    Else
        outinStartAt = CursorPosition
    End If

    If Not API_FlushConsoleInputBuffer(in_hndl) Then _
        throw Exps.IfError
'    Operation(OPERATION_CONSOLE_READSTATE).State = False
    Dim vKey As Long
looper:
If in_hndl = INVALID_HANDLE_VALUE Then throw Exps.InvalidOperationException
'Function ReadChar ============================================================
func_readchar_loop:                                                          '|
'        SUCCESS = API_ReadConsoleInput(in_hndl, Char, 1, BytesTransferred)   '|
'        If SUCCESS = 0 Then throw Exps.IfError
        If Char.EventType = KEY_EVENT Then                                   '|
            vKey = Char.Event.KeyEvent.wVirtualKeyCode                       '|
            If Char.Event.KeyEvent.UCHAR <> 0 Then                           '|
                If CBool(Char.Event.KeyEvent.bKeyDown) Then _
                    GoTo func_readchar_exit                                  '|
            ElseIf (vKey = vbKeyUp Or vKey = vbKeyDown Or _
                    vKey = vbKeyLeft Or vKey = vbKeyRight Or _
                    vKey = vbKeyHome Or vKey = vbKeyEnd) Then                '|
                If CBool(Char.Event.KeyEvent.bKeyDown) Then _
                    GoTo func_readchar_exit                                  '|
            End If                                                           '|
        End If                                                               '|
        If supportUIEventProcessing Then                                     '|
            If Char.EventType <> MOUSE_EVENTC Then DoEvents                  '|
        End If                                                               '|
        Call API_Sleep(10)                                                   '|
'If Not Operation(OPERATION_CONSOLE_READSTATE).State Then GoTo func_readchar_loop              '|
func_readchar_exit:                                                          '|
'------------------------------------------------------------------------------
    Dim chCode As Long
    Dim lChar As String * 1
    chCode = Char.Event.KeyEvent.UCHAR
    chChar = Chr(chCode)
    If chCode = BackClearerChar Then 'Backspace
        If retVal_Len <= 0 Then
            Call VBA.Beep
            GoTo write_ended
        Else
            If (retVal_CIndex - 1) >= 0 Then
                i = 0
                lChar = Mid(RetVal, retVal_CIndex, 1)
                If lChar = vbCr Or lChar = vbLf Then
                    If currentRelativeY - 1 >= 0 Then
                        Dim start_i, cLineLength As Long, cLine As Long
                        cLine = 0
                        cLineLength = 0
                        start_i = retVal_CIndex + 1
                        For i = retVal_CIndex - 1 To 1 Step -1
                            lChar = Mid(RetVal, i, 1)
                            If lChar = vbCr Or lChar = vbLf Or lChar = vbCrLf Then
                                start_i = i + 1
                                Exit For
                            End If
                        Next
                        If start_i > retVal_Len Then start_i = retVal_Len
                        For i = retVal_CIndex + 1 To retVal_Len
                            lChar = Mid(RetVal, i, 1)
                            If lChar = vbCr Or lChar = vbLf Or lChar = vbCrLf Then
                                Call FillLine(mint_constructor.Position(IIf(cLine + currentRelativeY <= 0, start_X, 0) + currentRelativeX, start_Y + currentRelativeY + cLine), cLineLength, 0)
                                cLine = cLine + 1
                                cLineLength = 0
                            Else
                                cLineLength = cLineLength + 1
                            End If
                        Next
                        Call FillLine(mint_constructor.Position(IIf(cLine + currentRelativeY <= 0, start_X, 0) + currentRelativeX, start_Y + currentRelativeY + cLine), cLineLength, 0)
                        currentRelativeY = currentRelativeY - 1
                        currentRelativeX = 0
                        For i = retVal_CIndex - 1 To 1 Step -1
                            lChar = Mid(RetVal, i, 1)
                            If lChar = vbCr Or lChar = vbLf Or lChar = vbCrLf Then _
                                Exit For
                            currentRelativeX = currentRelativeX + 1
                        Next
                        'currentRelativeX = i - 1 'DEBUG
                    Else
                        If currentRelativeX - 1 >= 0 Then currentRelativeX = currentRelativeX - 1
                    End If
                Else
                    If currentRelativeX - 1 >= 0 Then currentRelativeX = currentRelativeX - 1
                End If
                Dim cuX As Long, cuY As Long
                cuX = 0
                cuY = 0
                Dim iCh As Long
                For iCh = 1 To retVal_Len
                    lChar = Mid(RetVal, iCh, 1)
                    If lChar = vbCr Or lChar = vbLf Then
                        cuX = 0
                        cuY = cuY + 1
                    Else
                        cuX = cuX + 1
                    End If
                Next
                Call SetChar(mint_constructor.Position(IIf(cuY = 0, start_X, 0) + cuX - 1, start_Y + cuY), 0)
                retVal_CIndex = retVal_CIndex - 1
                Dim rVal As String
                rVal = Right(RetVal, retVal_Len - retVal_CIndex - 1)
                RetVal = Left(RetVal, retVal_CIndex) & rVal
                CursorPosition = mint_constructor.Position(IIf(i <= 0, start_X, 0) + currentRelativeX, start_Y + currentRelativeY)
                Call WriteStr(rVal)
                CursorPosition = mint_constructor.Position(IIf(i <= 0, start_X, 0) + currentRelativeX, start_Y + currentRelativeY)
                retVal_Len = retVal_Len - 1
            Else
                Call VBA.Beep
                GoTo write_ended
            End If
        End If
    ElseIf chCode = FrontClearerChar Then 'DEL
        If retVal_Len <= 0 Then
            Call VBA.Beep
            GoTo write_ended
        Else
            If retVal_CIndex >= retVal_Len Then
                Call VBA.Beep
                GoTo write_ended
            Else
                Title = "Front Clearer"
            End If
        End If
    ElseIf vKey = vbKeyLeft Then '<
        If retVal_CIndex - 1 >= 0 Then
            retVal_CIndex = retVal_CIndex - 1
            If currentRelativeX - 1 >= 0 Then
                currentRelativeX = currentRelativeX - 1
            ElseIf currentRelativeY > 0 Then
                Dim lS As Long, lE As Long
                If retVal_CIndex - 2 > 0 Then _
                    Call CalculateLineRange(RetVal, retVal_CIndex - 2, lS, lE)
                If lS <> lE Then
                    currentRelativeY = currentRelativeY - 1
                    currentRelativeX = (lE - lS) + 1
                End If
            End If
            If currentRelativeY = 0 Then
                CursorPosition = mint_constructor.Position(start_X + currentRelativeX, start_Y + currentRelativeY)
            Else
                CursorPosition = mint_constructor.Position(currentRelativeX, start_Y + currentRelativeY)
            End If
        End If
    ElseIf vKey = vbKeyRight Then '>
        If retVal_CIndex + 1 <= retVal_Len Then
            retVal_CIndex = retVal_CIndex + 1
            If Mid(RetVal, retVal_CIndex, 1) = vbCr Then
                currentRelativeX = 0
                currentRelativeY = currentRelativeY + 1
            ElseIf currentRelativeX + 1 <= maxX Then
                currentRelativeX = currentRelativeX + 1
            Else
                currentRelativeX = 0
                currentRelativeY = currentRelativeY + 1
            End If
            If currentRelativeY = 0 Then
                CursorPosition = mint_constructor.Position(start_X + currentRelativeX, start_Y + currentRelativeY)
            Else
                CursorPosition = mint_constructor.Position(currentRelativeX, start_Y + currentRelativeY)
            End If
        End If
    ElseIf vKey = vbKeyUp Then '^
        If currentRelativeY - 1 >= 0 Then

        Else
            currentRelativeX = 0
        End If
        CursorPosition = mint_constructor.Position(start_X + currentRelativeX, start_Y + currentRelativeY)
    ElseIf vKey = vbKeyDown Then 'v
        If currentRelativeY + 1 >= 0 Then

        Else
            currentRelativeX = 0
        End If
        CursorPosition = mint_constructor.Position(start_X + currentRelativeX, start_Y + currentRelativeY)
    ElseIf vKey = vbKeyHome Then 'Home
        retVal_CIndex = 0
        If currentRelativeX = 0 Then
            currentRelativeX = 0
            currentRelativeY = 0
        Else
            currentRelativeX = 0
        End If
        CursorPosition = mint_constructor.Position(start_X + currentRelativeX, start_Y + currentRelativeY)
    ElseIf vKey = vbKeyEnd Then 'End
        Dim crPos As Long
        For i = retVal_CIndex To retVal_Len
            If Mid(RetVal, i, 1) = vbCr Then Exit For
            crPos = crPos + 1
        Next
        If retVal_CIndex = crPos Then
            For i = retVal_CIndex To retVal_Len
                If Mid(RetVal, i, 1) = vbCr Then
                    currentRelativeY = currentRelativeY + 1
                    currentRelativeX = 0
                End If
                currentRelativeX = currentRelativeX + 1
            Next
            retVal_CIndex = retVal_Len - 1
        Else
            currentRelativeX = crPos
            retVal_CIndex = retVal_CIndex + crPos
        End If
        CursorPosition = mint_constructor.Position(start_X + currentRelativeX, start_Y + currentRelativeY)
    Else
        For i = 0 To breakerLength - 1
            If chCode = BreakAtCharacters(i) Then
                outBreakChar = chCode
                GoTo outside_of_looper
            End If
        Next

        Dim execOnce As Boolean
        execOnce = False
        If retVal_Len < Length Then
        For i = 0 To availableInputsLength - 1
            If chCode >= AvailableInputs_(i).LPart And chCode <= AvailableInputs_(i).UPart Then
                execOnce = True
                If haveIntegeralCheck Then
                    Dim integeralCheck_Value As Double
                    integeralCheck_Value = VBA.Val(RetVal & chChar)
                    If Not ((integeralCheck_Value >= NumberAcceptedRange_Lower And _
                        integeralCheck_Value <= NumberAcceptedRange_Upper)) Then
                        Call VBA.Beep
                        GoTo write_ended
                    End If
                End If
                If chChar = vbCr Or chChar = vbLf Or chChar = vbCrLf Then
                    chChar = vbCr
                    Dim fillLength As Long, cInd As Long
                    fillLength = 0
                    For cInd = retVal_CIndex + 1 To retVal_Len
                        lChar = Mid(RetVal, cInd, 1)
                        If lChar = vbCr Or lChar = vbLf Or lChar = vbCrLf Then
                            Exit For
                        Else
                            fillLength = fillLength + 1
                        End If
                    Next
                    If currentRelativeY = 0 Then
                        Call FillLine(mint_constructor.Position(start_X + currentRelativeX, start_Y + currentRelativeY), fillLength, 0)
                    Else
                        Call FillLine(mint_constructor.Position(currentRelativeX, start_Y + currentRelativeY), fillLength, 0)
                    End If
                    CursorPosition = mint_constructor.Position(0, start_Y + currentRelativeY)
                    currentRelativeY = currentRelativeY + 1
                    currentRelativeX = 0
                Else 'Codes below doesnt tested yet.
                    Dim cX As Long
                    If currentRelativeY = 0 Then
                        cX = start_X + currentRelativeX
                    Else
                        cX = currentRelativeX
                    End If
                    If cX + 1 <= maxX Then
                        currentRelativeX = currentRelativeX + 1
                        If currentRelativeY = 0 Then
                            CursorPosition = mint_constructor.Position(start_X + currentRelativeX - 1, start_Y + currentRelativeY)
                        Else
                            CursorPosition = mint_constructor.Position(currentRelativeX - 1, start_Y + currentRelativeY)
                        End If
                    Else
                        currentRelativeY = currentRelativeY + 1
                        currentRelativeX = 1
                        CursorPosition = mint_constructor.Position(0, start_Y + currentRelativeY)
                    End If
                End If
                RetVal = Left(RetVal, retVal_CIndex) & chChar & Right(RetVal, retVal_Len - retVal_CIndex)
                retVal_Len = retVal_Len + 1
                retVal_CIndex = retVal_CIndex + 1
                Call WriteStr(Mid(RetVal, retVal_CIndex, retVal_Len - retVal_CIndex + 1))
                'End If
                If currentRelativeY = 0 Then
                    CursorPosition = mint_constructor.Position(start_X + currentRelativeX, start_Y + currentRelativeY)
                Else
                    CursorPosition = mint_constructor.Position(currentRelativeX, start_Y + currentRelativeY)
                End If
write_ended:
                Exit For
            End If
        Next
        If Not execOnce Then Call VBA.Beep
        Else
            Call VBA.Beep
        End If
    End If

'If Not Operation(OPERATION_CONSOLE_READSTATE).State Then _
'    GoTo looper

outside_of_looper:

    If must_modify Then _
        Call API_SetConsoleTextAttribute(out_hndl, CSBINFO.wAttributes)

    outEndAt = CursorPosition

    ReadString = Left(RetVal, retVal_Len)
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_READ)
End Function 'ReadString
'-------------------------------------------------
'==================================================
'-------------------------------------------------
Public Function ReadLine(Optional ByVal BackColor As ConsoleColors = ccDefault, Optional ByVal ForeColor As ConsoleColors = ccDefault) As String
    Dim breakChar As Byte, bts() As Long
    Dim AvailableInputs() As Console_Input_Available_Record
'    ReDim AvailableInputs(0) 'All characters.
'    AvailableInputs(0).LPart = 0
'    AvailableInputs(0).UPart = 255
    'bts = Array_Long(10, 13, 27)
    ReadLine = ReadString(AvailableInputs, LinearBufferSize, bts, BACKSPACECHAR, DELETECHAR, breakChar, mint_constructor.Position(0, 0), mint_constructor.Position(0, 0), 0, 0, BackColor, ForeColor)
    Call WriteStr(vbCrLf)
    If breakChar = 27 Then ReadLine = ""
End Function
Public Function ReadStr(Optional ByVal BackColor As ConsoleColors = ccDefault, Optional ByVal ForeColor As ConsoleColors = ccDefault) As String
    Dim breakChar As Byte, bts() As Long
    Dim AvailableInputs() As Console_Input_Available_Record
'    ReDim AvailableInputs(0) 'All characters.
'    AvailableInputs(0).LPart = 0
'    AvailableInputs(0).UPart = 255
    'bts = Array_Long(10, 13, 27)
    ReadStr = ReadString(AvailableInputs, LinearBufferSize, bts, BACKSPACECHAR, DELETECHAR, breakChar, mint_constructor.Position(0, 0), mint_constructor.Position(0, 0), 0, 0, BackColor, ForeColor)
    If breakChar = 27 Then ReadStr = ""
End Function
Public Sub ReadVar(Var, Optional ByVal BackColor As ConsoleColors = ccDefault, Optional ByVal ForeColor As ConsoleColors = ccDefault)
    Dim breakChar As Byte, varBuff As Variant, bts() As Long
    Dim AvailableInputs() As Console_Input_Available_Record
    'bts = Array_Long(10, 13, 27)
    If IsObject(Var) Then
        Set varBuff = Var
    Else
        varBuff = Var
    End If


    Dim i As Long, outBreaker As Byte
    Dim retStr As String
    Dim Avail_Nums(0) As Console_Input_Available_Record
    Avail_Nums(0).LPart = 48
    Avail_Nums(0).UPart = 57
    'Var = ReadString(AvailableInputs, Array_Byte(13, 27), BACKSPACECHAR, breakChar, BackColor, ForeColor)

    On Error GoTo Err
    Select Case VarType(Var)
        Case VbVarType.vbString
            retStr = ReadStr(BackColor, ForeColor)
            Var = retStr
        Case VbVarType.vbArray
            On Error Resume Next
            'Call mySignalEmitter.BlockSignals
            For i = 0 To ArraySize(Var) - 1
                Call ReadVar(Var(i), BackColor, ForeColor)
            Next
            On Error GoTo 0
            'Call mySignalEmitter.UnblockSignals
            'Call mySignalEmitter.Emit(SIGNAL_READ)
        Case VbVarType.vbBoolean
            retStr = ReadStr(BackColor, ForeColor)
            retStr = LCase(retStr)
            If retStr = "true" Or retStr = "1" Or retStr = "t" Or Val(retStr) <> 0 Then
                Var = True
            Else
                Var = False
            End If
        Case VbVarType.vbByte
            retStr = ReadString( _
                Array_Range(CreateRange(48, 57), CreateRange1(43)), _
                4, bts, BACKSPACECHAR, DELETECHAR, outBreaker, mint_constructor.Position(0, 0), mint_constructor.Position(0, 0), 0, 255, BackColor, ForeColor)
            If retStr = "" Then
                Var = 0
            Else
                Var = CByte(retStr)
            End If
        Case VbVarType.vbInteger
            retStr = ReadString( _
                Array_Range(CreateRange(48, 57), CreateRange1(45), CreateRange1(43)), _
                6, bts, BACKSPACECHAR, DELETECHAR, outBreaker, mint_constructor.Position(0, 0), mint_constructor.Position(0, 0), INTEGER_MIN, INTEGER_MAX, BackColor, ForeColor)
            If retStr = "" Then
                Var = 0
            Else
                Var = CInt(retStr)
            End If
        Case VbVarType.vbLong
            retStr = ReadString( _
                Array_Range(CreateRange(48, 57), CreateRange1(45), CreateRange1(43)), _
                11, bts, BACKSPACECHAR, DELETECHAR, outBreaker, mint_constructor.Position(0, 0), mint_constructor.Position(0, 0), LONG_MIN, LONG_MAX, BackColor, ForeColor)
            If retStr = "" Then
                Var = 0
            Else
                Var = CLng(retStr)
            End If
        Case VbVarType.vbDouble
            retStr = ReadString( _
                Array_Range(CreateRange(48, 57), CreateRange1(45), CreateRange1(43), CreateRange1(46), CreateRange1(69), CreateRange1(101), CreateRange1(102), CreateRange1(70)), _
                26, bts, BACKSPACECHAR, DELETECHAR, outBreaker, mint_constructor.Position(0, 0), mint_constructor.Position(0, 0), 0, 0, BackColor, ForeColor)
            If retStr = "" Then
                Var = 0
            Else
                Var = CDbl(retStr)
            End If
        Case VbVarType.vbSingle
            retStr = ReadString( _
                Array_Range(CreateRange(48, 57), CreateRange1(45), CreateRange1(43), CreateRange1(46), CreateRange1(69), CreateRange1(101), CreateRange1(102), CreateRange1(70)), _
                16, bts, BACKSPACECHAR, DELETECHAR, outBreaker, mint_constructor.Position(0, 0), mint_constructor.Position(0, 0), 0, 0, BackColor, ForeColor)
            Var = CSng(retStr)
        Case VbVarType.vbCurrency
            'retStr = ReadStr(BackColor, ForeColor)
            retStr = ReadString(Avail_Nums, LinearBufferSize, bts, _
                BACKSPACECHAR, DELETECHAR, outBreaker, mint_constructor.Position(0, 0), mint_constructor.Position(0, 0), BackColor, ForeColor)
            If retStr = "" Then
                Var = 0
            Else
                Var = CCur(retStr)
            End If
        Case VbVarType.vbDate
            'retStr = ReadStr(BackColor, ForeColor)
            retStr = ReadString(Avail_Nums, LinearBufferSize, bts, _
                BACKSPACECHAR, DELETECHAR, outBreaker, mint_constructor.Position(0, 0), mint_constructor.Position(0, 0), BackColor, ForeColor)
            If retStr = "" Then
                Var = Now
            Else
                Var = CDate(retStr)
            End If
        Case Else
            throw Exps.InvalidArgumentException("Var is unknown.")
    End Select
    GoTo continue_func
Err:
    throw Exps.Exception("An error occured in Console.ReadVar(...) {" & Err.Description & "}")
continue_func:
    If breakChar = 27 Then
        If IsObject(varBuff) Then
            Set Var = varBuff
        Else
            Var = varBuff
        End If
    End If
End Sub

Public Property Get IsKeyPressed() As Boolean
    Dim BytesTransferred As Long
    Dim Char As API_INPUT_RECORD
    Dim SUCCESS As Long

'    SUCCESS = API_ReadConsoleInput(in_hndl, Char, 1, BytesTransferred)
    If Char.EventType = KEY_EVENT Then
        If Char.Event.KeyEvent.UCHAR <> 0 Then _
            IsKeyPressed = CBool(Char.Event.KeyEvent.bKeyDown)
    End If
End Property

Public Property Get CapsLock() As Boolean
    Dim BytesTransferred As Long
    Dim Char As API_INPUT_RECORD
    Dim SUCCESS As Long

    'SUCCESS = API_ReadConsoleInput(in_hndl, Char, 1, BytesTransferred)
    If Char.EventType = KEY_EVENT Then
        If Char.Event.KeyEvent.wVirtualKeyCode = vbKeyCapital Then _
            CapsLock = CBool(Char.Event.KeyEvent.bKeyDown)
    End If
End Property
Public Property Get NumLock() As Boolean
    Dim BytesTransferred As Long
    Dim Char As API_INPUT_RECORD
    Dim SUCCESS As Long

    'SUCCESS = API_ReadConsoleInput(in_hndl, Char, 1, BytesTransferred)
    If Char.EventType = KEY_EVENT Then
        If Char.Event.KeyEvent.wVirtualKeyCode = vbKeyNumlock Then _
            NumLock = CBool(Char.Event.KeyEvent.bKeyDown)
    End If
End Property
Public Property Get ScrollLock() As Boolean
    Dim BytesTransferred As Long
    Dim Char As API_INPUT_RECORD
    Dim SUCCESS As Long

    'SUCCESS = API_ReadConsoleInput(in_hndl, Char, 1, BytesTransferred)
    If Char.EventType = KEY_EVENT Then
        If Char.Event.KeyEvent.wVirtualKeyCode = vbKeyScrollLock Then _
            ScrollLock = CBool(Char.Event.KeyEvent.bKeyDown)
    End If
End Property
Public Property Get InsertMode() As Boolean

End Property
Public Property Let InsertMode(Value As Boolean)

End Property

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter.ExternEmitter
End Function
Private Property Get ISupportSlot_AutoManage() As Boolean
    ISupportSlot_AutoManage = False
End Property
Private Sub ISupportSlot_ManualManageCall(SlotName As String)
    Select Case SlotName
        Case SIGNAL_CLEAR
            Call Clear
        Case SIGNAL_BEEP
            Call Beep
        Case SIGNAL_RESETCOLORS
            Call ResetColors
        Case Else
            throw Exps.InvalidCallException
    End Select
End Sub
Private Function ISupportSlot_Methods() As String
    ISupportSlot_Methods = _
        "1" & SLOT_CLEAR & _
        " 0" & SLOT_BEEP & _
        " 0" & SLOT_RESETCOLORS
End Function

Private Function ITargetStream_Available(Dir As StreamDirection) As Long
    
End Function

'<NOT CLOSABLE>
Private Sub ITargetStream_CloseStream(Optional Direction As StreamDirection = 3&)
    If (Direction And sdInStream) = sdInStream Then
        inState = False
    End If
    If (Direction And sdOutStream) = sdOutStream Then
        outState = False
    End If
End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean
    Dim Val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        Val = inState
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        Val = Val Or outState
    End If
    ITargetStream_getState = Val
End Function

Private Sub ITargetStream_inStream(Data As Variant, Optional Length As Long = -1&)
    If Not inState Then throw Exps.InvalidOperationException
    If IsObject(Data) Then
        If TypeOf Data Is ObjectBuffer Then
            Dim objBuffer As ObjectBuffer
            Set objBuffer = Data
            If objBuffer.IsStreamingValue Then
                Call objBuffer.thrownotargs(1)
                If objBuffer.ArgType(0) = vbString Then
                    Dim strFirst As String
                    strFirst = objBuffer.First
                    If strFirst = OBJECTBUFFER_VALUES_STREAMING_CLEAR Then
                        Call Clear
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_NEWLINE Then
                        Call WriteStr(vbCrLf)
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_FLUSH Then
                        'Do nothing...
                        Exit Sub
                    Else
                        throw Exps.InvalidArgumentException("Streaming Command '" & strFirst & "' Is Not Available For Console.")
                    End If
                Else
                    throw Exps.InvalidArgumentException("Streaming Type Is Not Available For Console.")
                End If
            Else
                throw Exps.InvalidArgumentException
            End If
        Else
            throw Exps.InvalidArgumentException
        End If
        throw Exps.InvalidArgumentException
    End If
    Call ReadVar(Data)
End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    Dim rtLength As Long
    Dim Avai_Input() As Console_Input_Available_Record
    Dim braChar() As Long
    Dim outBreak As Byte
    'braChar = Array_Long(10, 13, 27)
    rtLength = Length
    ITargetStream_readBytes = StringToByteArray(ReadString(Avai_Input, rtLength, braChar, BACKSPACECHAR, DELETECHAR, outBreak, mint_constructor.Position(0, 0), mint_constructor.Position(0, 0), 0, 0))
End Function

Private Sub ITargetStream_OpenStream(Optional Direction As StreamDirection = StreamDirection.sdBoth)
    If (Direction And sdInStream) = sdInStream Then
        inState = True
    End If
    If (Direction And sdOutStream) = sdOutStream Then
        outState = True
    End If
End Sub
Private Sub ITargetStream_outStream(Data As Variant, Optional Length As Long = -1)
    If Not outState Then throw Exps.InvalidOperationException
    If IsObject(Data) Then
        If TypeOf Data Is ObjectBuffer Then
            Dim objBuffer As ObjectBuffer
            Set objBuffer = Data
            If objBuffer.IsStreamingValue Then
                Call objBuffer.thrownotargs(1)
                If objBuffer.ArgType(0) = vbString Then
                    Dim strFirst As String
                    strFirst = objBuffer.First
                    If strFirst = OBJECTBUFFER_VALUES_STREAMING_CLEAR Then
                        Call Clear
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_NEWLINE Then
                        Call WriteStr(vbCrLf)
                        Exit Sub
                    Else
                        throw Exps.InvalidArgumentException("Streaming Command '" & strFirst & "' Is Not Available For ByteArray.")
                    End If
                Else
                    throw Exps.InvalidArgumentException("Streaming Type Is Not Available For ByteArray.")
                End If
            End If
        End If
        Call internal_WriteByteArray(mint_get_byte_array_of(Data, Length))
        Exit Sub
    ElseIf IsArray(Data) Then
        Call internal_WriteByteArray(mint_get_byte_array_of(Data))
        Exit Sub
    End If
    Dim btOut() As Byte
    If Length = -1 Then
        btOut = StringToByteArray(CStr(Data))
    Else
        btOut = StringToByteArray(Left(CStr(Data), Length))
    End If
    Call internal_WriteByteArray(btOut)
End Sub

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved)
    Call mint_setstream_state(State, inState, inState_LOCK, Reserved)
End Sub
Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved)
    Call mint_setstream_state(State, outState, outState_LOCK, Reserved)
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property



Private Function IClassStream_Available(ByVal Direction As StreamDirection) As Long
    Dim Val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        Val = ((in_hndl <> INVALID_HANDLE_VALUE) And inState)
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        Val = Val Or ((out_hndl <> INVALID_HANDLE_VALUE) And outState)
    End If
    IClassStream_Available = Val
End Function
Private Sub IClassStream_CloseStream(Optional ByVal Direction As StreamDirection = 3&)

End Sub
Private Function IClassStream_GetState(ByVal Dir As StreamDirection) As Boolean
    
End Function
Private Sub IClassStream_InStream(ByVal Buffer As ClassStreamBuffer, Optional ByVal Length As Long = -1&)

End Sub
Private Sub IClassStream_OpenStream(Optional ByVal Direction As StreamDirection = 3&)

End Sub
Private Sub IClassStream_OutStream(ByVal Buffer As ClassStreamBuffer, Optional ByVal Length As Long = -1&)

End Sub
Private Function IClassStream_ReadBytes(ByVal Length As Long, Optional ByVal StartAt As Long = 0&) As Byte()

End Function
Private Sub IClassStream_SetInState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)

End Sub
Private Sub IClassStream_SetOutState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)

End Sub
