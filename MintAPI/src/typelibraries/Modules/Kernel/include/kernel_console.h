#ifndef __KERNEL_CONSOLE_H__
#define __KERNEL_CONSOLE_H__


#define MAX_DEFAULTCHAR     2
#define MAX_LEADBYTES       12  //  5 ranges, 2 bytes ea., 0 term.


#pragma pack(4)
typedef struct API_SMALL_RECT {
    Integer Left;
    Integer Top;
    Integer Right;
    Integer Bottom;
} API_SMALL_RECT;
typedef struct API_COORD {
    Integer X;
    Integer Y;
} API_COORD;
typedef struct API_CONSOLE_SCREEN_BUFFER_INFO {
    API_COORD dwSize;
    API_COORD dwCursorPosition;
    Integer wAttributes;
    API_SMALL_RECT srWindow;
    API_COORD dwMaximumWindowSize;
} API_CONSOLE_SCREEN_BUFFER_INFO;

typedef struct API_KEY_EVENT_RECORD {
    Long bKeyDown;
    Integer wRepeatCount;
    Integer wVirtualKeyCode;
    Integer wVirtualScanCode;
    Byte uChar;
    Long dwControlKeyState;
} API_KEY_EVENT_RECORD;
typedef struct API_MOUSE_EVENT_RECORD {
    API_COORD dwMousePosition;
    Long dwButtonState;
    Long dwControlKeyState;
    Long dwEventFlags;
} API_MOUSE_EVENT_RECORD;
typedef struct API_WINDOW_BUFFER_SIZE_RECORD {
    API_COORD dwSize;
} API_WINDOW_BUFFER_SIZE_RECORD;
typedef struct API_MENU_EVENT_RECORD {
    Long dwCommandId;
} API_MENU_EVENT_RECORD;
typedef struct API_FOCUS_EVENT_RECORD {
    Long bSetFocus;
} API_FOCUS_EVENT_RECORD;
typedef struct API_CONSOLE_INPUT_RECORD_EVENT {
    API_KEY_EVENT_RECORD KeyEvent;
    API_MOUSE_EVENT_RECORD MouseEvent;
    API_WINDOW_BUFFER_SIZE_RECORD WindowBufferSizeEvent;
    API_MENU_EVENT_RECORD MenuEvent;
    API_FOCUS_EVENT_RECORD FocusEvent;
} API_CONSOLE_INPUT_RECORD_EVENT;
typedef struct API_INPUT_RECORD {
    Byte EventType;
    API_CONSOLE_INPUT_RECORD_EVENT Event;
} API_INPUT_RECORD;

typedef struct API_CONSOLE_CURSOR_INFO {
    Long dwSize;
    Long bVisible;
} API_CONSOLE_CURSOR_INFO;

typedef struct API_CPINFO {
    Long MaxCharSize;                     //  max length (Byte) of a char
    Byte DefaultChar[MAX_DEFAULTCHAR];    //  default character
    Byte LeadByte[MAX_LEADBYTES];         //  lead byte ranges
} API_CPINFO;
typedef struct API_CHAR_INFO {
    Integer Char;
    Integer Attributes;
} API_CHAR_INFO;

#pragma pack()
[
    dllname("Kernel32.dll"),
    helpstring("Access to console API functions within the Kernel32.dll system file.")
]
module KernelConsole {
[entry("AllocConsole"), usesgetlasterror]
    long API_AllocConsole();
[entry("FreeConsole"), usesgetlasterror]
    long API_FreeConsole();
[entry("AttachConsole"), usesgetlasterror]
    MBOOL API_AttachConsole([in] long dwProcessId);
//========================================
[entry("SetConsoleTitleA"), usesgetlasterror]
    long API_SetConsoleTitle([in] String lpConsoleTitle);
[entry("SetConsoleTitleW"), usesgetlasterror]
    long API_SetConsoleTitleUnicode([in] String lpConsoleTitle);
    
[entry("GetConsoleTitleA"), usesgetlasterror]
    long API_GetConsoleTitle([in] String lpConsoleTitle, [in] long nSize);
[entry("GetConsoleTitleW"), usesgetlasterror]
    long API_GetConsoleTitleUnicode([in] String lpConsoleTitle, [in] long nSize);
    
[entry("SetConsoleCursorPosition"), usesgetlasterror]
    long API_SetConsoleCursorPosition([in] long hConsoleOutput, [in] long dwCursorPosition);
[entry("GetConsoleScreenBufferInfo"), usesgetlasterror]
    long API_GetConsoleScreenBufferInfo([in] long hConsoleOutput, [out] API_CONSOLE_SCREEN_BUFFER_INFO* lpConsoleScreenBufferInfo);
[entry("SetConsoleTextAttribute"), usesgetlasterror]
    long API_SetConsoleTextAttribute([in] long hConsoleOutput, [in] long wAttributes);
[entry("FillConsoleOutputCharacterA"), usesgetlasterror]
    long API_FillConsoleOutputCharacter([in] long hConsoleOutput, [in] Byte cCharacter, [in] long nLength, [in] long dwWriteCoord, [out] long* lpNumberOfCharsWritten);
[entry("FillConsoleOutputAttribute"), usesgetlasterror]
    long API_FillConsoleOutputAttribute([in] long hConsoleOutput, [in] long wAttribute, [in] long nLength, [in] long dwWriteCoord, [out] long* lpNumberOfAttrsWritten);
[entry("SetConsoleScreenBufferSize"), usesgetlasterror]
    long API_SetConsoleScreenBufferSize([in] long hConsoleOutput, [in] long dwSize);
[entry("SetConsoleCursorInfo"), usesgetlasterror]
    long API_SetConsoleCursorInfo([in] long hConsoleOutput, [out] API_CONSOLE_CURSOR_INFO* lpConsoleCursorInfo);
[entry("GetConsoleCursorInfo"), usesgetlasterror]
    long API_GetConsoleCursorInfo([in] long hConsoleOutput, [out] API_CONSOLE_CURSOR_INFO* lpConsoleCursorInfo);
[entry("SetConsoleWindowInfo"), usesgetlasterror]
    long API_SetConsoleWindowInfo([in] long hConsoleOutput, [in] long bAbsolute, [out] API_SMALL_RECT* lpConsoleWindow);
//========================================
[entry("WriteConsoleA"), usesgetlasterror]
    long API_WriteConsole([in] long hConsoleOutput, [in] Any lpBuffer, [in] long nNumberOfCharsToWrite, [out] long* lpNumberOfCharsWritten, [out] Any lpReserved);
[entry("WriteConsoleW"), usesgetlasterror]
    long API_WriteConsoleUnicode([in] long hConsoleOutput, [in] Any lpBuffer, [in] long nNumberOfCharsToWrite, [out] long* lpNumberOfCharsWritten, [out] Any lpReserved);

[entry("FlushConsoleInputBuffer"), usesgetlasterror]
    long API_FlushConsoleInputBuffer([in] long hConsoleInput);
[entry("ReadConsoleInputA"), usesgetlasterror]
    long API_ReadConsoleInput([in] long hConsoleInput, [out] API_INPUT_RECORD* lpBuffer, [in] long nLength, [out] long* lpNumberOfEventsRead);
//========================================
[entry("CloseConsoleHandle"), usesgetlasterror]
    long API_CloseConsoleHandle([in] long hConsoleHandle);
    
[entry("SetConsoleMode"), usesgetlasterror]
    long API_SetConsoleMode([in] long hConsoleHandle, [in] long dwMode);
[entry("SetConsoleCtrlHandler"), usesgetlasterror]
    long API_SetConsoleCtrlHandler([in] long HandlerRoutine, [in] long Add);
[entry("GetLargestConsoleWindowSize"), usesgetlasterror]
    API_COORD API_GetLargestConsoleWindowSize([in] long hConsoleOutput);
};

#endif //__KERNEL_CONSOLE_H__