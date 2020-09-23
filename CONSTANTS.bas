Attribute VB_Name = "CONSTANTS"

Public Const ENABLE_ECHO_INPUT = &H4
Public Const ENABLE_LINE_INPUT = &H2

Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&

Public Type COORD

    X As Integer
    Y As Integer

End Type

Public Type SMALL_RECT

    Left    As Integer
    Top     As Integer
    Right   As Integer
    Bottom  As Integer

End Type

Public Type CONSOLE_SCREEN_BUFFER_INFO

    dwSize              As COORD
    dwCursorPosition    As COORD
    wAttributes         As Integer
    srWindow            As SMALL_RECT
    dwMaximumWindowSize As COORD

End Type

Public Type CONSOLE_CURSOR_INFO
        dwSize As Long
        bVisible As Long
End Type


Public Declare Function AllocConsole Lib "kernel32" () As Long
Public Declare Function FreeConsole Lib "kernel32" () As Long
Public Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Public Declare Function ReadFileNULL Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Long) As Long
Public Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, ByVal dwMode As Long) As Long
Public Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Public Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, dwCursorPosition As COORD) As Long
Public Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Public Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Public Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Public Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
Public Declare Function SetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Public Declare Function GetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Public Declare Function SetConsoleWindowInfo Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal bAbsolute As Boolean, lpConsoleWindow As SMALL_RECT) As Long
Public Declare Function GetLargestConsoleWindowSize Lib "kernel32" (ByVal hConsoleOutput As Long) As COORD
Public Declare Function SetConsoleScreenBufferSize Lib "kernel32" (ByVal hConsoleOutput As Long, dwSize As COORD) As Long
Public Declare Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As Long, ByVal Add As Boolean) As Long
Public Declare Function FlushConsoleInputBuffer Lib "kernel32" (ByVal hConsoleInput As Long) As Long


Public Const CTRL_BREAK_EVENT = 1
Public Const CTRL_C_EVENT = 0
Public Const CTRL_CLOSE_EVENT = 2
Public Const CTRL_LOGOFF_EVENT = 5
Public Const CTRL_SHUTDOWN_EVENT = 6







Public Function Disable(CritEvent As Long) As Boolean
'Callback API !
'If return is True, then the default critical event handlers are not called
'I was thinking of adding code to notify the end user
'about CTRL+BREAK or other critical events , any standard
'VB keyword used in this sub (Set =, For x= etc) causes VB to crash !
    Disable = True
End Function
