VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CONSOLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum ForeColor
     FOREGROUND_GREEN = &H2     '  text color contains green.
     FOREGROUND_BLUE = &H1     '  text color contains blue.
     FOREGROUND_RED = &H4     '  text color contains red.
     FOREGROUND_INTENSITY = &H8     '  text color is intensified.
End Enum

Public Enum BackColor
     BACKGROUND_BLUE = &H10    '  background color contains blue.
     BACKGROUND_GREEN = &H20    '  background color contains green.
     BACKGROUND_INTENSITY = &H80    '  background color is intensified.
     BACKGROUND_RED = &H40    '  background color contains red.
End Enum


Dim hConsOutput As Long, hConsInput As Long, Cnt As Long
Dim ContDisabledFlag As Boolean
'NOTE: IN ALL THE SUBS/FUNTIONS THE PAREMETER "ERROR" IS AN INPUTVARIABLE
'THAT WILL HOLD ANY ERRORS RETURNED BY THE ACTIVEX DLL



Sub CreateConsole(Optional Error As String)
    Dim res As Long
    'Check whether console is already created .
    'A process can have only one console !
    If Cnt = 1 Then
        Error = "Cannot allocate two consoles"
        Exit Sub
    End If
    If Cnt = 0 Then
        Cnt = Cnt + 1
    End If
        res = AllocConsole
    If res = 0 Then
        Error = "Failed to initialize console"
        Exit Sub
    End If
    ' Get input and output handles
    hConsOutput = GetStdHandle(STD_OUTPUT_HANDLE)
    hConsInput = GetStdHandle(STD_INPUT_HANDLE)
    ' Enable input "echo"
    SetConsoleMode hConsInput, ENABLE_ECHO_INPUT
    Error = ""
End Sub

Sub OutputConsole(ByVal Text As String, Optional NewLineFlag As Boolean = True)
    If NewLineFlag = True Then
        Text = Text + Chr$(10) + Chr$(13)
    End If
    Dim A As Long
    Dim L As Long
    
    If Len(Text) < 1 Then Exit Sub
    
    'Convert unicode to ansi.
    ReDim ByteBuffer(0 To Len(Text)) As Byte
    
    For A = 1 To Len(Text)
        ByteBuffer(A) = Asc(Mid$(Text, A, 1))
    Next A

    WriteConsole hConsOutput, ByVal VarPtr(ByteBuffer(1)), UBound(ByteBuffer), L, ByVal 0&

End Sub

Function InputConsole() As String

    Dim T As String
    Dim B As String

    Do

        T = Inkey()

        If T > "" Then

            Select Case Asc(T)
            Case 10
            Case vbKeyBack: If B > "" Then B = Left$(B, Len(B) - 1)
            Case vbKeyReturn: Exit Do
            Case Else:  B = B & T
            End Select

        End If

    Loop

    InputConsole = B

End Function


Private Function Inkey() As String
    Dim T As Byte
    Dim L As Long
'Apparently an undocumented API. Come to know about it from
'the internet
    ReadFileNULL hConsInput, T, Len(T), L, ByVal 0&

    If L Then Inkey = Chr$(T)
End Function

Sub GetCursorPosition(X As Long, Y As Long, Optional Error As String)
'Don't know whether this works properly !
    Dim lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO
    Dim A As Long
    A = GetConsoleScreenBufferInfo(hConsOutput, lpConsoleScreenBufferInfo)
    If A = 0 Then
        Error = "Unable to get cursor info"
        Exit Sub
    End If
    X = lpConsoleScreenBufferInfo.dwCursorPosition.X
    Y = lpConsoleScreenBufferInfo.dwCursorPosition.Y
End Sub


Sub Locate(ByVal Row As Long, ByVal Column As Long)
'Don't know whether this works properly !
    Dim dwCursorPosition As COORD
    dwCursorPosition.X = Row
    dwCursorPosition.Y = Column
    SetConsoleCursorPosition hConsOutput, dwCursorPosition
End Sub

Sub SetCTitle(Title As String, Optional Error As String)
    Dim A As Long
'Call standard APIs for setting console title !
    
    A = SetConsoleTitle(Title)
    If A = 0 Then
        Error = "Unable to set console title"
        Exit Sub
    End If
End Sub


Sub SetColor(Textcolor As ForeColor, BackColor As BackColor, Optional Error As String)
'Set Console Title !
    Dim A As Long, CColor As Long
    Dim Text As Long, Back As Long
    Text = Textcolor
    Back = BackColor
    CColor = Text Or Back
    A = SetConsoleTextAttribute(hConsOutput, CColor)
    If A = 0 Then
        Error = "Unable to set color attributes"
        Exit Sub
    End If
End Sub

Sub CloseConsole(Optional Error As String)
Dim A As Long
'Call API to close console
    A = FreeConsole
    If A = 0 Then
        Error = "Unable to close console"
        Exit Sub
    End If
    Cnt = 0
End Sub

Sub GetCTitle(Title As String, Optional Error As String)
    Dim Size As Long, T As String
    Size = 255
    T = String(255, " ")
'   Get console title
    s = GetConsoleTitle(T, Size)
    If s = 0 Then
        Error = "Unable to retrieve console title"
        Exit Sub
    End If
    Title = Trim(Left$(T, s))
End Sub
Sub GetCursor(CurSize As Long, VisibilityFlag As Boolean, Optional Error As String)
Dim Info As CONSOLE_CURSOR_INFO, A As Long
A = GetConsoleCursorInfo(hConsOutput, Info)
If A = 0 Then
    Error = "Unable to get cursor attributes"
    Exit Sub
End If
CurSize = Info.dwSize
VisibilityFlag = Info.bVisible
End Sub

Sub SetCursor(CurSize As Long, VisibilityFlag As Boolean, Optional Error As String)
Dim Info As CONSOLE_CURSOR_INFO, A As Long
Info.bVisible = VisibilityFlag
Info.dwSize = CurSize
A = SetConsoleCursorInfo(hConsOutput, Info)
If A = 0 Then
    Error = "Unable to set cursor attributes"
    Exit Sub
End If
End Sub

Sub Newline()
    OutputConsole "", True
End Sub

Private Sub Class_Terminate()
    If Cnt <> 0 Then FreeConsole
End Sub

Sub DisableCriticalEvents(Optional Error As String)
'Disable default CTRL+BREAK handlers !
    res = SetConsoleCtrlHandler(AddressOf Disable, True)
    
    If res = 0 Then
        Error = "Cannot set critical event trap"
    End If
    ContDisabledFlag = True
End Sub

Sub EnableCriticalEvents(Optional Error As String, Optional ProcessPrevEvents As Boolean = False)
'Reenable CTRL+BREAK and other default critical event handlers
    If ContDisabledFlag = True Then
    If ProcessPrevEvents <> True Then
        res = FlushConsoleInputBuffer(hConsInput)
        If res = 0 Then
            Error = "Unable to clear previous events"
            Exit Sub
        End If
    End If
        res = SetConsoleCtrlHandler(AddressOf Disable, False)
        If res = 0 Then
            Error = "Cannot disable critical event trap"
            Exit Sub
        End If
        ContDisabledFlag = False
    Else
        Error = "Critical event trap not set"
    End If
End Sub

Sub ClearInputBuffer(Optional Error As String)
        res = FlushConsoleInputBuffer(hConsInput)
        If res = 0 Then
            Error = "Unable to clear the input buffer"
            Exit Sub
        End If
End Sub

Property Get OutputHandle() As Long
    OutputHandle = hConsOutput
End Property

Property Get InputHandle() As Long
    GetInputHandle = hConsInput
End Property


