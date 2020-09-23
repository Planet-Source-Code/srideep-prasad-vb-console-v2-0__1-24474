Attribute VB_Name = "CONSOLE"
'Color constants
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Sub Main()
    Dim C As VBCONSOLE.CONSOLE
    
    Set C = CreateObject("VBCONSOLE.CONSOLE")
    'Use the methods of the Console1 object to create and
    'manage a console
    Dim Error As String, N As String
    C.CreateConsole Error
    If Error <> "" Then
        MsgBox Error
        Exit Sub
    End If
    C.SetCTitle "VBConsole Demo", Error
    If Error <> "" Then
        MsgBox Error
        Exit Sub
    End If
    
    C.OutputConsole "Thank you for using VBConsole"
    C.OutputConsole "Please enter your name:", False
    N$ = C.InputConsole
    C.SetColor FOREGROUND_BLUE Or FOREGROUND_INTENSITY, BACKGROUND_BLUE Or BACKGROUND_GREEN Or BACKGROUND_RED Or BACKGROUND_INTENSITY
    C.OutputConsole "Welcome " + N$
    C.SetCTitle "Welcome " + N$ + " - VBConsole Demo"
    C.Newline
    C.OutputConsole "VBConsole provides a VB Programmer a flexible console interface"
    C.OutputConsole "To learn more read the Readme.txt file"
    C.OutputConsole ""
    C.SetColor FOREGROUND_GREEN Or FOREGROUND_INTENSITY, BACKGROUND_BLUE Or BACKGROUND_INTENSITY
    C.OutputConsole "Press the ENTER key to continue"
    N$ = C.InputConsole
    C.CloseConsole
End Sub
