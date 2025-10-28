Attribute VB_Name = "M_WinAPI"
' WinAPI.bas
Option Explicit


' 进程相关 API
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
    (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Public Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

' 管道相关 API
Public Declare Function CreatePipe Lib "kernel32" _
    (phReadPipe As Long, phWritePipe As Long, _
    lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long

Public Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long

Public Declare Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, ByVal lpBuffer As Long, _
    ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long


Public Declare Function PeekNamedPipe Lib "kernel32" _
    (ByVal hNamedPipe As Long, lpBuffer As Any, _
    ByVal nBufferSize As Long, lpBytesRead As Long, _
    lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long

' 高精度计时器 API
Public Declare Function QueryPerformanceCounter Lib "kernel32" _
    (lpPerformanceCount As Currency) As Long

Public Declare Function QueryPerformanceFrequency Lib "kernel32" _
    (lpFrequency As Currency) As Long

' 定时器 API
Public Declare Function SetTimer Lib "user32" _
    (ByVal hWnd As Long, ByVal nIDEvent As Long, _
    ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Public Declare Function KillTimer Lib "user32" _
    (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

' 其他 API
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long
Public Declare Function VarPtr Lib "msvbvm60.dll" (lp As Any) As Long

' 常量定义
Public Const STILL_ACTIVE As Long = &H103
Public Const INFINITE As Long = &HFFFFFFFF
Public Const NORMAL_PRIORITY_CLASS As Long = &H20
Public Const CREATE_NO_WINDOW As Long = &H8000000

' 结构体定义
Public Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

