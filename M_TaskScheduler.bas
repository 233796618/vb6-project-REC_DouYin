Attribute VB_Name = "M_TaskScheduler"
Option Explicit


' ListView 消息常量
Public Const LVM_FIRST As Long = &H1000
Public Const LVM_HITTEST As Long = (LVM_FIRST + 18)
Public Const WM_SETREDRAW As Long = &HB
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Public Const LVM_GETITEMCOUNT As Long = (LVM_FIRST + 4)
Public Const LVM_SETITEMSTATE As Long = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETTOPINDEX As Long = (LVM_FIRST + 39)
Public Const LVM_ENSUREVISIBLE As Long = (LVM_FIRST + 19)

Public Const LVS_EX_DOUBLEBUFFER As Long = &H10000
Public Const LVS_EX_FULLROWSELECT As Long = &H20
Public Const LVS_EX_TRACKSELECT As Long = &H8

' ListView 选择状态相关常量
Public Const LVIS_SELECTED As Long = &H2
Public Const LVIS_FOCUSED As Long = &H1

' 鼠标点击测试结构
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type LVHITTESTINFO
    pt As POINTAPI
    flags As Long
    iItem As Long
    iSubItem As Long
End Type

Public Type LVITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' API 声明
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

' 进程创建相关API
Public Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
'Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
'Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long





'Public Type PROCESS_INFORMATION
'    hProcess As Long
'    hThread As Long
'    dwProcessId As Long
'    dwThreadId As Long
'End Type

' 常量
Public Const SCHEDULER_TIMER_INTERVAL As Long = 200    ' 调度器定时器间隔

' 进程状态常量
'Public Const STILL_ACTIVE As Long = &H103
'Public Const NORMAL_PRIORITY_CLASS As Long = &H20
'Public Const CREATE_NO_WINDOW As Long = &H8000000

' 管道相关常量
Public Const STARTF_USESTDHANDLES As Long = &H100
Public Const STARTF_USESHOWWINDOW As Long = &H1
Public Const SW_HIDE As Long = 0

' Shell 函数常量
Public Const vbHide As Integer = 0
Public Const vbNormalFocus As Integer = 1
Public Const vbMinimizedFocus As Integer = 2
Public Const vbMaximizedFocus As Integer = 3
Public Const vbNormalNoFocus As Integer = 4
Public Const vbMinimizedNoFocus As Integer = 6

' 定时器相关变量
Public g_schedulerTimerID As Long
Public g_schedulerTimerRunning As Boolean

' 添加高精度计时 API
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

' 添加高精度时间获取函数
Public Function GetHighResolutionTime() As Currency
    Dim count As Currency
    QueryPerformanceCounter count
    GetHighResolutionTime = count
End Function

Public Function HighResTimeToMs(ByVal highResTime As Currency) As Currency
    Dim freq As Currency
    QueryPerformanceFrequency freq
    HighResTimeToMs = highResTime / freq * 1000
End Function


' 定时器回调函数 - 重命名
Public Sub SchedulerTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Debug.Print ">>> 调度器定时器触发: " & Format(Now, "hh:mm:ss")
    
    If Not g_scheduler Is Nothing Then
        g_scheduler.ProcessQueuesByTimer
    Else
        Debug.Print ">>> 错误: 调度器为空，停止定时器"
        StopSchedulerTimer
    End If
End Sub

' 启动调度器定时器
Public Sub StartSchedulerTimer()
    If g_schedulerTimerID = 0 Then
        g_schedulerTimerID = SetTimer(0, 0, SCHEDULER_TIMER_INTERVAL, AddressOf SchedulerTimerProc)
        If g_schedulerTimerID <> 0 Then
            g_schedulerTimerRunning = True
            Debug.Print ">>> 调度器定时器启动成功 (ID: " & g_schedulerTimerID & ")"
        Else
            Debug.Print ">>> 错误：无法启动调度器定时器"
        End If
    Else
        Debug.Print ">>> 调度器定时器已在运行 (ID: " & g_schedulerTimerID & ")"
    End If
End Sub

' 停止调度器定时器
Public Sub StopSchedulerTimer()
    If g_schedulerTimerID <> 0 Then
        KillTimer 0, g_schedulerTimerID
        g_schedulerTimerID = 0
        g_schedulerTimerRunning = False
        Debug.Print ">>> 调度器定时器已停止"
    Else
        Debug.Print ">>> 调度器定时器未运行"
    End If
End Sub

' 检查调度器定时器是否在运行
Public Function IsSchedulerTimerRunning() As Boolean
    IsSchedulerTimerRunning = (g_schedulerTimerID <> 0)
End Function

Public Sub InitializeScheduler()
    Debug.Print ">>> 初始化调度器"
    Debug.Print ">>> 调度器初始化完成"
End Sub

Public Sub CleanupScheduler()
    Debug.Print ">>> 开始清理调度器"
    StopSchedulerTimer
    Debug.Print ">>> 调度器清理完成"
End Sub

' 强制停止所有定时器
Public Sub ForceStopAllTimers()
    Debug.Print ">>> 强制停止所有定时器"
    StopSchedulerTimer
    Debug.Print ">>> 定时器强制停止完成"
End Sub
