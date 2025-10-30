' M__Global.bas （精简/重构版）
' 说明：
'   - 统一全局实例管理：Scheduler、LiveroomManager、全局开关等
'   - 提供初始化/清理/优雅关闭流程（SafeShutdownApplication）
'   - 保持最小全局变量，鼓励后续通过参数传入依赖（依赖注入风格）
Option Explicit

Public Const LiveRoomManager_TIMER_INTERVAL As Long = 300
Public Const SCHEDULER_TIMER_INTERVAL As Long = 200

Public g_scheduler As TaskScheduler
Public g_liveroomManager As LiveroomManager
Public g_liverooms As Collection

Public g_appShuttingDown As Boolean

Public Sub InitializeGlobalVariables()
    On Error GoTo EH
    g_appShuttingDown = False

    If g_scheduler Is Nothing Then
        Set g_scheduler = New TaskScheduler
        g_scheduler.Initialize
    End If

    If g_liveroomManager Is Nothing Then
        Set g_liveroomManager = New LiveroomManager
        g_liveroomManager.SetScheduler g_scheduler
    End If

    If g_liverooms Is Nothing Then
        Set g_liverooms = New Collection
    End If

    SetLiveroomManagerInstance g_liveroomManager
    StartLiveroomManagerTimer

    Debug.Print "全局初始化完成：Scheduler 和 LiveroomManager 已就绪"
    Exit Sub
EH:
    Debug.Print "InitializeGlobalVariables 错误: " & Err.Description
    Err.Clear
End Sub

Public Sub CleanupGlobalVariables()
    On Error Resume Next
    Debug.Print "开始全局清理..."

    If Not g_scheduler Is Nothing Then
        g_scheduler.InitiateShutdown
        Set g_scheduler = Nothing
    End If

    If Not g_liveroomManager Is Nothing Then
        g_liveroomManager.StopTimer
        Set g_liveroomManager = Nothing
    End If

    If Not g_liverooms Is Nothing Then
        Set g_liverooms = Nothing
    End If

    CleanupLiveroomManagerTimer
    CleanupAllCopyDataSubclass

    Debug.Print "全局清理完成"
End Sub

Public Sub SafeShutdownApplication()
    On Error Resume Next
    If g_appShuttingDown Then Exit Sub
    g_appShuttingDown = True

    Debug.Print "开始优雅关闭应用..."

    If Not g_liveroomManager Is Nothing Then
        g_liveroomManager.StopTimer
    End If

    If Not g_scheduler Is Nothing Then
        g_scheduler.InitiateShutdown
    End If

    WaitForSafeExit
End Sub

Public Sub WaitForSafeExit()
    On Error Resume Next
    Dim waitCount As Long: waitCount = 0
    Dim maxWaitCount As Long: maxWaitCount = 300

    Debug.Print "等待安全退出中..."
    Do While waitCount < maxWaitCount
        If CanSafeExit() Then
            Debug.Print "达到安全退出条件"
            Exit Do
        End If
        DoEvents
        Sleep 100
        waitCount = waitCount + 1
        If waitCount Mod 10 = 0 Then Debug.Print "等待安全退出... " & (waitCount / 10) & "s"
    Loop

    If waitCount >= maxWaitCount Then
        Debug.Print "等待超时，将强制清理"
    End If

    CompleteShutdown
End Sub

Public Function CanSafeExit() As Boolean
    On Error Resume Next
    If g_scheduler Is Nothing Then
        CanSafeExit = True
    Else
        CanSafeExit = g_scheduler.CanSafeExit()
    End If
End Function

Public Sub CompleteShutdown()
    On Error Resume Next
    Debug.Print "CompleteShutdown: 开始释放全局资源..."
    CleanupGlobalVariables
    ForceStopAllTimers
    CleanupLiveroomManagerTimer
    CleanupAllCopyDataSubclass
    Debug.Print "CompleteShutdown: 释放完成，程序可以退出"
End Sub

Public Sub StartSchedulerTimer()
    On Error Resume Next
    StartSchedulerTimer
End Sub

Public Sub StopSchedulerTimer()
    On Error Resume Next
    StopSchedulerTimer
End Sub

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' End of M__Global.bas

