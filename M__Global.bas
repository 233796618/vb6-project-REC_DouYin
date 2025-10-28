Attribute VB_Name = "M__Global"
Option Explicit

' 全局调度器实例
Public g_scheduler As TaskScheduler

' 全局直播间集合，使用 web_rid 作为 key
Public g_liverooms As Collection

' 全局直播间管理器
Public g_liveroomManager As LiveroomManager

' 全局关闭标志
Public g_appShuttingDown As Boolean

' 修改 InitializeGlobalVariables 过程
Public Sub InitializeGlobalVariables()
    ' 初始化关闭标志
    g_appShuttingDown = False
    
    Set g_scheduler = New TaskScheduler
    g_scheduler.Initialize
    
    Set g_liverooms = New Collection
    Set g_liveroomManager = New LiveroomManager
    g_liveroomManager.SetScheduler g_scheduler
    
    ' 设置全局实例
    SetLiveroomManagerInstance g_liveroomManager
    
    ' 启动 LiveroomManager 定时器
    StartLiveroomManagerTimer
    
    Debug.Print "全局变量初始化完成"
End Sub

' 修改 CleanupGlobalVariables 过程
Public Sub CleanupGlobalVariables()
    If Not g_scheduler Is Nothing Then
        g_scheduler.StopScheduler
        Set g_scheduler = Nothing
    End If
    
    If Not g_liverooms Is Nothing Then
        Set g_liverooms = Nothing
    End If
    
    If Not g_liveroomManager Is Nothing Then
        Set g_liveroomManager = Nothing
    End If
    
    ' 清理定时器
    CleanupLiveroomManagerTimer
    
    Debug.Print "全局变量清理完成"
End Sub

' 安全关闭应用程序
Public Sub SafeShutdownApplication()
    If g_appShuttingDown Then Exit Sub
    
    g_appShuttingDown = True
    Debug.Print "=== 应用程序开始安全关闭 ==="
    
    ' 1. 停止 LiveroomManager
    If Not g_liveroomManager Is Nothing Then
        g_liveroomManager.StopTimer
        Debug.Print ">>> LiveroomManager 已停止"
    End If
    
    ' 2. 停止调度器并开始关闭流程
    If Not g_scheduler Is Nothing Then
        g_scheduler.InitiateShutdown
        Debug.Print ">>> 调度器关闭流程已启动"
    End If
    
    ' 3. 等待所有任务完成（在定时器中处理）
    WaitForSafeExit
End Sub

' 等待安全退出条件
Public Sub WaitForSafeExit()
    On Error Resume Next
    
    Dim waitCount As Long
    waitCount = 0
    Dim maxWaitCount As Long
    maxWaitCount = 300 ' 最多等待30秒 (300 * 100ms)
    
    Debug.Print ">>> 等待安全退出条件..."
    
    Do While waitCount < maxWaitCount
        ' 检查是否可以安全退出
        If CanSafeExit() Then
            Debug.Print ">>> 达到安全退出条件，继续关闭流程"
            Exit Do
        End If
        
        ' 处理消息队列，保持UI响应
        DoEvents
        
        ' 等待100ms
        'Sleep 100
        waitCount = waitCount + 1
        
        If waitCount Mod 10 = 0 Then ' 每1秒输出一次状态
            Debug.Print ">>> 等待安全退出... (" & waitCount / 10 & "秒)"
        End If
    Loop
    
    If waitCount >= maxWaitCount Then
        Debug.Print ">>> 警告: 等待超时，强制退出"
    End If
    
    ' 继续最终清理
    CompleteShutdown
End Sub

' 检查是否可以安全退出
Public Function CanSafeExit() As Boolean
    On Error Resume Next
    
    If g_scheduler Is Nothing Then
        CanSafeExit = True
        Exit Function
    End If
    
    CanSafeExit = g_scheduler.CanSafeExit()
End Function

' 完成关闭流程
Public Sub CompleteShutdown()
    On Error Resume Next
    
    Debug.Print "=== 执行最终清理 ==="
    
    ' 清理全局变量
    CleanupGlobalVariables
    
    ' 清理所有定时器
    ForceStopAllTimers
    CleanupLiveroomManagerTimer
    
    ' 清理 CopyData 子类化
    CleanupAllCopyDataSubclass
    
    Debug.Print "=== 安全关闭完成，可以退出程序 ==="
End Sub

' 停止指定录制任务（全局函数）
Public Function StopRecordTask(ByVal web_rid As String) As Boolean
    On Error Resume Next
    
    If g_scheduler Is Nothing Then
        Debug.Print "错误: 调度器未初始化"
        StopRecordTask = False
        Exit Function
    End If
    
    StopRecordTask = g_scheduler.StopRecordTask(web_rid)
End Function

' 停止所有录制任务（全局函数）
Public Function StopAllRecordTasks() As Boolean
    On Error Resume Next
    
    If g_scheduler Is Nothing Then
        Debug.Print "错误: 调度器未初始化"
        StopAllRecordTasks = False
        Exit Function
    End If
    
    StopAllRecordTasks = g_scheduler.StopAllRecordTasks()
End Function

' 获取直播间统计信息
Public Function GetLiveroomStats() As String
    On Error Resume Next
    
    If g_liveroomManager Is Nothing Then
        GetLiveroomStats = "LiveroomManager 未初始化"
        Exit Function
    End If
    
    GetLiveroomStats = g_liveroomManager.GetDetailedStats()
End Function

' 获取直播间数量
Public Sub GetLiveroomCounts(ByRef total As Long, ByRef enabled As Long, ByRef disabled As Long, _
                            ByRef parsing As Long, ByRef recording As Long)
    On Error Resume Next
    
    total = 0
    enabled = 0
    disabled = 0
    parsing = 0
    recording = 0
    
    If Not g_liveroomManager Is Nothing Then
        g_liveroomManager.GetCountStats total, enabled, disabled, parsing, recording
    End If
End Sub

' ========== 工具函数 ==========

' 检查应用程序是否正在关闭
Public Function IsAppShuttingDown() As Boolean
    IsAppShuttingDown = g_appShuttingDown
End Function

' 设置关闭标志
Public Sub SetAppShuttingDown(ByVal shuttingDown As Boolean)
    g_appShuttingDown = shuttingDown
    If shuttingDown Then
        Debug.Print "应用程序关闭标志已设置"
    Else
        Debug.Print "应用程序关闭标志已清除"
    End If
End Sub

' 优化内存使用
Public Sub OptimizeMemoryUsage()
    On Error Resume Next
    ' 强制垃圾回收
    Dim i As Long
    For i = 1 To 3
        DoEvents
    Next i
    
    Debug.Print "内存优化完成"
End Sub

' 修改：简化添加直播间函数 - 修复参数顺序
Public Function AddLiveroomSimple(ByVal web_rid As String, _
                                  Optional ByVal AnchorName As String = "", _
                                  Optional ByVal checkIntervalMinutes As Long = 5, _
                                  Optional ByVal checkAllTime As Boolean = True, _
                                  Optional ByVal watchTimeStart As String = "", _
                                  Optional ByVal watchTimeEnd As String = "", _
                                  Optional ByVal taskPriority As taskPriority = tpMedium, _
                                  Optional ByVal enabled As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    If g_liveroomManager Is Nothing Then
        Debug.Print "错误: LiveroomManager 未初始化"
        AddLiveroomSimple = False
        Exit Function
    End If
    
    ' 直接调用 LiveroomManager 的方法，避免递归
    AddLiveroomSimple = g_liveroomManager.AddLiveroomSimple(web_rid, AnchorName, checkIntervalMinutes, checkAllTime, watchTimeStart, watchTimeEnd, taskPriority, enabled)
    
    Exit Function
    
ErrorHandler:
    Debug.Print "AddLiveroomSimple 错误: " & Err.Description & " (Web_rid: " & web_rid & ")"
    AddLiveroomSimple = False
End Function

' ========== 全局辅助函数 ==========

Public Function GetPriorityText(ByVal priority As taskPriority) As String
    Select Case priority
        Case tpHigh: GetPriorityText = "高"
        Case tpMedium: GetPriorityText = "中"
        Case tpLow: GetPriorityText = "低"
        Case Else: GetPriorityText = "未知"
    End Select
End Function

Public Function GetWatchTimeDisplay(ByRef room As LiveRoom) As String
    If room.checkAllTime Then
        GetWatchTimeDisplay = "全天"
    ElseIf room.watchTimeStart <> "" And room.watchTimeEnd <> "" Then
        GetWatchTimeDisplay = room.watchTimeStart & "-" & room.watchTimeEnd
    Else
        GetWatchTimeDisplay = "未设置"
    End If
End Function


' 添加这些辅助函数到 M__Global.bas
Public Sub LogMessage(ByVal message As String)
    Debug.Print ">>> " & Format(Now, "hh:mm:ss") & " - " & message
End Sub

Public Function GetSelectedWebRidFromGrid() As String
    On Error Resume Next
    If frmMain.MSHFlexGrid.row < 1 Then
        GetSelectedWebRidFromGrid = ""
    Else
        GetSelectedWebRidFromGrid = frmMain.MSHFlexGrid.TextMatrix(frmMain.MSHFlexGrid.row, 1)
    End If
End Function


' 新增：获取任务状态显示文本
Public Function GetTaskStatusDisplay(ByRef room As LiveRoom) As String
    Dim statusText As String
    
    ' 录制状态优先显示
    If room.isRecording Then
        If room.RecordTaskStatus = "停止中" Then
            statusText = "录制停止中"
        Else
            statusText = "录制中"
        End If
    ElseIf room.RecordTaskStatus = "排队中" Then
        statusText = "录制排队中"
    ElseIf room.RecordTaskStatus = "运行中" Then
        statusText = "录制运行中"
    ElseIf room.ParseTaskStatus = "运行中" Then
        statusText = "解析中"
    ElseIf room.ParseTaskStatus = "排队中" Then
        statusText = "解析排队中"
    ElseIf room.roomStatus = "直播中" Then
        statusText = "直播中"
    Else
        statusText = room.roomStatus
    End If
    
    GetTaskStatusDisplay = statusText
End Function

' 修改：计算行背景颜色 - 添加停止中状态
Public Function CalculateRowBackgroundColor(ByRef room As LiveRoom) As Long
    If room.roomStatus = "直播中" And room.isRecording Then
        If room.RecordTaskStatus = "停止中" Then
            CalculateRowBackgroundColor = &HCC99FF ' 停止中状态颜色
        Else
            CalculateRowBackgroundColor = &HCCCCFF ' 录制中颜色
        End If
    ElseIf room.roomStatus = "直播中" Then
        CalculateRowBackgroundColor = &HCCFFCC ' 直播中颜色
    ElseIf room.ParseTaskStatus = "运行中" Then
        CalculateRowBackgroundColor = &HFFCCFF ' 解析中颜色
    ElseIf room.ParseTaskStatus = "排队中" Then
        CalculateRowBackgroundColor = &HFFE6CC ' 解析排队中颜色
    ElseIf Not room.enabled Then
        CalculateRowBackgroundColor = &HE0E0E0 ' 禁用状态颜色
    ElseIf room.roomStatus = "离线" Then
        CalculateRowBackgroundColor = &HF0F0F0 ' 离线状态颜色
    Else
        CalculateRowBackgroundColor = vbWhite ' 默认颜色
    End If
End Function
