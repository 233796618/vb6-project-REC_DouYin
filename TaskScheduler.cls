VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 成员变量
Private m_queueManager As TaskSchedulerQueueManager
Private WithEvents m_taskManager As TaskSchedulerTaskManager
Attribute m_taskManager.VB_VarHelpID = -1

' 关闭管理
Private m_shutdownInProgress As Boolean
Private m_shutdownInitiated As Boolean

' 事件定义
Public Event TaskAdded(ByRef task As task)
Public Event TaskStarted(ByRef task As task)
Public Event TaskCompleted(ByRef task As task)
Public Event TaskFailed(ByRef task As task, ByVal errorMsg As String)
Public Event TaskTimeout(ByRef task As task)
Public Event TaskOutput(ByRef task As task, ByVal outputData As String, ByVal isErrorOutput As Boolean)

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    ' 先初始化所有对象
    Set m_queueManager = New TaskSchedulerQueueManager
    Set m_taskManager = New TaskSchedulerTaskManager
    
    ' 然后设置父引用
    m_taskManager.SetParent Me
    
    m_shutdownInProgress = False
    m_shutdownInitiated = False
    
    Debug.Print "任务调度器初始化完成"
    Exit Sub
    
ErrorHandler:
    Debug.Print ">>> 任务调度器初始化出错: " & Err.Description
    ' 尝试重新初始化
    Set m_queueManager = New TaskSchedulerQueueManager
    Set m_taskManager = New TaskSchedulerTaskManager
    m_taskManager.SetParent Me
End Sub

Private Sub Class_Terminate()
    Cleanup
    Debug.Print "任务调度器已清理"
End Sub

' ========== 初始化方法 ==========

Public Sub Initialize()
    On Error GoTo ErrorHandler
    
    Debug.Print ">>> 开始初始化任务调度器..."
    
    ' 检查对象是否已初始化
    If m_queueManager Is Nothing Then
        Set m_queueManager = New TaskSchedulerQueueManager
        Debug.Print ">>> 重新初始化队列管理器"
    End If
    
    If m_taskManager Is Nothing Then
        Set m_taskManager = New TaskSchedulerTaskManager
        m_taskManager.SetParent Me
        Debug.Print ">>> 重新初始化任务管理器"
    End If
    
    ' 添加任务类型配置
    Debug.Print ">>> 添加任务类型配置..."
    m_queueManager.AddTaskType "parser", 1, 3000, 10
    m_queueManager.AddTaskType "recorder", 20, 500, 1800
    m_queueManager.AddTaskType "default", 3, 1000, 300
    
    Debug.Print ">>> 任务队列初始化完成"
    Exit Sub
    
ErrorHandler:
    Debug.Print ">>> 初始化任务调度器时出错: " & Err.Description
    Debug.Print ">>> 错误来源: " & Err.Source
    Debug.Print ">>> 错误号: " & Err.Number
    
    ' 尝试重新初始化
    If Err.Number = 91 Then ' 对象变量未设置
        Debug.Print ">>> 检测到对象未设置错误，尝试重新初始化..."
        Set m_queueManager = New TaskSchedulerQueueManager
        Set m_taskManager = New TaskSchedulerTaskManager
        m_taskManager.SetParent Me
        
        ' 重试初始化
        Resume
    End If
End Sub

' ========== 任务管理方法 ==========

Public Function AddTask(ByVal taskType As String, _
                       ByVal commandLine As String, _
                       Optional ByVal parameters As String = "", _
                       Optional ByVal taskName As String = "", _
                       Optional ByVal priority As taskPriority = tpMedium, _
                       Optional ByVal timeoutSeconds As Long = -1, _
                       Optional ByVal monitorPipe As Boolean = False, _
                       Optional ByVal web_rid As String = "", _
                       Optional ByVal workingDirectory As String = "") As Long
    
    ' 在关闭期间拒绝新任务
    If m_shutdownInProgress Then
        Debug.Print ">>> 关闭进行中，拒绝新任务: " & taskName
        AddTask = -1
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    ' 检查调度器是否已初始化
    If Not IsInitialized() Then
        Reinitialize
    End If
    
    Dim task As task
    Set task = New task
    
    ' 设置任务属性
    task.taskType = taskType
    task.commandLine = commandLine
    task.parameters = parameters
    task.priority = priority
    task.monitorPipe = monitorPipe
    task.web_rid = web_rid
    task.workingDirectory = workingDirectory
    
    If timeoutSeconds >= 0 Then
        task.timeoutSeconds = timeoutSeconds
    Else
        task.timeoutSeconds = m_queueManager.GetDefaultTimeout(taskType)
    End If
    
    ' 设置任务名称
    If taskName = "" Then
        task.taskName = taskType & "_" & task.taskID
    Else
        task.taskName = taskName
    End If
    
    ' 获取队列并添加任务
    Dim queue As TaskQueue
    Set queue = m_queueManager.GetQueue(taskType, priority)
    If queue Is Nothing Then
        Set queue = m_queueManager.GetDefaultQueue()
    End If
    
    queue.AddTask task
    AddTask = task.taskID
    
    Debug.Print ">>> 任务已添加: " & task.taskName & " (ID: " & task.taskID & ") 到队列: " & queue.queueName
    
    RaiseEvent TaskAdded(task)
    
    ProcessQueuesImmediately
    
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> 添加任务时出错: " & Err.Description
    AddTask = -1
End Function

' ========== 队列处理方法 ==========

Public Sub ProcessQueuesByTimer()
    On Error GoTo ErrorHandler
    
    Debug.Print ">>> 定时器触发队列处理 - 开始 " & Format(Now, "hh:mm:ss")
    
    ' 检查调度器是否已初始化
    If Not IsInitialized() Then
        Debug.Print ">>> 调度器未初始化，跳过处理"
        Exit Sub
    End If
    
    ' 检查完成进程和管道输出
    m_taskManager.CheckCompletedProcesses m_queueManager
    m_taskManager.CheckPipeOutputs m_queueManager
    m_taskManager.CheckTaskTimeouts m_queueManager
    
    ' 在关闭期间不处理新任务
    If Not m_shutdownInProgress Then
        ProcessQueuesByPriority
    Else
        Debug.Print ">>> 关闭进行中，跳过新任务处理"
        ' 检查是否可以安全退出
        If CanSafeExit() Then
            Debug.Print ">>> 达到安全退出条件"
            If IsSchedulerTimerRunning() Then
                StopSchedulerTimer
                Debug.Print ">>> 关闭监控完成，停止定时器"
            End If
        End If
    End If
    
    m_taskManager.CleanCompletedRunners
    
    ' 管理定时器状态
    ManageTimerByRunningState
    
    Debug.Print ">>> 定时器触发队列处理 - 结束"
    Exit Sub
    
ErrorHandler:
    Debug.Print ">>> 定时器处理出错: " & Err.Description
End Sub

' 按优先级处理队列
Private Sub ProcessQueuesByPriority()
    Dim priority As taskPriority
    Dim queue As TaskQueue
    Dim task As task
    Dim runner As Object
    Dim executedAnyTask As Boolean
    Dim totalExecuted As Long
    
    executedAnyTask = False
    totalExecuted = 0
    
    Debug.Print ">>> 开始按优先级处理队列 - " & Format(Now, "hh:mm:ss")
    
    ' 先检查完成进程和管道输出
    m_taskManager.CheckCompletedProcesses m_queueManager
    m_taskManager.CheckPipeOutputs m_queueManager
    m_taskManager.CheckTaskTimeouts m_queueManager
    
    ' 按优先级从高到低处理
    For priority = tpHigh To tpLow
        For Each queue In m_queueManager.GetAllQueues()
            If queue.priority = priority And queue.PendingCount > 0 Then
                'Debug.Print ">>> 检查队列: " & queue.queueName & _
                '           " (优先级:" & GetPriorityName(priority) & _
                '           " 等待:" & queue.PendingCount & _
                '           " 运行中:" & queue.RunningCount & ")"
                
                ' 详细检查队列是否可以运行任务
                Dim CanRun As Boolean
                Dim Reason As String
                m_queueManager.CheckQueueRunTaskDetailed queue, CanRun, Reason
                
                If CanRun Then
                    ' 处理当前队列中的任务
                    Do While queue.PendingCount > 0
                        ' 再次检查是否可以运行任务
                        m_queueManager.CheckQueueRunTaskDetailed queue, CanRun, Reason
                        If Not CanRun Then
                            Debug.Print ">>> 队列 '" & queue.queueName & "' 无法运行更多任务: " & Reason
                            Exit Do
                        End If
                        
                        ' 获取下一个任务
                        Set task = queue.GetNextTask()
                        If task Is Nothing Then
                            Debug.Print ">>> 无法从队列获取任务，跳过"
                            Exit Do
                        End If
                        
                        ' 创建运行器并执行任务
                        If ExecuteTask(queue, task) Then
                            executedAnyTask = True
                            totalExecuted = totalExecuted + 1
                        Else
                            ' 将任务重新放回队列前端
                            ReAddTaskToFront queue, task
                            Exit Do
                        End If
                        
                        ' 检查是否达到限制
                        m_queueManager.CheckQueueRunTaskDetailed queue, CanRun, Reason
                        If Not CanRun Then
                            'Debug.Print ">>> 达到限制，停止处理当前队列: " & Reason
                            Exit Do
                        End If
                    Loop
                Else
                    'Debug.Print ">>> 跳过队列 '" & queue.queueName & "' - " & Reason
                End If
            End If
        Next queue
    Next priority
    
    ' 清理完成的运行器
    m_taskManager.CleanCompletedRunners
    
    ' 管理定时器状态
    ManageTimerByRunningState
    
    ' 输出处理结果
    If totalExecuted > 0 Then
        'Debug.Print ">>> 本轮处理完成 - 启动了 " & totalExecuted & " 个任务"
    Else
        'Debug.Print ">>> 本轮处理完成 - 没有启动新任务"
        m_queueManager.LogDetailedWaitingReasons
    End If
    
    ' 记录调度器状态
    m_queueManager.LogSchedulerStatus
End Sub

' 执行任务
' 执行任务
Private Function ExecuteTask(ByRef queue As TaskQueue, ByRef task As task) As Boolean
    On Error GoTo ErrorHandler
    
    ' 创建运行器
    Dim runner As Object
    Set runner = m_taskManager.CreateRunner(task.monitorPipe)
    If runner Is Nothing Then
        Debug.Print ">>> 无法创建运行器: " & task.taskName
        ExecuteTask = False
        Exit Function
    End If
    
    ' 设置任务和运行器
    Set runner.CurrentTask = task
    m_taskManager.AddRunner runner
    
    ' 将任务标记为运行中
    queue.MoveToRunning task
    
    ' 重要：更新任务类型的最后运行时间
    m_queueManager.UpdateTaskTypeLastRunTime task.taskType
    
    ' 标记任务开始运行
    MarkTaskRunning task
    
    ' 触发任务开始事件
    RaiseEvent TaskStarted(task)
    
    ' 执行任务
    runner.ExecuteTask
    
    Debug.Print ">>> 启动任务: " & task.taskName & _
               " (类型:" & task.taskType & _
               " 优先级:" & GetPriorityName(task.priority) & _
               " 队列:" & queue.queueName & ")"
    
    ' 启动定时器（如果还没有运行）
    If Not IsSchedulerTimerRunning() Then
        StartSchedulerTimer
        'Debug.Print ">>> 启动任务后启动调度器定时器"
    End If
    
    ExecuteTask = True
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> 执行任务时出错: " & Err.Description
    ExecuteTask = False
End Function

' ========== 定时器管理方法 ==========

Private Sub ManageTimerByRunningState()
    ' 检查调度器是否已初始化
    If Not IsInitialized() Then
        Exit Sub
    End If
    
    Dim hasTasks As Boolean
    hasTasks = m_queueManager.HasAnyTasks()
    
    Debug.Print ">>> 定时器状态管理 - 总任务数: " & m_queueManager.GetTotalTaskCount() & _
               " (等待:" & m_queueManager.GetTotalPendingCount() & " 运行中:" & m_queueManager.GetTotalRunningCount() & _
               "), 定时器运行: " & IsSchedulerTimerRunning()
    
    If hasTasks Then
        If Not IsSchedulerTimerRunning() Then
            StartSchedulerTimer
            'Debug.Print ">>> 有任务存在，启动定时器"
        End If
    Else
        If IsSchedulerTimerRunning() Then
            StopSchedulerTimer
            'Debug.Print ">>> 所有队列无任务，停止定时器"
        End If
    End If
End Sub

Private Function ProcessQueuesImmediately() As Boolean
    On Error GoTo ErrorHandler
    
    ' 检查调度器是否已初始化
    If Not IsInitialized() Then
        ProcessQueuesImmediately = False
        Exit Function
    End If
    
    m_taskManager.CheckCompletedProcesses m_queueManager
    m_taskManager.CheckPipeOutputs m_queueManager
    
    Dim totalPending As Long
    Dim totalRunning As Long

    totalPending = m_queueManager.GetTotalPendingCount()
    totalRunning = m_queueManager.GetTotalRunningCount()
    
    'Debug.Print ">>> 系统状态 - 等待任务: " & totalPending & ", 运行中任务: " & totalRunning
    
    ProcessQueuesByPriority
    
    m_taskManager.CleanCompletedRunners
    
    totalPending = m_queueManager.GetTotalPendingCount()
    totalRunning = m_queueManager.GetTotalRunningCount()
    
    'Debug.Print ">>> 处理结果 - 总等待: " & totalPending & ", 总运行中: " & totalRunning
    
    Dim hasTasks As Boolean
    hasTasks = m_queueManager.HasAnyTasks()
    
    'Debug.Print ">>> 需要定时器: " & hasTasks & " (总任务: " & (totalPending + totalRunning) & ")"
    
    If hasTasks And Not IsSchedulerTimerRunning() Then
        StartSchedulerTimer
        'Debug.Print ">>> 有任务存在但定时器未运行，立即启动定时器"
    End If
    
    ProcessQueuesImmediately = hasTasks
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> 立即处理队列时出错: " & Err.Description
    ProcessQueuesImmediately = False
End Function

' ========== 任务管理器事件处理 ==========

Private Sub m_taskManager_TaskCompleted(ByRef task As task)
    Debug.Print ">>> 收到任务完成事件: " & task.taskName & " (ID: " & task.taskID & ")"
    ' 转发事件到外部
    RaiseEvent TaskCompleted(task)
    
    ' 调用任务完成处理方法
    OnTaskCompleted task
End Sub

Private Sub m_taskManager_TaskFailed(ByRef task As task, ByVal errorMsg As String)
    Debug.Print ">>> 收到任务失败事件: " & task.taskName & " (ID: " & task.taskID & ") - " & errorMsg
    ' 转发事件到外部
    RaiseEvent TaskFailed(task, errorMsg)
    
    ' 调用任务失败处理方法
    OnTaskFailed task, errorMsg
End Sub

Private Sub m_taskManager_TaskTimeout(ByRef task As task)
    Debug.Print ">>> 收到任务超时事件: " & task.taskName & " (ID: " & task.taskID & ")"
    ' 转发事件到外部
    RaiseEvent TaskTimeout(task)
End Sub

Private Sub m_taskManager_TaskOutput(ByRef task As task, ByVal outputData As String, ByVal isErrorOutput As Boolean)
    Debug.Print ">>> 收到任务输出事件: " & task.taskName & " (ID: " & task.taskID & ") - " & outputData
    ' 转发事件到外部
    RaiseEvent TaskOutput(task, outputData, isErrorOutput)
End Sub

' ========== 任务事件处理方法 ==========

Public Sub OnTaskStarted(ByRef task As task)
    Debug.Print ">>> 任务开始运行: " & task.taskName & " (ID: " & task.taskID & ")"
    RaiseEvent TaskStarted(task)
End Sub


' ========== 直播间任务标记方法 ==========

Private Sub MarkTaskRunning(ByRef task As task)
    On Error Resume Next
    Dim room As LiveRoom
    Set room = FindLiveroomByTask(task)
    If Not room Is Nothing Then
        If task.taskType = "parser" Then
            room.MarkParseTaskRunning
        ElseIf task.taskType = "recorder" Then
            room.MarkRecordTaskRunning
        End If
    End If
End Sub

Private Sub MarkTaskCompleted(ByRef task As task)
    On Error Resume Next
    Dim room As LiveRoom
    Set room = FindLiveroomByTask(task)
    If Not room Is Nothing Then
        If task.taskType = "parser" Then
            room.MarkParseTaskCompleted
        ElseIf task.taskType = "recorder" Then
            room.MarkRecordTaskCompleted
        End If
    End If
End Sub

' ========== 辅助方法 ==========

' 辅助方法：将任务重新放回队列前端
Private Sub ReAddTaskToFront(ByRef queue As TaskQueue, ByRef task As task)
    On Error Resume Next
    
    ' 创建一个临时集合来保存所有任务
    Dim tempCollection As New Collection
    tempCollection.Add task ' 先添加当前任务
    
    ' 添加队列中现有的所有任务
    Dim i As Long
    For i = 1 To queue.PendingCount
        Dim existingTask As task
        Set existingTask = queue.GetPendingTask(1) ' 总是获取第一个
        If Not existingTask Is Nothing Then
            tempCollection.Add existingTask
        End If
    Next i
    
    ' 清空原队列
    Set queue = New TaskQueue
    queue.queueName = queue.queueName ' 保持队列名称
    queue.taskType = queue.taskType
    queue.priority = queue.priority
    queue.maxConcurrent = queue.maxConcurrent
    queue.minInterval = queue.minInterval
    queue.DefaultTimeoutSeconds = queue.DefaultTimeoutSeconds
    
    ' 重新添加所有任务（保持原有顺序）
    For i = 1 To tempCollection.count
        queue.AddTask tempCollection(i)
    Next i
    
    Debug.Print ">>> 任务 " & task.taskName & " 已重新放回队列前端"
End Sub


' 检查队列是否可以运行任务
Private Function CheckQueueCanRunTask(ByRef queue As TaskQueue, ByRef Reason As String) As Boolean
    Dim CanRun As Boolean
    m_queueManager.CheckQueueRunTaskDetailed queue, CanRun, Reason
    CheckQueueCanRunTask = CanRun
End Function

' ========== 公共接口方法 ==========

Public Sub ProcessQueues()
    ProcessQueuesImmediately
End Sub

Public Sub StartScheduler()
    Debug.Print ">>> 调度器启动"
    ProcessQueuesImmediately
End Sub

Public Sub StopScheduler()
    Debug.Print ">>> 调度器停止"
    If IsSchedulerTimerRunning() Then
        StopSchedulerTimer
    End If
End Sub

Public Sub RaiseTaskOutput(ByRef task As task, ByVal outputData As String, ByVal isErrorOutput As Boolean)
    RaiseEvent TaskOutput(task, outputData, isErrorOutput)
End Sub

Public Function GetQueues() As Collection
    Set GetQueues = m_queueManager.GetAllQueues()
End Function

' ========== 任务类型相关方法 ==========

Public Function CanRunTaskType(ByVal taskType As String) As Boolean
    CanRunTaskType = m_queueManager.CanRunTaskType(taskType)
End Function

Public Function GetTaskTypeRunningCount(ByVal taskType As String) As Long
    GetTaskTypeRunningCount = m_queueManager.GetTaskTypeRunningCount(taskType)
End Function

Public Function GetTaskTypeMaxConcurrent(ByVal taskType As String) As Long
    GetTaskTypeMaxConcurrent = m_queueManager.GetTaskTypeMaxConcurrent(taskType)
End Function

Public Function GetTaskTypeLastRunTime(ByVal taskType As String) As Currency
    GetTaskTypeLastRunTime = m_queueManager.GetTaskTypeLastRunTime(taskType)
End Function

' ========== 初始化状态检查 ==========

' 检查调度器是否已正确初始化
Public Function IsInitialized() As Boolean
    On Error Resume Next
    If m_queueManager Is Nothing Then
        IsInitialized = False
        Exit Function
    End If
    
    If m_taskManager Is Nothing Then
        IsInitialized = False
        Exit Function
    End If
    
    ' 测试调用一个简单的方法
    Dim queueCount As Long
    queueCount = m_queueManager.GetTotalTaskCount()
    
    IsInitialized = (Err.Number = 0)
End Function

' 重新初始化调度器
Public Sub Reinitialize()
    On Error Resume Next
    Debug.Print ">>> 重新初始化调度器..."
    
    Cleanup
    
    Set m_queueManager = New TaskSchedulerQueueManager
    Set m_taskManager = New TaskSchedulerTaskManager
    m_taskManager.SetParent Me
    
    WithEvents m_taskManager
    Set WithEvents m_taskManager = m_taskManager
    
    Initialize
    
    Debug.Print ">>> 调度器重新初始化完成"
End Sub

' 调试初始化过程
Public Sub DebugInitialize()
    Debug.Print ">>> 调试初始化过程..."
    Debug.Print ">>> m_queueManager: " & IIf(m_queueManager Is Nothing, "未设置", "已设置")
    Debug.Print ">>> m_taskManager: " & IIf(m_taskManager Is Nothing, "未设置", "已设置")
    
    If Not m_queueManager Is Nothing Then
        On Error Resume Next
        Debug.Print ">>> 队列数量: " & m_queueManager.GetAllQueues().count
    End If
End Sub

' ========== 关闭管理功能 ==========

Public Sub InitiateShutdown()
    If m_shutdownInProgress Then Exit Sub
    
    Debug.Print ">>> 开始关闭程序 - 停止接受新任务"
    m_shutdownInProgress = True
    m_shutdownInitiated = True
    
    StopScheduler
    SendQuitToAllFFmpegProcesses
    StartShutdownMonitor
End Sub

Public Function CanSafeExit() As Boolean
    If Not m_shutdownInitiated Then
        CanSafeExit = False
        Exit Function
    End If
    
    Dim tasksRunning As Boolean
    tasksRunning = m_queueManager.HasAnyTasks()
    
    If tasksRunning Then
        Debug.Print ">>> 仍有任务存在，等待完成..."
        CanSafeExit = False
    Else
        Debug.Print ">>> 所有任务已完成，可以安全退出"
        CanSafeExit = True
    End If
End Function

Public Property Get ShutdownInProgress() As Boolean
    ShutdownInProgress = m_shutdownInProgress
End Property

Public Property Get ShutdownInitiated() As Boolean
    ShutdownInitiated = m_shutdownInitiated
End Property

' ========== 清理方法 ==========

Private Sub Cleanup()
    If IsSchedulerTimerRunning() Then
        StopSchedulerTimer
    End If
    
    Set m_queueManager = Nothing
    Set m_taskManager = Nothing
End Sub

' ========== 管道命令发送功能 ==========
Public Function SendQuitToFFmpeg(ByRef task As task) As Boolean
    On Error GoTo ErrorHandler
    
    Debug.Print ">>> 发送退出命令到: " & task.taskName
    
    ' 检查任务是否已经被处理过
    If task.status = tsCompleted Or task.status = tsFailed Then
        Debug.Print ">>> 任务已完成或失败，跳过处理"
        SendQuitToFFmpeg = True
        Exit Function
    End If
    
    If task.hPipeWrite = 0 Then
        Debug.Print ">>> 管道句柄无效"
        SendQuitToFFmpeg = True
        Exit Function
    End If
    
    ' 检查进程状态
    If Not IsProcessStillRunning(task.hProcess) Then
        Debug.Print ">>> 进程已退出，无需发送命令"
        ' 更新任务状态
        task.status = tsCompleted
        SendQuitToFFmpeg = True
        Exit Function
    End If
    
    Debug.Print ">>> 进程仍在运行，准备发送退出命令"
    
    ' 使用字节数组发送命令
    Dim commandBytes() As Byte
    commandBytes = StrConv("q" & vbCrLf, vbFromUnicode)
    
    Dim bytesWritten As Long
    Dim success As Long
    
    success = WriteFile(task.hPipeWrite, commandBytes(0), UBound(commandBytes) + 1, bytesWritten, ByVal 0&)
    
    If success <> 0 Then
        If bytesWritten > 0 Then
            Debug.Print ">>> 成功发送退出命令，写入字节: " & bytesWritten
            ' 等待进程退出
            WaitForProcessExit task.hProcess, 3000 ' 等待3秒
            
            If IsProcessStillRunning(task.hProcess) Then
                Debug.Print ">>> 进程未响应退出命令，强制终止"
                TerminateProcess task.hProcess, 1
                task.status = tsCompleted
            Else
                Debug.Print ">>> 进程已优雅退出"
                task.status = tsCompleted
            End If
        Else
            Debug.Print ">>> 写入0字节，管道可能已关闭，强制终止进程"
            TerminateProcess task.hProcess, 1
            task.status = tsCompleted
        End If
    Else
        Debug.Print ">>> WriteFile 失败，错误代码: " & Err.LastDllError
        ' 写入失败，直接终止进程
        Debug.Print ">>> 管道写入失败，强制终止进程"
        TerminateProcess task.hProcess, 1
        task.status = tsCompleted
    End If
    
    ' 清理资源
    task.Cleanup
    
    Debug.Print ">>> 退出命令处理完成"
    SendQuitToFFmpeg = True
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> 发送退出命令时出错: " & Err.Description
    ' 出错时也标记为完成
    task.status = tsCompleted
    task.Cleanup
    SendQuitToFFmpeg = True
End Function

' 检查管道是否可写
Private Function IsPipeWriteable(ByVal hPipe As Long) As Boolean
    On Error Resume Next
    Dim testData(0) As Byte
    Dim bytesWritten As Long
    Dim success As Long
    
    ' 尝试写入一个测试字节
    success = WriteFile(hPipe, testData(0), 0, bytesWritten, ByVal 0&)
    
    If success <> 0 Then
        IsPipeWriteable = True
        Debug.Print ">>> 管道可写检查: 通过"
    Else
        IsPipeWriteable = False
        Debug.Print ">>> 管道可写检查: 失败，错误: " & Err.LastDllError
    End If
End Function

' 等待进程退出
Private Sub WaitForProcessExit(ByVal hProcess As Long, ByVal timeoutMs As Long)
    On Error Resume Next
    Dim startTime As Long
    startTime = GetTickCount()
    
    Do While GetTickCount() - startTime < timeoutMs
        If Not IsProcessStillRunning(hProcess) Then
            Debug.Print ">>> 进程已正常退出"
            Exit Sub
        End If
        Sleep 100 ' 等待100ms
    Loop
    
    If IsProcessStillRunning(hProcess) Then
        Debug.Print ">>> 等待超时，进程仍在运行"
    End If
End Sub

' 检查进程是否运行
Private Function IsProcessStillRunning(ByVal hProcess As Long) As Boolean
    On Error Resume Next
    Dim exitCode As Long
    Dim result As Long
    
    result = GetExitCodeProcess(hProcess, exitCode)
    
    If result = 0 Then
        IsProcessStillRunning = False
    Else
        IsProcessStillRunning = (exitCode = STILL_ACTIVE)
    End If
End Function



' 停止所有录制任务
Public Function StopAllRecordTasks() As Boolean
    On Error GoTo ErrorHandler
    
    ' 检查调度器是否已初始化
    If Not IsInitialized() Then
        Debug.Print ">>> 调度器未初始化，无法停止任务"
        StopAllRecordTasks = False
        Exit Function
    End If
    
    Dim stoppedCount As Long
    Dim removedCount As Long
    stoppedCount = 0
    removedCount = 0
    
    Dim queue As TaskQueue
    Dim task As task
    Dim i As Long
    
    ' 先停止所有运行中的录制任务
    For Each queue In m_queueManager.GetAllQueues()
        For i = 1 To queue.RunningCount
            Set task = queue.GetRunningTask(i)
            If Not task Is Nothing Then
                If task.status = tsRunning And task.taskType = "recorder" Then
                    If SendQuitToFFmpeg(task) Then
                        stoppedCount = stoppedCount + 1
                        Debug.Print ">>> 已向录制任务发送退出命令: " & task.taskName
                        
                        ' 更新对应的直播间状态
                        UpdateLiveroomRecordStatus task.web_rid, False
                    End If
                End If
            End If
        Next i
    Next queue
    
    ' 然后移除所有等待中的录制任务
    For Each queue In m_queueManager.GetAllQueues()
        i = 1
        While i <= queue.PendingCount
            Set task = queue.GetPendingTask(i)
            If Not task Is Nothing Then
                If task.taskType = "recorder" Then
                    RemoveTaskFromPending queue, task.taskID
                    removedCount = removedCount + 1
                    Debug.Print ">>> 已从等待队列移除录制任务: " & task.taskName
                    
                    ' 更新对应的直播间状态
                    UpdateLiveroomRecordStatus task.web_rid, False
                Else
                    i = i + 1
                End If
            Else
                i = i + 1
            End If
        Wend
    Next queue
    
    If stoppedCount > 0 Or removedCount > 0 Then
        StopAllRecordTasks = True
        Debug.Print ">>> 成功停止 " & stoppedCount & " 个运行中录制任务，移除 " & removedCount & " 个等待中录制任务"
    Else
        StopAllRecordTasks = False
        Debug.Print ">>> 没有找到需要停止的录制任务"
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> 停止所有录制任务时出错: " & Err.Description
    StopAllRecordTasks = False
End Function

' 向所有 FFmpeg 进程发送退出命令
Private Sub SendQuitToAllFFmpegProcesses()
    On Error Resume Next
    
    ' 检查调度器是否已初始化
    If Not IsInitialized() Then
        Debug.Print ">>> 调度器未初始化，无法发送退出命令"
        Exit Sub
    End If
    
    Dim queue As TaskQueue
    Dim task As task
    Dim i As Long
    Dim ffmpegCount As Long
    
    ffmpegCount = 0
    
    For Each queue In m_queueManager.GetAllQueues()
        For i = 1 To queue.RunningCount
            Set task = queue.GetRunningTask(i)
            If Not task Is Nothing Then
                If task.status = tsRunning And task.taskType = "recorder" Then
                    If task.hProcess <> 0 And task.monitorPipe Then
                        ' 向 FFmpeg 发送 'q' 命令使其安全退出
                        If SendQuitToFFmpeg(task) Then
                            ffmpegCount = ffmpegCount + 1
                            Debug.Print ">>> 已向 FFmpeg 进程发送退出命令: " & task.taskName
                            
                            ' 更新对应的直播间状态
                            UpdateLiveroomRecordStatus task.web_rid, False
                        End If
                    End If
                End If
            End If
        Next i
    Next queue
    
    If ffmpegCount > 0 Then
        Debug.Print ">>> 总共向 " & ffmpegCount & " 个 FFmpeg 进程发送了退出命令"
    Else
        Debug.Print ">>> 没有找到需要发送退出命令的 FFmpeg 进程"
    End If
End Sub

' ========== 辅助方法 ==========

' 从等待队列中移除任务
Private Sub RemoveTaskFromPending(ByRef queue As TaskQueue, ByVal taskID As Long)
    On Error Resume Next
    Dim i As Long
    For i = 1 To queue.PendingCount
        Dim pendingTask As task
        Set pendingTask = queue.GetPendingTask(i)
        If Not pendingTask Is Nothing Then
            If pendingTask.taskID = taskID Then
                ' 创建一个临时集合来保存剩余任务
                Dim tempCollection As New Collection
                Dim j As Long
                
                ' 添加除了要移除的任务之外的所有任务
                For j = 1 To queue.PendingCount
                    Set pendingTask = queue.GetPendingTask(j)
                    If Not pendingTask Is Nothing Then
                        If pendingTask.taskID <> taskID Then
                            tempCollection.Add pendingTask
                        End If
                    End If
                Next j
                
                ' 清空原队列
                Set queue = New TaskQueue
                queue.queueName = queue.queueName
                queue.taskType = queue.taskType
                queue.priority = queue.priority
                queue.maxConcurrent = queue.maxConcurrent
                queue.minInterval = queue.minInterval
                queue.DefaultTimeoutSeconds = queue.DefaultTimeoutSeconds
                
                ' 重新添加所有剩余任务
                For j = 1 To tempCollection.count
                    queue.AddTask tempCollection(j)
                Next j
                
                Debug.Print ">>> 从等待队列移除任务 ID: " & taskID
                Exit For
            End If
        End If
    Next i
End Sub

' 更新直播间录制状态
Private Sub UpdateLiveroomRecordStatus(ByVal web_rid As String, ByVal isRecording As Boolean)
    On Error Resume Next
    Dim room As LiveRoom
    Set room = g_liveroomManager.GetLiveroom(web_rid)
    If Not room Is Nothing Then
        room.isRecording = isRecording
        If Not isRecording Then
            room.RecordTaskID = 0
            room.RecordTaskStatus = "未录制"
            Debug.Print ">>> 更新直播间录制状态: " & web_rid & " -> 未录制"
        End If
    End If
End Sub
Private Sub StartShutdownMonitor()
    If m_shutdownInitiated And Not IsSchedulerTimerRunning() Then
        StartSchedulerTimer
        Debug.Print ">>> 启动关闭监控定时器"
    End If
End Sub

' ========== 辅助函数 ==========

Private Function GetPriorityName(ByVal priority As taskPriority) As String
    Select Case priority
        Case tpHigh: GetPriorityName = "高"
        Case tpMedium: GetPriorityName = "中"
        Case tpLow: GetPriorityName = "低"
        Case Else: GetPriorityName = "未知"
    End Select
End Function



' 新增：停止录制任务时标记为停止中状态
Public Function StopRecordTask(ByVal web_rid As String) As Boolean
    On Error GoTo ErrorHandler
    
    If m_shutdownInProgress Then
        Debug.Print ">>> 关闭进行中，无法停止单个任务"
        StopRecordTask = False
        Exit Function
    End If
    
    ' 检查调度器是否已初始化
    If Not IsInitialized() Then
        Debug.Print ">>> 调度器未初始化，无法停止任务"
        StopRecordTask = False
        Exit Function
    End If
    
    Dim taskFound As Boolean
    taskFound = False
    
    Dim queue As TaskQueue
    Dim task As task
    Dim i As Long
    
    ' 在所有队列中查找对应的录制任务
    For Each queue In m_queueManager.GetAllQueues()
        ' 先检查运行中的任务
        For i = 1 To queue.RunningCount
            Set task = queue.GetRunningTask(i)
            If Not task Is Nothing Then
                If task.status = tsRunning And task.taskType = "recorder" And task.web_rid = web_rid Then
                    Debug.Print ">>> 找到运行中的录制任务，准备停止: " & task.taskName & " (web_rid: " & web_rid & ")"
                    
                    ' 检查任务是否已经被处理过
                    If task.hProcess = 0 Then
                        Debug.Print ">>> 任务已被处理过，跳过"
                        taskFound = True
                        Exit For
                    End If
                    
                    ' 标记直播间录制任务为停止中状态
                    MarkRecordTaskStopping web_rid
                    
                    ' 发送退出命令
                    If SendQuitToFFmpeg(task) Then
                        taskFound = True
                        Debug.Print ">>> 已处理录制任务: " & task.taskName
                        
                        ' 立即从运行队列中移除
                        queue.RemoveFromRunning task.taskID
                    End If
                    
                    Exit For
                End If
            End If
        Next i
        
        ' 如果没有找到运行中的任务，检查等待中的任务
        If Not taskFound Then
            For i = 1 To queue.PendingCount
                Set task = queue.GetPendingTask(i)
                If Not task Is Nothing Then
                    If task.taskType = "recorder" And task.web_rid = web_rid Then
                        Debug.Print ">>> 找到等待中的录制任务，直接移除: " & task.taskName & " (web_rid: " & web_rid & ")"
                        
                        ' 从等待队列中移除
                        RemoveTaskFromPending queue, task.taskID
                        taskFound = True
                        Debug.Print ">>> 已从等待队列移除录制任务: " & task.taskName
                        Exit For
                    End If
                End If
            Next i
        End If
        
        If taskFound Then Exit For
    Next queue
    
    If taskFound Then
        StopRecordTask = True
        Debug.Print ">>> 成功停止录制任务: " & web_rid
    Else
        StopRecordTask = False
        Debug.Print ">>> 未找到运行中或等待中的录制任务: " & web_rid
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> 停止录制任务时出错: " & Err.Description
    StopRecordTask = False
End Function

' 新增：标记录制任务为停止中状态
Private Sub MarkRecordTaskStopping(ByVal web_rid As String)
    On Error Resume Next
    Dim room As LiveRoom
    Set room = FindLiveroomByTask(web_rid)
    If Not room Is Nothing Then
        room.MarkRecordTaskStopping
        Debug.Print ">>> 标记录制任务为停止中状态: " & web_rid
    End If
End Sub

' 修改：根据 web_rid 查找直播间
Private Function FindLiveroomByTask(ByVal web_rid As String) As LiveRoom
    On Error Resume Next
    Set FindLiveroomByTask = g_liveroomManager.GetLiveroom(web_rid)
End Function

' 修改：任务完成时的处理 - 同步状态到直播间
Public Sub OnTaskCompleted(ByRef task As task)
    Debug.Print ">>> 收到任务完成通知: " & task.taskName & " (ID: " & task.taskID & ")"
    
    ' 清理资源
    task.Cleanup
    
    ' 找到队列并移动任务
    Dim queue As TaskQueue
    For Each queue In m_queueManager.GetAllQueues()
        queue.MoveToCompleted task
    Next queue
    
    ' 根据任务类型同步状态到直播间
    If task.taskType = "parser" Then
        MarkParseTaskCompleted task
    ElseIf task.taskType = "recorder" Then
        MarkRecordTaskCompleted task
    End If
    
    ' 检查是否需要停止定时器
    If Not m_queueManager.HasAnyTasks() Then
        If IsSchedulerTimerRunning() Then
            Debug.Print ">>> 所有任务完成，停止定时器"
            StopSchedulerTimer
        End If
    Else
        Debug.Print ">>> 仍有任务存在，保持定时器运行"
    End If
End Sub

' 修改：任务失败时的处理 - 同步状态到直播间
Public Sub OnTaskFailed(ByRef task As task, ByVal errorMsg As String)
    Debug.Print ">>> 收到任务失败通知: " & task.taskName & " (ID: " & task.taskID & ") - " & errorMsg
    
    ' 清理资源
    task.Cleanup
    
    ' 找到队列并移动任务
    Dim queue As TaskQueue
    For Each queue In m_queueManager.GetAllQueues()
        queue.MoveToCompleted task
    Next queue
    
    ' 根据任务类型同步状态到直播间
    If task.taskType = "parser" Then
        MarkParseTaskCompleted task
    ElseIf task.taskType = "recorder" Then
        MarkRecordTaskCompleted task
    End If
    
    RaiseEvent TaskFailed(task, errorMsg)
    
    ' 检查是否需要停止定时器
    If Not m_queueManager.HasAnyTasks() Then
        If IsSchedulerTimerRunning() Then
            Debug.Print ">>> 所有任务完成，停止定时器"
            StopSchedulerTimer
        End If
    Else
        Debug.Print ">>> 仍有任务存在，保持定时器运行"
    End If
End Sub

' 新增：标记解析任务完成
Private Sub MarkParseTaskCompleted(ByRef task As task)
    On Error Resume Next
    Dim room As LiveRoom
    Set room = FindLiveroomByTask(task.web_rid)
    If Not room Is Nothing Then
        room.MarkParseTaskCompleted
        Debug.Print ">>> 同步解析任务完成状态到直播间: " & task.web_rid
    End If
End Sub

' 新增：标记录制任务完成
Private Sub MarkRecordTaskCompleted(ByRef task As task)
    On Error Resume Next
    Dim room As LiveRoom
    Set room = FindLiveroomByTask(task.web_rid)
    If Not room Is Nothing Then
        room.MarkRecordTaskCompleted
        Debug.Print ">>> 同步录制任务完成状态到直播间: " & task.web_rid
    End If
End Sub

