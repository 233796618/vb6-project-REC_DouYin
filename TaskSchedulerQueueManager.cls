VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskSchedulerQueueManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 队列集合
Private m_queues As Collection
Private m_taskTypeConfig As Collection
' 新增：任务类型最后运行时间跟踪
Private m_taskTypeLastRunTimes As Collection

Private Sub Class_Initialize()
    Set m_queues = New Collection
    Set m_taskTypeConfig = New Collection
    Set m_taskTypeLastRunTimes = New Collection
    Debug.Print "TaskSchedulerQueueManager 初始化完成"
End Sub

' ========== 队列管理方法 ==========

' 添加任务类型配置
Public Sub AddTaskType(ByVal taskType As String, ByVal maxConcurrent As Long, ByVal minInterval As Long, ByVal timeoutSeconds As Long)
    On Error GoTo ErrorHandler
    
    ' 检查集合是否已初始化
    If m_queues Is Nothing Then
        Set m_queues = New Collection
        Debug.Print ">>> 重新初始化队列集合"
    End If
    
    If m_taskTypeConfig Is Nothing Then
        Set m_taskTypeConfig = New Collection
        Debug.Print ">>> 重新初始化任务类型配置集合"
    End If
    
    If m_taskTypeLastRunTimes Is Nothing Then
        Set m_taskTypeLastRunTimes = New Collection
        Debug.Print ">>> 重新初始化任务类型时间跟踪集合"
    End If
    
    Dim config As New Collection
    config.Add maxConcurrent, "maxConcurrent"
    config.Add minInterval, "minInterval"
    config.Add timeoutSeconds, "timeoutSeconds"
    
    ' 如果已存在该任务类型，先移除
    On Error Resume Next
    m_taskTypeConfig.Remove taskType
    m_taskTypeLastRunTimes.Remove taskType
    On Error GoTo ErrorHandler
    
    m_taskTypeConfig.Add config, taskType
    ' 初始化任务类型的最后运行时间为0
    m_taskTypeLastRunTimes.Add 0, taskType
    
    Dim priorityNames(2) As String
    priorityNames(tpHigh) = "高"
    priorityNames(tpMedium) = "中"
    priorityNames(tpLow) = "低"
    
    Dim i As Long
    For i = tpHigh To tpLow
        Dim queueName As String
        queueName = GetQueueName(taskType, i)
        
        ' 如果已存在该队列，先移除
        On Error Resume Next
        m_queues.Remove queueName
        On Error GoTo ErrorHandler
        
        Dim queue As TaskQueue
        Set queue = New TaskQueue
        queue.queueName = queueName
        queue.taskType = taskType
        queue.priority = i
        queue.maxConcurrent = maxConcurrent
        queue.minInterval = minInterval
        queue.DefaultTimeoutSeconds = timeoutSeconds
        
        m_queues.Add queue, queueName
        Debug.Print "添加队列: " & queueName & " (优先级:" & priorityNames(i) & " 类型并发:" & maxConcurrent & " 间隔:" & minInterval & "ms)"
    Next i
    
    Exit Sub
    
ErrorHandler:
    Debug.Print ">>> 添加任务类型时出错: " & Err.Description & " (任务类型: " & taskType & ")"
    ' 重新抛出错误
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' ========== 任务类型时间管理方法 ==========

' 更新任务类型的最后运行时间
Public Sub UpdateTaskTypeLastRunTime(ByVal taskType As String)
    On Error Resume Next
    If m_taskTypeLastRunTimes Is Nothing Then
        Set m_taskTypeLastRunTimes = New Collection
    End If
    
    Dim currentTime As Currency
    currentTime = GetHighResTimeMs()
    
    ' 如果任务类型不存在，先添加
    On Error Resume Next
    Dim existingTime As Currency
    existingTime = m_taskTypeLastRunTimes(taskType)
    If Err.Number <> 0 Then
        m_taskTypeLastRunTimes.Add currentTime, taskType
    Else
        m_taskTypeLastRunTimes.Remove taskType
        m_taskTypeLastRunTimes.Add currentTime, taskType
    End If
    
    Debug.Print ">>> 更新任务类型 '" & taskType & "' 最后运行时间: " & Format(currentTime, "0.000") & "ms"
End Sub

' 获取任务类型的最后运行时间
Public Function GetTaskTypeLastRunTime(ByVal taskType As String) As Currency
    On Error Resume Next
    If m_taskTypeLastRunTimes Is Nothing Then
        GetTaskTypeLastRunTime = 0
        Exit Function
    End If
    
    GetTaskTypeLastRunTime = m_taskTypeLastRunTimes(taskType)
    If Err.Number <> 0 Then
        GetTaskTypeLastRunTime = 0
    End If
End Function

' ========== 队列运行检查方法 ==========

' 详细检查队列是否可以运行任务 - 使用输出参数
Public Sub CheckQueueRunTaskDetailed(ByRef queue As TaskQueue, ByRef CanRun As Boolean, ByRef Reason As String)
    On Error GoTo ErrorHandler
    
    ' 检查1: 队列级别的并发限制
    If queue.RunningCount >= queue.maxConcurrent Then
        CanRun = False
        Reason = "队列并发限制: " & queue.RunningCount & "/" & queue.maxConcurrent & " 运行中"
        Exit Sub
    End If
    
    ' 检查2: 任务类型级别的并发限制
    Dim currentRunning As Long
    Dim maxConcurrent As Long
    currentRunning = GetTaskTypeRunningCount(queue.taskType)
    maxConcurrent = GetTaskTypeMaxConcurrent(queue.taskType)
    
    If currentRunning >= maxConcurrent Then
        CanRun = False
        Reason = "任务类型并发限制: " & currentRunning & "/" & maxConcurrent & " 运行中"
        Exit Sub
    End If
    
    ' 检查3: 任务类型级别的时间间隔限制
    Dim lastTaskTypeRunTime As Currency
    lastTaskTypeRunTime = GetTaskTypeLastRunTime(queue.taskType)
    Dim currentTime As Currency
    currentTime = GetHighResTimeMs()
    Dim timeDiff As Currency
    timeDiff = currentTime - lastTaskTypeRunTime
    
    If timeDiff < queue.minInterval Then
        CanRun = False
        Reason = "任务类型时间间隔限制: 距离上次运行 " & Format(timeDiff, "0.0") & "ms < " & queue.minInterval & "ms"
        Exit Sub
    End If
    
    ' 检查4: 队列级别的时间间隔限制
    Dim queueTimeDiff As Currency
    queueTimeDiff = currentTime - queue.LastRunTimeMs
    
    If queueTimeDiff < queue.minInterval Then
        CanRun = False
        Reason = "队列时间间隔限制: 距离上次运行 " & Format(queueTimeDiff, "0.0") & "ms < " & queue.minInterval & "ms"
        Exit Sub
    End If
    
    ' 所有检查通过
    CanRun = True
    Reason = "可以运行任务"
    Exit Sub
    
ErrorHandler:
    Debug.Print ">>> 检查队列运行状态时出错: " & Err.Description
    CanRun = False
    Reason = "检查队列状态时出错: " & Err.Description
End Sub
' ========== 队列管理方法 ==========


' 获取队列
Public Function GetQueue(ByVal taskType As String, ByVal priority As taskPriority) As TaskQueue
    On Error Resume Next
    Dim queueName As String
    queueName = GetQueueName(taskType, priority)
    Set GetQueue = m_queues(queueName)
End Function

' 获取所有队列
Public Function GetAllQueues() As Collection
    Set GetAllQueues = m_queues
End Function

' 获取默认队列
Public Function GetDefaultQueue() As TaskQueue
    Set GetDefaultQueue = GetQueue("default", tpMedium)
End Function

' ========== 任务统计方法 ==========

' 获取总任务数
Public Function GetTotalTaskCount() As Long
    GetTotalTaskCount = GetTotalPendingCount() + GetTotalRunningCount()
End Function

' 获取总等待任务数
Public Function GetTotalPendingCount() As Long
    On Error Resume Next
    Dim queue As TaskQueue
    Dim totalPending As Long
    
    totalPending = 0
    If Not m_queues Is Nothing Then
        For Each queue In m_queues
            totalPending = totalPending + queue.PendingCount
        Next queue
    End If
    
    GetTotalPendingCount = totalPending
End Function

' 获取总运行中任务数
Public Function GetTotalRunningCount() As Long
    On Error Resume Next
    Dim queue As TaskQueue
    Dim totalRunning As Long
    
    totalRunning = 0
    If Not m_queues Is Nothing Then
        For Each queue In m_queues
            totalRunning = totalRunning + queue.RunningCount
        Next queue
    End If
    
    GetTotalRunningCount = totalRunning
End Function

' 检查是否有任何任务
Public Function HasAnyTasks() As Boolean
    HasAnyTasks = (GetTotalTaskCount() > 0)
End Function

' 检查是否有运行中任务
Public Function HasRunningTasks() As Boolean
    HasRunningTasks = (GetTotalRunningCount() > 0)
End Function

' ========== 任务类型管理方法 ==========

' 获取任务类型的运行中任务数
Public Function GetTaskTypeRunningCount(ByVal taskType As String) As Long
    On Error Resume Next
    Dim queue As TaskQueue
    Dim totalRunning As Long
    
    totalRunning = 0
    If Not m_queues Is Nothing Then
        For Each queue In m_queues
            If queue.taskType = taskType Then
                totalRunning = totalRunning + queue.RunningCount
            End If
        Next queue
    End If
    
    GetTaskTypeRunningCount = totalRunning
End Function

' 获取任务类型的最大并发数
Public Function GetTaskTypeMaxConcurrent(ByVal taskType As String) As Long
    On Error Resume Next
    Dim config As Collection
    Set config = m_taskTypeConfig(taskType)
    If Err.Number = 0 Then
        GetTaskTypeMaxConcurrent = config("maxConcurrent")
    Else
        GetTaskTypeMaxConcurrent = 3
    End If
End Function



' 检查任务类型是否可以运行
Public Function CanRunTaskType(ByVal taskType As String) As Boolean
    Dim currentRunning As Long
    Dim maxConcurrent As Long
    
    currentRunning = GetTaskTypeRunningCount(taskType)
    maxConcurrent = GetTaskTypeMaxConcurrent(taskType)
    
    If currentRunning >= maxConcurrent Then
        Debug.Print ">>> 任务类型 '" & taskType & "' 并发限制: " & currentRunning & "/" & maxConcurrent
        CanRunTaskType = False
    Else
        CanRunTaskType = True
    End If
End Function

' 获取默认超时时间
Public Function GetDefaultTimeout(ByVal taskType As String) As Long
    On Error Resume Next
    Dim config As Collection
    Set config = m_taskTypeConfig(taskType)
    If Err.Number = 0 Then
        GetDefaultTimeout = config("timeoutSeconds")
    Else
        GetDefaultTimeout = 300
    End If
End Function

' ========== 队列运行检查方法 ==========


' 简化版本 - 只返回布尔值
Public Function CanQueueRunTask(ByRef queue As TaskQueue) As Boolean
    Dim CanRun As Boolean
    Dim Reason As String
    CheckQueueRunTaskDetailed queue, CanRun, Reason
    CanQueueRunTask = CanRun
End Function

' ========== 日志和调试方法 ==========

' 记录详细的等待原因
Public Sub LogDetailedWaitingReasons()
    On Error Resume Next
    Dim totalPending As Long
    totalPending = GetTotalPendingCount()
    
    If totalPending > 0 Then
        Debug.Print ">>> 详细等待原因分析:"
        Debug.Print ">>> 仍有 " & totalPending & " 个任务在等待"
        
        Dim queue As TaskQueue
        If Not m_queues Is Nothing Then
            For Each queue In m_queues
                If queue.PendingCount > 0 Then
                    Dim CanRun As Boolean
                    Dim Reason As String
                    CheckQueueRunTaskDetailed queue, CanRun, Reason
                    
                    Debug.Print ">>>   " & queue.queueName & ": " & queue.PendingCount & " 等待 - " & Reason
                    
                    ' 输出队列详细信息
                    Debug.Print ">>>     队列状态: 运行中 " & queue.RunningCount & "/" & queue.maxConcurrent & _
                               ", 最后运行: " & Format(queue.LastRunTimeMs, "0.0") & "ms"
                    
                    ' 输出任务类型详细信息
                    Dim typeRunning As Long
                    Dim typeMaxConcurrent As Long
                    typeRunning = GetTaskTypeRunningCount(queue.taskType)
                    typeMaxConcurrent = GetTaskTypeMaxConcurrent(queue.taskType)
                    Debug.Print ">>>     任务类型状态: " & queue.taskType & " 运行中 " & typeRunning & "/" & typeMaxConcurrent & _
                               ", 最后运行: " & Format(GetTaskTypeLastRunTime(queue.taskType), "0.0") & "ms"
                End If
            Next queue
        End If
    End If
End Sub

' 记录调度器状态
Public Sub LogSchedulerStatus()
    On Error Resume Next
    Dim totalPending As Long
    Dim totalRunning As Long
    Dim totalCompleted As Long
    Dim queue As TaskQueue
    
    totalPending = GetTotalPendingCount()
    totalRunning = GetTotalRunningCount()
    totalCompleted = 0
    
    If Not m_queues Is Nothing Then
        For Each queue In m_queues
            totalCompleted = totalCompleted + queue.completedCount
        Next queue
    End If
    
    Debug.Print ">>> 调度器状态 - 等待:" & totalPending & " 运行中:" & totalRunning & " 已完成:" & totalCompleted
    
    ' 输出各任务类型状态
    Dim taskTypes As New Collection
    Dim taskType As Variant
    
    ' 收集所有任务类型
    If Not m_queues Is Nothing Then
        For Each queue In m_queues
            On Error Resume Next
            taskTypes.Add queue.taskType, queue.taskType
        Next queue
    End If
    
    ' 输出每个任务类型的详细状态
    For Each taskType In taskTypes
        Dim typeRunning As Long
        Dim typeMaxConcurrent As Long
        typeRunning = GetTaskTypeRunningCount(taskType)
        typeMaxConcurrent = GetTaskTypeMaxConcurrent(taskType)
        
        Debug.Print ">>>   类型 '" & taskType & "': " & typeRunning & "/" & typeMaxConcurrent & " 运行中"
    Next taskType
End Sub

' ========== 辅助函数 ==========

Private Function GetQueueName(ByVal taskType As String, ByVal priority As taskPriority) As String
    Dim priorityNames(2) As String
    priorityNames(tpHigh) = "高"
    priorityNames(tpMedium) = "中"
    priorityNames(tpLow) = "低"
    
    GetQueueName = taskType & "-" & priorityNames(priority)
End Function

Private Function GetHighResTimeMs() As Currency
    Dim count As Currency
    Dim freq As Currency
    
    QueryPerformanceFrequency freq
    QueryPerformanceCounter count
    
    If freq > 0 Then
        GetHighResTimeMs = (count / freq) * 1000
    Else
        GetHighResTimeMs = Timer * 1000
    End If
End Function

