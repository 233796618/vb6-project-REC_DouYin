VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 队列配置
Public queueName As String
Public taskType As String
Public priority As taskPriority
Public maxConcurrent As Long
Public minInterval As Long
Public DefaultTimeoutSeconds As Long

' 高精度时间相关
Private m_lastRunTimeHighRes As Currency
Private m_performanceFrequency As Currency

' 任务集合
Private m_pendingTasks As Collection
Private m_runningTasks As Collection
Private m_completedTasks As Collection

' API 声明
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Public Property Get PendingCount() As Long
    PendingCount = m_pendingTasks.count
End Property

Public Property Get RunningCount() As Long
    RunningCount = m_runningTasks.count
End Property

Public Property Get completedCount() As Long
    completedCount = m_completedTasks.count
End Property

Public Property Get TotalCount() As Long
    TotalCount = PendingCount + RunningCount + completedCount
End Property

' 获取最后运行时间（毫秒精度）
Public Property Get LastRunTimeMs() As Currency
    LastRunTimeMs = m_lastRunTimeHighRes
End Property

Private Sub Class_Initialize()
    Set m_pendingTasks = New Collection
    Set m_runningTasks = New Collection
    Set m_completedTasks = New Collection
    
    ' 初始化队列配置
    maxConcurrent = 1
    minInterval = 1000  ' 默认1秒间隔
    DefaultTimeoutSeconds = 300
    
    ' 初始化高精度计时器
    InitializeHighResTimer
    
    ' 设置初始时间为当前时间
    m_lastRunTimeHighRes = GetHighResTimeMs()
    
    Debug.Print "[" & queueName & "] 队列初始化完成，最小间隔: " & minInterval & "ms"
End Sub

' 初始化高精度计时器
Private Sub InitializeHighResTimer()
    Dim success As Long
    success = QueryPerformanceFrequency(m_performanceFrequency)
    If success = 0 Then
        Debug.Print "警告: 无法获取高精度计时器频率，使用低精度计时"
        m_performanceFrequency = 1000 ' 默认1kHz
    Else
        Debug.Print "高精度计时器频率: " & m_performanceFrequency & " Hz"
    End If
End Sub

' 获取高精度时间（毫秒）
Private Function GetHighResTimeMs() As Currency
    Dim count As Currency
    Dim success As Long
    
    success = QueryPerformanceCounter(count)
    If success = 0 Then
        ' 如果高精度计时失败，使用低精度计时
        GetHighResTimeMs = Timer * 1000
    Else
        ' 转换为毫秒
        GetHighResTimeMs = (count / m_performanceFrequency) * 1000
    End If
End Function

' 添加任务到队列
Public Sub AddTask(ByRef task As task)
    m_pendingTasks.Add task
    Debug.Print "[" & queueName & "] 添加任务: " & task.taskName & " (ID: " & task.taskID & ")"
    
    ' 记录调试信息
    Dim queueInfo As String
    queueInfo = "等待:" & PendingCount & " 运行中:" & RunningCount & " 已完成:" & completedCount
    Debug.Print "[" & queueName & "] 队列状态: " & queueInfo
End Sub

' 获取下一个可执行的任务
Public Function GetNextTask() As task
    If m_pendingTasks.count = 0 Then
        Set GetNextTask = Nothing
        Exit Function
    End If
    
    Set GetNextTask = m_pendingTasks(1)
    m_pendingTasks.Remove 1
    Debug.Print "[" & queueName & "] 获取下一个任务: " & GetNextTask.taskName
End Function

' 获取运行中的任务
Public Function GetRunningTask(ByVal Index As Long) As task
    On Error Resume Next
    If Index > 0 And Index <= m_runningTasks.count Then
        Set GetRunningTask = m_runningTasks(Index)
    Else
        Set GetRunningTask = Nothing
    End If
End Function

' 获取等待队列中的任务
Public Function GetPendingTask(ByVal Index As Long) As task
    On Error Resume Next
    If Index > 0 And Index <= m_pendingTasks.count Then
        Set GetPendingTask = m_pendingTasks(Index)
    Else
        Set GetPendingTask = Nothing
    End If
End Function

' 获取完成队列中的任务
Public Function GetCompletedTask(ByVal Index As Long) As task
    On Error Resume Next
    If Index > 0 And Index <= m_completedTasks.count Then
        Set GetCompletedTask = m_completedTasks(Index)
    Else
        Set GetCompletedTask = Nothing
    End If
End Function

' 将任务移动到运行中状态
Public Sub MoveToRunning(ByRef task As task)
    On Error Resume Next
    m_runningTasks.Add task
    task.status = tsRunning
    task.TimeStarted = Now
    
    ' 更新高精度最后运行时间
    m_lastRunTimeHighRes = GetHighResTimeMs()
    
    Debug.Print "[" & queueName & "] 任务开始运行: " & task.taskName & _
                " (运行中任务: " & RunningCount & "/" & maxConcurrent & ")"
End Sub

' 将任务移动到完成状态
Public Sub MoveToCompleted(ByRef task As task)
    On Error Resume Next
    RemoveFromRunning task.taskID
    m_completedTasks.Add task
    task.status = tsCompleted
    task.TimeCompleted = Now
    Debug.Print "[" & queueName & "] 任务完成: " & task.taskName
End Sub

' 从运行中队列移除任务
Public Sub RemoveFromRunning(ByVal taskID As Long)
    Dim i As Long
    For i = 1 To m_runningTasks.count
        If m_runningTasks(i).taskID = taskID Then
            m_runningTasks.Remove i
            Debug.Print "[" & queueName & "] 从运行中移除任务 ID: " & taskID
            Exit For
        End If
    Next i
End Sub

' 检查是否可以运行新任务（修正后的逻辑）
Public Function CanRunTask(ByRef scheduler As TaskScheduler) As Boolean
    ' 检查1: 队列级别的并发限制
    If RunningCount >= maxConcurrent Then
        Debug.Print "[" & queueName & "] 队列并发限制: " & RunningCount & "/" & maxConcurrent
        CanRunTask = False
        Exit Function
    End If
    
    ' 检查2: 任务类型级别的并发限制
    If Not scheduler.CanRunTaskType(taskType) Then
        Debug.Print "[" & queueName & "] 任务类型并发限制: " & taskType
        CanRunTask = False
        Exit Function
    End If
    
    ' 检查3: 任务类型级别的时间间隔限制
    Dim lastTaskTypeRunTime As Currency
    lastTaskTypeRunTime = scheduler.GetTaskTypeLastRunTime(taskType)
    Dim currentTime As Currency
    currentTime = GetHighResTimeMs()
    Dim timeDiff As Currency
    timeDiff = currentTime - lastTaskTypeRunTime
    
    ' 调试输出时间信息
    Debug.Print "[" & queueName & "] 任务类型时间检查 - 当前:" & Format(currentTime, "0.000") & _
                "ms, 最后任务类型运行:" & Format(lastTaskTypeRunTime, "0.000") & _
                "ms, 差值:" & Format(timeDiff, "0.000") & "ms, 最小间隔:" & minInterval & "ms"
    
    If timeDiff < minInterval Then
        Debug.Print "[" & queueName & "] 任务类型时间间隔限制: " & Format(timeDiff, "0.000") & "ms < " & minInterval & "ms"
        CanRunTask = False
        Exit Function
    End If
    
    ' 检查4: 队列级别的时间间隔限制（必须检查，无论是否有运行中任务）
    Dim queueTimeDiff As Currency
    queueTimeDiff = currentTime - m_lastRunTimeHighRes
    
    ' 调试输出队列时间信息
    Debug.Print "[" & queueName & "] 队列时间检查 - 当前:" & Format(currentTime, "0.000") & _
                "ms, 最后队列运行:" & Format(m_lastRunTimeHighRes, "0.000") & _
                "ms, 差值:" & Format(queueTimeDiff, "0.000") & "ms, 最小间隔:" & minInterval & "ms"
    
    If queueTimeDiff < minInterval Then
        Debug.Print "[" & queueName & "] 队列时间间隔限制: " & Format(queueTimeDiff, "0.000") & "ms < " & minInterval & "ms"
        CanRunTask = False
        Exit Function
    End If
    
    Debug.Print "[" & queueName & "] 可以运行任务，任务类型间隔: " & Format(timeDiff, "0.000") & "ms, 队列间隔: " & Format(queueTimeDiff, "0.000") & "ms"
    CanRunTask = True
End Function

' 获取队列统计信息
Public Function GetQueueStats() As String
    Dim stats As String
    stats = "队列[" & queueName & "] - " & _
            "等待:" & PendingCount & " " & _
            "运行中:" & RunningCount & "/" & maxConcurrent & " " & _
            "已完成:" & completedCount & " " & _
            "最后运行:" & Format(m_lastRunTimeHighRes, "0.000") & "ms"
    GetQueueStats = stats
End Function

' 清理已完成的任务（防止内存泄漏）
Public Sub CleanCompletedTasks(Optional ByVal keepCount As Long = 100)
    While m_completedTasks.count > keepCount
        m_completedTasks.Remove 1
    Wend
    
    If m_completedTasks.count > 0 Then
        Debug.Print "[" & queueName & "] 清理后保留 " & m_completedTasks.count & " 个已完成任务"
    End If
End Sub

' 重置队列状态（用于调试或重置）
Public Sub ResetQueue()
    Set m_pendingTasks = New Collection
    Set m_runningTasks = New Collection
    Set m_completedTasks = New Collection
    
    m_lastRunTimeHighRes = GetHighResTimeMs()
    
    Debug.Print "[" & queueName & "] 队列已重置"
End Sub

' 获取队列详细信息（用于调试）
Public Sub DebugQueueInfo()
    Dim info As String
    info = "=== 队列 [" & queueName & "] 详细信息 ===" & vbCrLf
    info = info & "配置: 类型=" & taskType & " 优先级=" & GetPriorityName(priority) & _
           " 最大并发=" & maxConcurrent & " 最小间隔=" & minInterval & "ms" & vbCrLf
    info = info & "状态: 等待=" & PendingCount & " 运行中=" & RunningCount & " 已完成=" & completedCount & vbCrLf
    info = info & "时间: 最后运行=" & Format(m_lastRunTimeHighRes, "0.000") & "ms" & vbCrLf
    
    ' 等待任务列表
    If PendingCount > 0 Then
        info = info & "等待任务:" & vbCrLf
        Dim i As Long
        For i = 1 To PendingCount
            Dim pendingTask As task
            Set pendingTask = GetPendingTask(i)
            If Not pendingTask Is Nothing Then
                info = info & "  " & pendingTask.taskName & " (ID:" & pendingTask.taskID & ")" & vbCrLf
            End If
        Next i
    End If
    
    ' 运行中任务列表
    If RunningCount > 0 Then
        info = info & "运行中任务:" & vbCrLf
        For i = 1 To RunningCount
            Dim runningTask As task
            Set runningTask = GetRunningTask(i)
            If Not runningTask Is Nothing Then
                info = info & "  " & runningTask.taskName & " (ID:" & runningTask.taskID & _
                       " 运行时间:" & Format(runningTask.GetRunningTime(), "0.0") & "s)" & vbCrLf
            End If
        Next i
    End If
    
    Debug.Print info
End Sub

' 辅助函数：获取优先级名称
Private Function GetPriorityName(ByVal priority As taskPriority) As String
    Select Case priority
        Case tpHigh: GetPriorityName = "高"
        Case tpMedium: GetPriorityName = "中"
        Case tpLow: GetPriorityName = "低"
        Case Else: GetPriorityName = "未知"
    End Select
End Function

' 强制更新最后运行时间（用于特殊情况）
Public Sub ForceUpdateLastRunTime()
    m_lastRunTimeHighRes = GetHighResTimeMs()
    Debug.Print "[" & queueName & "] 强制更新最后运行时间: " & Format(m_lastRunTimeHighRes, "0.000") & "ms"
End Sub

