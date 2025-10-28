Attribute VB_Name = "M_SchedulerUtils"
Option Explicit
' 任务调度器工具函数


Public Function CanRunTaskType( _
    ByRef scheduler As TaskScheduler, _
    ByVal taskType As String) As Boolean
    
    Dim currentRunning As Long
    Dim maxConcurrent As Long
    
    currentRunning = scheduler.GetTaskTypeRunningCount(taskType)
    maxConcurrent = scheduler.GetTaskTypeMaxConcurrent(taskType)
    
    If currentRunning >= maxConcurrent Then
        Debug.Print ">>> 任务类型 '" & taskType & "' 并发限制: " & currentRunning & "/" & maxConcurrent
        CanRunTaskType = False
    Else
        CanRunTaskType = True
    End If
End Function

Public Function GetPriorityName(ByVal priority As taskPriority) As String
    Select Case priority
        Case tpHigh: GetPriorityName = "高"
        Case tpMedium: GetPriorityName = "中"
        Case tpLow: GetPriorityName = "低"
        Case Else: GetPriorityName = "未知"
    End Select
End Function

Public Function GetQueueName(ByVal taskType As String, ByVal priority As taskPriority) As String
    Dim priorityNames(2) As String
    priorityNames(tpHigh) = "高"
    priorityNames(tpMedium) = "中"
    priorityNames(tpLow) = "低"
    
    GetQueueName = taskType & "-" & priorityNames(priority)
End Function

Public Function GetDefaultTimeout(ByRef taskTypeConfig As Collection, ByVal taskType As String) As Long
    On Error Resume Next
    Dim config As Collection
    Set config = taskTypeConfig(taskType)
    If Err.Number = 0 Then
        GetDefaultTimeout = config("timeoutSeconds")
    Else
        GetDefaultTimeout = 300
    End If
End Function

Public Sub LogSchedulerStatus(ByRef queues As Collection)
    Dim totalPending As Long
    Dim totalRunning As Long
    Dim totalCompleted As Long
    Dim queue As TaskQueue
    
    totalPending = 0
    totalRunning = 0
    totalCompleted = 0
    
    For Each queue In queues
        totalPending = totalPending + queue.PendingCount
        totalRunning = totalRunning + queue.RunningCount
        totalCompleted = totalCompleted + queue.completedCount
    Next queue
    
    Debug.Print ">>> 调度器状态 - 等待:" & totalPending & " 运行中:" & totalRunning & " 已完成:" & totalCompleted
    
    ' 输出各任务类型状态
    Dim taskTypes As New Collection
    Dim taskType As Variant
    
    ' 收集所有任务类型
    For Each queue In queues
        On Error Resume Next
        taskTypes.Add queue.taskType, queue.taskType
    Next queue
    
    ' 输出每个任务类型的详细状态
    For Each taskType In taskTypes
        Dim typeRunning As Long
        Dim typeMaxConcurrent As Long
        typeRunning = GetTaskTypeRunningCountFromQueues(queues, taskType)
        typeMaxConcurrent = GetTaskTypeMaxConcurrentFromQueues(queues, taskType)
        
        Debug.Print ">>>   类型 '" & taskType & "': " & typeRunning & "/" & typeMaxConcurrent & " 运行中"
    Next taskType
End Sub

Private Function GetTaskTypeRunningCountFromQueues(ByRef queues As Collection, ByVal taskType As String) As Long
    Dim queue As TaskQueue
    Dim totalRunning As Long
    
    totalRunning = 0
    For Each queue In queues
        If queue.taskType = taskType Then
            totalRunning = totalRunning + queue.RunningCount
        End If
    Next queue
    
    GetTaskTypeRunningCountFromQueues = totalRunning
End Function

Private Function GetTaskTypeMaxConcurrentFromQueues(ByRef queues As Collection, ByVal taskType As String) As Long
    Dim queue As TaskQueue
    
    For Each queue In queues
        If queue.taskType = taskType Then
            GetTaskTypeMaxConcurrentFromQueues = queue.maxConcurrent
            Exit Function
        End If
    Next queue
    
    GetTaskTypeMaxConcurrentFromQueues = 3 ' 默认值
End Function

