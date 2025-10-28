VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskSchedulerTaskManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 运行器集合
Private m_runners As Collection
Private m_processTaskMap As Collection

' 父调度器引用
Private m_parent As TaskScheduler

' 事件定义 - 用于通知父调度器
Public Event TaskCompleted(ByRef task As task)
Public Event TaskFailed(ByRef task As task, ByVal errorMsg As String)
Public Event TaskTimeout(ByRef task As task)
Public Event TaskOutput(ByRef task As task, ByVal outputData As String, ByVal isErrorOutput As Boolean)

Private Sub Class_Initialize()
    Set m_runners = New Collection
    Set m_processTaskMap = New Collection
    Debug.Print "TaskSchedulerTaskManager 初始化完成"
End Sub

' 设置父调度器
Public Sub SetParent(ByRef parentScheduler As TaskScheduler)
    Set m_parent = parentScheduler
    Debug.Print ">>> 任务管理器父调度器已设置"
End Sub

' ========== 运行器管理方法 ==========

' 创建运行器
Public Function CreateRunner(ByVal monitorPipe As Boolean) As Object
    On Error GoTo ErrorHandler
    
    Dim runner As ProcessRunner
    Set runner = New ProcessRunner
    
    Debug.Print ">>> 创建统一进程运行器，监控管道: " & monitorPipe
    
    If Not m_parent Is Nothing Then
        Set runner.Parent = m_parent
    End If
    
    Debug.Print ">>> 运行器创建完成"
    Set CreateRunner = runner
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> 创建运行器时出错: " & Err.Description
    Set CreateRunner = Nothing
End Function

' 清理完成的运行器
Public Sub CleanCompletedRunners()
    On Error Resume Next
    Dim i As Long
    i = 1
    While i <= m_runners.count
        If Not m_runners(i).IsRunning Then
            m_runners.Remove i
        Else
            i = i + 1
        End If
    Wend
End Sub

' 添加运行器
Public Sub AddRunner(ByRef runner As Object)
    On Error Resume Next
    m_runners.Add runner
End Sub

' ========== 进程和管道管理方法 ==========

' 检查完成进程
' 修改：检查完成进程 - 添加 ResetFFmpegPipe 调用
Public Sub CheckCompletedProcesses(ByRef queueManager As TaskSchedulerQueueManager)
    On Error Resume Next
    
    If queueManager Is Nothing Then
        Exit Sub
    End If
    
    Dim queues As Collection
    Set queues = queueManager.GetAllQueues()
    
    If queues Is Nothing Then
        Exit Sub
    End If
    
    Dim queue As TaskQueue
    Dim task As task
    Dim i As Long
    Dim completedTasks As New Collection
    
    ' 第一遍：收集已完成的任务
    For Each queue In queues
        For i = 1 To queue.RunningCount
            Set task = queue.GetRunningTask(i)
            If Not task Is Nothing And task.status = tsRunning Then
                If task.hProcess <> 0 Then
                    Dim exitCode As Long
                    Dim result As Long
                    result = GetExitCodeProcess(task.hProcess, exitCode)
                    
                    If result <> 0 Then
                        If exitCode <> STILL_ACTIVE Then
                            completedTasks.Add task
                        End If
                    End If
                End If
            End If
        Next i
    Next queue
    
    ' 第二遍：处理已完成的任务
    For Each task In completedTasks
        If Not task Is Nothing Then
            task.exitCode = exitCode
            task.TimeCompleted = Now
            task.status = tsCompleted
            
            Debug.Print ">>> 进程完成: " & task.taskName & " (退出代码: " & task.exitCode & ")"
            
            ' 如果是录制任务，重置管道信息
            If task.taskType = "recorder" And task.monitorPipe Then
                ResetFFmpegPipe task.web_rid
            End If
            
            ' 清理资源
            task.Cleanup
            
            ' 找到对应的队列并移动
            Dim foundQueue As TaskQueue
            Set foundQueue = FindQueueByTask(queues, task)
            If Not foundQueue Is Nothing Then
                foundQueue.MoveToCompleted task
            End If
            
            ' 触发事件 - 通过自己的事件通知父调度器
            RaiseEvent TaskCompleted(task)
        End If
    Next task
End Sub

' 新增：重置 FFmpeg 管道信息
Private Sub ResetFFmpegPipe(ByVal web_rid As String)
    On Error Resume Next
    Dim room As LiveRoom
    Set room = g_liveroomManager.GetLiveroom(web_rid)
    If Not room Is Nothing Then
        room.ResetFFmpegInfo
        Debug.Print ">>> 已重置 FFmpeg 管道信息: " & web_rid
    End If
End Sub

' 检查管道输出
Public Sub CheckPipeOutputs(ByRef queueManager As TaskSchedulerQueueManager)
    On Error Resume Next
    
    If queueManager Is Nothing Then
        Exit Sub
    End If
    
    Dim queues As Collection
    Set queues = queueManager.GetAllQueues()
    
    If queues Is Nothing Then
        Exit Sub
    End If
    
    Dim queue As TaskQueue
    Dim task As task
    Dim i As Long
    Dim hasOutput As Boolean
    
    hasOutput = False
    
    For Each queue In queues
        For i = 1 To queue.RunningCount
            Set task = queue.GetRunningTask(i)
            If Not task Is Nothing And task.status = tsRunning Then
                If task.monitorPipe And task.hPipeRead <> 0 Then
                    Dim output As String
                    output = ReadPipeOutput(task.hPipeRead)
                    
                    If Len(output) > 0 Then
                        hasOutput = True
                        
                        Debug.Print ">>> 管道输出 [" & task.taskName & "]: " & output
                        UseFFmpegPipe task.web_rid, output
                        
                        task.outputData = task.outputData & output & vbCrLf
                        ' 触发事件 - 通过自己的事件通知父调度器
                        RaiseEvent TaskOutput(task, output, False)
                    End If
                End If
            End If
        Next i
    Next queue
    
    If hasOutput Then
        Debug.Print ">>> 本轮检测到管道输出"
    End If
End Sub

' 检查任务超时
Public Sub CheckTaskTimeouts(ByRef queueManager As TaskSchedulerQueueManager)
    On Error GoTo ErrorHandler
    
    If queueManager Is Nothing Then
        Exit Sub
    End If
    
    Dim queues As Collection
    Set queues = queueManager.GetAllQueues()
    
    If queues Is Nothing Then
        Exit Sub
    End If
    
    Dim queue As TaskQueue
    Dim task As task
    Dim i As Long
    
    For Each queue In queues
        For i = 1 To queue.RunningCount
            Set task = queue.GetRunningTask(i)
            If Not task Is Nothing Then
                If task.status = tsRunning Then
                    Debug.Print ">>> 检查运行中任务: " & task.taskName & " (运行时间: " & Format(task.GetRunningTime(), "0.0") & "秒, 超时: " & task.timeoutSeconds & "秒)"
                End If
                
                If task.IsTimeout() Then
                    Debug.Print ">>> ！！！任务超时: " & task.taskName & " (运行时间: " & Format(task.GetRunningTime(), "0.0") & "秒)！！！"
                    
                    If TerminateTaskProcess(task) Then
                        Debug.Print ">>> 成功终止超时任务的进程"
                    End If
                    
                    task.status = tsTimeout
                    task.TimeCompleted = Now
                    task.ErrorData = "任务执行超时（" & task.timeoutSeconds & "秒），进程已被终止"
                    task.exitCode = -1
                    
                    task.Cleanup
                    queue.RemoveFromRunning task.taskID
                    
                    ' 触发事件 - 通过自己的事件通知父调度器
                    RaiseEvent TaskTimeout(task)
                    RaiseEvent TaskFailed(task, task.ErrorData)
                End If
            End If
        Next i
    Next queue
    
    Exit Sub
    
ErrorHandler:
    Debug.Print ">>> 检查任务超时时出错: " & Err.Description
End Sub

' ========== 辅助方法 ==========

' 终止任务进程
Private Function TerminateTaskProcess(ByRef task As task) As Boolean
    On Error GoTo ErrorHandler
    
    If task.hProcess <> 0 Then
        Debug.Print ">>> 找到进程句柄: " & task.hProcess & "，开始终止..."
        
        Dim result As Long
        result = TerminateProcess(task.hProcess, 1)
        
        If result <> 0 Then
            Debug.Print ">>> 进程终止成功: 句柄 " & task.hProcess
            TerminateTaskProcess = True
        Else
            Debug.Print ">>> 进程终止失败，错误代码: " & Err.LastDllError
            TerminateTaskProcess = False
        End If
    Else
        Debug.Print ">>> 未找到任务的进程句柄，任务: " & task.taskName
        TerminateTaskProcess = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> 终止进程时出错: " & Err.Description
    TerminateTaskProcess = False
End Function

' 根据任务找到队列
Private Function FindQueueByTask(ByRef queues As Collection, ByRef taskToFind As task) As TaskQueue
    On Error Resume Next
    Dim queue As TaskQueue
    For Each queue In queues
        Dim i As Long
        For i = 1 To queue.RunningCount
            Dim CurrentTask As task
            Set CurrentTask = queue.GetRunningTask(i)
            If Not CurrentTask Is Nothing Then
                If CurrentTask.taskID = taskToFind.taskID Then
                    Set FindQueueByTask = queue
                    Exit Function
                End If
            End If
        Next i
    Next queue
    Set FindQueueByTask = Nothing
End Function

' 读取管道输出
Public Function ReadPipeOutput(ByVal hPipe As Long) As String
    On Error Resume Next
    
    Dim bytesAvailable As Long
    Dim success As Long
    
    ' 检查是否有数据可读
    success = PeekNamedPipe(hPipe, ByVal 0&, 0, ByVal 0&, bytesAvailable, ByVal 0&)
    
    If success <> 0 And bytesAvailable > 0 Then
        Dim buffer() As Byte
        Dim bytesRead As Long
        Dim totalOutput As String
        
        ReDim buffer(0 To bytesAvailable - 1) As Byte
        
        success = ReadFile(hPipe, ByVal VarPtr(buffer(0)), bytesAvailable, bytesRead, ByVal 0&)
        
        If success <> 0 And bytesRead > 0 Then
            ' 将字节数据转换为字符串
            totalOutput = StrConv(buffer, vbUnicode)
            totalOutput = Left$(totalOutput, bytesRead)
            
            ' 清理输出
            totalOutput = CleanOutputText(totalOutput)
            
            ReadPipeOutput = totalOutput
        Else
            Debug.Print "读取管道失败，错误代码: " & Err.LastDllError
        End If
    End If
End Function
