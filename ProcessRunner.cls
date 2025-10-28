VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 运行模式枚举
Public Enum ProcessRunnerMode
    prmSimple = 0
    prmPipe = 1
End Enum

Private m_task As task
Private m_isRunning As Boolean
Private m_parent As TaskScheduler
Private m_runnerMode As ProcessRunnerMode

Public Event TaskCompleted(ByRef task As task)
Public Event TaskFailed(ByRef task As task, ByVal errorMsg As String)

Public Property Set Parent(ByRef scheduler As TaskScheduler)
    Set m_parent = scheduler
End Property

Public Property Get Parent() As TaskScheduler
    Set Parent = m_parent
End Property

Public Property Set CurrentTask(ByRef task As task)
    Set m_task = task
    ' 根据任务设置运行模式
    If task.monitorPipe Then
        m_runnerMode = prmPipe
    Else
        m_runnerMode = prmSimple
    End If
End Property

Public Property Get CurrentTask() As task
    Set CurrentTask = m_task
End Property

Public Property Get IsRunning() As Boolean
    IsRunning = m_isRunning
End Property

Public Property Get RunnerMode() As ProcessRunnerMode
    RunnerMode = m_runnerMode
End Property

Public Sub ExecuteTask()
    If m_task Is Nothing Then
        Debug.Print "错误: 任务为空"
        Exit Sub
    End If
    
    m_isRunning = True
    m_task.status = tsRunning
    m_task.TimeStarted = Now
    
    Debug.Print "开始执行任务[" & GetModeName() & "]: " & m_task.taskName
    Debug.Print GetFullHighResTimestamp() & "开始执行任务[" & GetModeName() & "]: " & m_task.taskName
    
    Select Case m_runnerMode
        Case prmSimple
            ExecuteSimple
        Case prmPipe
            ExecuteWithPipe
    End Select
End Sub

Private Sub ExecuteSimple()
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim success As Long
    Dim fullCommandLine As String
    Dim workingDir As String
    
    On Error GoTo ErrorHandler
    
    fullCommandLine = BuildCommandLine()
    workingDir = GetWorkingDirectory()
    
    Debug.Print "简单模式命令行: " & fullCommandLine
    Debug.Print "工作目录: " & workingDir
    
    ' 设置启动信息
    si.cb = Len(si)
    si.dwFlags = 0
    si.wShowWindow = SW_HIDE
    
    pi.hProcess = 0
    pi.hThread = 0
    pi.dwProcessId = 0
    pi.dwThreadId = 0
    
    success = CreateProcessA( _
        vbNullString, _
        fullCommandLine, _
        0, _
        0, _
        0, _
        CREATE_NO_WINDOW, _
        0, _
        workingDir, _
        si, _
        pi)
    
    If success = 0 Then
        Debug.Print "创建进程失败，错误代码: " & Err.LastDllError
        GoTo ErrorHandler
    End If
    
    Debug.Print "进程创建成功，进程ID: " & pi.dwProcessId
    
    ' 保存进程信息
    m_task.hProcess = pi.hProcess
    m_task.ProcessId = pi.dwProcessId
    m_task.hPipeRead = 0
    m_task.hPipeWrite = 0
    
    ' 关闭线程句柄
    CloseHandle pi.hThread
    
    ' 通知父调度器
    NotifyTaskStarted
    
    m_isRunning = False
    Debug.Print "简单模式任务执行完成"
    Exit Sub
    
ErrorHandler:
    Debug.Print "简单模式执行出错: " & Err.Description
    CleanupProcess pi
    NotifyFailure "创建进程失败，错误代码: " & Err.LastDllError
End Sub

Private Sub ExecuteWithPipe()
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim sa As SECURITY_ATTRIBUTES
    Dim hStdOutRead As Long, hStdOutWrite As Long
    Dim hStdInRead As Long, hStdInWrite As Long
    Dim success As Long
    Dim commandLine As String
    Dim workingDir As String
    
    On Error GoTo ErrorHandler
    
    commandLine = BuildCommandLine()
    workingDir = GetWorkingDirectory()
    
    Debug.Print "管道模式命令行: " & commandLine
    Debug.Print "工作目录: " & workingDir
    
    ' 设置安全属性
    sa.nLength = Len(sa)
    sa.lpSecurityDescriptor = 0
    sa.bInheritHandle = 1
    
    ' 创建输出管道
    success = CreatePipe(hStdOutRead, hStdOutWrite, sa, 0)
    If success = 0 Then
        Debug.Print "创建输出管道失败，错误代码: " & Err.LastDllError
        GoTo ErrorHandler
    End If
    
    ' 创建输入管道
    success = CreatePipe(hStdInRead, hStdInWrite, sa, 0)
    If success = 0 Then
        Debug.Print "创建输入管道失败，错误代码: " & Err.LastDllError
        GoTo ErrorHandler
    End If
    
    ' 设置启动信息
    si.cb = Len(si)
    si.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    si.wShowWindow = SW_HIDE
    si.hStdOutput = hStdOutWrite
    si.hStdError = hStdOutWrite
    si.hStdInput = hStdInRead
    
    pi.hProcess = 0
    pi.hThread = 0
    pi.dwProcessId = 0
    pi.dwThreadId = 0
    
    success = CreateProcessA( _
        vbNullString, _
        commandLine, _
        0, _
        0, _
        1, _
        CREATE_NO_WINDOW, _
        0, _
        workingDir, _
        si, _
        pi)
    
    If success = 0 Then
        Debug.Print "创建管道进程失败，错误代码: " & Err.LastDllError
        GoTo ErrorHandler
    End If
    
    Debug.Print "管道进程创建成功，进程ID: " & pi.dwProcessId
    
    ' 关闭不需要的句柄
    CloseHandle hStdOutWrite
    CloseHandle hStdInRead
    CloseHandle pi.hThread
    
    ' 保存句柄到任务
    m_task.hProcess = pi.hProcess
    m_task.hPipeRead = hStdOutRead
    m_task.hPipeWrite = hStdInWrite
    m_task.ProcessId = pi.dwProcessId
    
    Debug.Print "保存进程句柄: 进程=" & m_task.hProcess & _
                ", 输出管道=" & m_task.hPipeRead & _
                ", 输入管道=" & m_task.hPipeWrite & _
                ", 进程ID=" & m_task.ProcessId
    
    ' 通知父调度器
    NotifyTaskStarted
    
    m_isRunning = False
    Debug.Print "管道模式任务执行完成"
    Exit Sub
    
ErrorHandler:
    Debug.Print "管道模式执行出错: " & Err.Description
    CleanupProcess pi
    CleanupPipes hStdOutRead, hStdOutWrite, hStdInRead, hStdInWrite
    NotifyFailure "创建管道进程失败，错误代码: " & Err.LastDllError
End Sub

' 辅助方法
Private Function BuildCommandLine() As String
    BuildCommandLine = m_task.commandLine
    If Len(m_task.parameters) > 0 Then
        BuildCommandLine = BuildCommandLine & " " & m_task.parameters
    End If
End Function

Private Function GetWorkingDirectory() As String
    GetWorkingDirectory = m_task.workingDirectory
    If GetWorkingDirectory = "" Then
        GetWorkingDirectory = CurDir
    End If
    
    If Dir(GetWorkingDirectory, vbDirectory) = "" Then
        Debug.Print "警告: 工作目录不存在，使用当前目录: " & GetWorkingDirectory
        GetWorkingDirectory = CurDir
    End If
End Function

Private Function GetModeName() As String
    Select Case m_runnerMode
        Case prmSimple: GetModeName = "简单模式"
        Case prmPipe: GetModeName = "管道模式"
        Case Else: GetModeName = "未知模式"
    End Select
End Function

' 清理方法
Private Sub CleanupProcess(pi As PROCESS_INFORMATION)
    If pi.hProcess <> 0 Then CloseHandle pi.hProcess
    If pi.hThread <> 0 Then CloseHandle pi.hThread
End Sub

Private Sub CleanupPipes(hStdOutRead As Long, hStdOutWrite As Long, hStdInRead As Long, hStdInWrite As Long)
    If hStdOutRead <> 0 Then CloseHandle hStdOutRead
    If hStdOutWrite <> 0 Then CloseHandle hStdOutWrite
    If hStdInRead <> 0 Then CloseHandle hStdInRead
    If hStdInWrite <> 0 Then CloseHandle hStdInWrite
End Sub

' 通知方法
Private Sub NotifyTaskStarted()
    If Not m_parent Is Nothing Then
        m_parent.OnTaskStarted m_task
    End If
End Sub

Private Sub NotifyCompletion()
    m_isRunning = False
    m_task.status = tsCompleted
    m_task.TimeCompleted = Now
    RaiseEvent TaskCompleted(m_task)
    
    If Not m_parent Is Nothing Then
        m_parent.OnTaskCompleted m_task
    End If
End Sub

Private Sub NotifyFailure(ByVal errorMsg As String)
    m_isRunning = False
    m_task.status = tsFailed
    m_task.ErrorData = errorMsg
    RaiseEvent TaskFailed(m_task, errorMsg)
    
    If Not m_parent Is Nothing Then
        m_parent.OnTaskFailed m_task, errorMsg
    End If
End Sub

