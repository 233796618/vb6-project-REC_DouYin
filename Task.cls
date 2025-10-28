VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 任务状态枚举
Public Enum TaskStatus
    tsPending = 0
    tsRunning = 1
    tsCompleted = 2
    tsFailed = 3
    tsTimeout = 4
End Enum

' 优先级枚举
Public Enum TaskPriority
    tpHigh = 0
    tpMedium = 1
    tpLow = 2
End Enum

' 任务属性
Public taskID As Long
Public taskType As String
Public taskName As String
Public commandLine As String
Public parameters As String
Public priority As TaskPriority
Public web_rid As String
Public workingDirectory As String

' 时间记录
Public TimeAdded As Date
Public TimeStarted As Date
Public TimeCompleted As Date
Public timeoutSeconds As Long

' 状态
Public status As TaskStatus
Public exitCode As Long
Public outputData As String
Public ErrorData As String

' 管道监控
Public monitorPipe As Boolean

' 进程和管道句柄
Public hProcess As Long
Public hPipeRead As Long     ' 输出管道（读取）
Public hPipeWrite As Long    ' 输入管道（写入）- 新增
Public ProcessId As Long

Private Sub Class_Initialize()
    ' 使用更可靠的方式生成 taskID
    taskID = GetTickCount() + Int(Rnd() * 10000)
    TimeAdded = Now
    status = tsPending
    priority = tpMedium
    timeoutSeconds = 300
    monitorPipe = False
    taskType = "default"
    parameters = ""
    web_rid = ""
    workingDirectory = ""
    taskName = ""  ' 初始化为空，由调用者设置
    
    ' 初始化句柄
    hProcess = 0
    hPipeRead = 0
    hPipeWrite = 0  ' 新增
    ProcessId = 0
End Sub

' 检查任务是否超时
Public Function IsTimeout() As Boolean
    If status <> tsRunning Then
        IsTimeout = False
        Exit Function
    End If
    
    If timeoutSeconds <= 0 Then
        IsTimeout = False
        Exit Function
    End If
    
    Dim elapsedSeconds As Double
    elapsedSeconds = (Now - TimeStarted) * 24 * 60 * 60
    
    IsTimeout = (elapsedSeconds > timeoutSeconds)
End Function

' 获取运行时间（秒）
Public Function GetRunningTime() As Double
    If TimeStarted = 0 Then
        GetRunningTime = 0
    ElseIf TimeCompleted <> 0 Then
        GetRunningTime = (TimeCompleted - TimeStarted) * 24 * 60 * 60
    Else
        GetRunningTime = (Now - TimeStarted) * 24 * 60 * 60
    End If
End Function

' 清理资源
Public Sub Cleanup()
    On Error Resume Next
    If hProcess <> 0 Then
        CloseHandle hProcess
        hProcess = 0
    End If
    If hPipeRead <> 0 Then
        CloseHandle hPipeRead
        hPipeRead = 0
    End If
    If hPipeWrite <> 0 Then    ' 新增
        CloseHandle hPipeWrite
        hPipeWrite = 0
    End If
    ProcessId = 0
End Sub
