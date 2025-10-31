' ProcessRunner.cls - 启动外部进程并管理输入输出（简化版）
Option Explicit

Private m_ProcessHandle As Long

Public Sub StartProcess(ByVal exePath As String, ByVal args As String)
    On Error Resume Next
    ' Mock 模式或实际启动外部进程
    If UseMockMode Then
        m_ProcessHandle = 0
        Debug.Print "[ProcessRunner] Mock start: " & exePath & " " & args
        Exit Sub
    End If
    ' TODO: 实现 CreateProcess + pipe 的启动逻辑
    Err.Raise vbObjectError + 1000, "ProcessRunner", "未实现真实启动逻辑"
End Sub

Public Function ReadOutput() As String
    On Error Resume Next
    If UseMockMode Then
        ReadOutput = MockOutputSample
    Else
        ReadOutput = ""
    End If
End Function

Public Sub StopProcess()
    On Error Resume Next
    If UseMockMode Then
        Debug.Print "[ProcessRunner] Mock stop"
        Exit Sub
    End If
    ' TODO: 实现真实停止逻辑
End Sub

Public Property Get ProcessHandle() As Long
    ProcessHandle = m_ProcessHandle
End Property

' End of ProcessRunner.cls