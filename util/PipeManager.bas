' PipeManager.bas - 简化的管道与 ffmpeg 输出解析封装（含 Mock 支持）
Option Explicit

Public Type PipeReadResult
    Success As Boolean
    Output As String
End Type

' 配置: 若 UseMockMode = True，则不启动外部进程，仅返回 sample data
Public UseMockMode As Boolean
Public MockOutputSample As String

Public Sub InitializePipeManager()
    On Error Resume Next
    UseMockMode = True ' 默认 Mock 模式，重构后可从配置加载
    MockOutputSample = "bitrate=1200k fps=25"
End Sub

Public Function ReadPipeOutput(ByVal procHandle As Long) As PipeReadResult
    Dim res As PipeReadResult
    On Error Resume Next
    If UseMockMode Then
        res.Success = True
        res.Output = MockOutputSample
        ReadPipeOutput = res
        Exit Function
    End If
    ' 真实实现应读取管道句柄 procHandle 的输出，这里只是占位
    res.Success = False
    res.Output = vbNullString
    ReadPipeOutput = res
End Function

Public Sub ClosePipe(ByVal procHandle As Long)
    On Error Resume Next
    ' 真实实现关闭句柄
End Sub

' End of PipeManager.bas