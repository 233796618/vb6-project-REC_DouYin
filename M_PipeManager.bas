Attribute VB_Name = "M_PipeManager"
Option Explicit



' 在 M_TaskScheduler.bas 或其他适当的模块中添加
Public Declare Function SetNamedPipeHandleState Lib "kernel32" ( _
    ByVal hNamedPipe As Long, _
    ByVal lpMode As Long, _
    ByVal lpMaxCollectionCount As Long, _
    ByVal lpCollectDataTimeout As Long) As Long

Public Declare Function GetNamedPipeHandleState Lib "kernel32" ( _
    ByVal hNamedPipe As Long, _
    lpState As Long, _
    lpCurInstances As Long, _
    lpMaxCollectionCount As Long, _
    lpCollectDataTimeout As Long, _
    lpUserName As String, _
    ByVal nMaxUserNameSize As Long) As Long

' 管道模式常量
Public Const PIPE_NOWAIT As Long = &H1
Public Const PIPE_READMODE_BYTE As Long = &H0
Public Const PIPE_READMODE_MESSAGE As Long = &H2


' 需要添加的声明
Private Declare Function SetHandleInformation Lib "kernel32" _
    (ByVal hObject As Long, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
    
Private Const HANDLE_FLAG_INHERIT As Long = &H1

' 管道管理函数
Public Function CreatePipeWithSecurity( _
    ByRef hReadPipe As Long, _
    ByRef hWritePipe As Long, _
    Optional ByVal bInheritRead As Boolean = False, _
    Optional ByVal bInheritWrite As Boolean = False) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim sa As SECURITY_ATTRIBUTES
    sa.nLength = Len(sa)
    sa.lpSecurityDescriptor = 0
    sa.bInheritHandle = IIf(bInheritRead Or bInheritWrite, 1, 0)
    
    Dim result As Long
    result = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    
    If result <> 0 Then
        ' 设置管道继承属性
        If Not bInheritRead Then
            Call SetHandleInformation(hReadPipe, HANDLE_FLAG_INHERIT, 0)
        End If
        If Not bInheritWrite Then
            Call SetHandleInformation(hWritePipe, HANDLE_FLAG_INHERIT, 0)
        End If
        
        CreatePipeWithSecurity = True
    Else
        CreatePipeWithSecurity = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "创建管道时出错: " & Err.Description
    CreatePipeWithSecurity = False
End Function

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

Public Function CleanOutputText(ByVal text As String) As String
    On Error Resume Next
    
    If Len(text) = 0 Then
        CleanOutputText = ""
        Exit Function
    End If
    
    Dim result As String
    result = text
    
    ' 移除回车符
    result = Replace(result, vbCr, "")
    
    ' 移除多个连续的空格
    While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Wend
    
    ' 移除开头和结尾的空格
    result = Trim(result)
    
    CleanOutputText = result
End Function

Public Function WriteToPipe(ByVal hPipe As Long, ByVal data As String) As Boolean
    On Error GoTo ErrorHandler
    
    If hPipe = 0 Then
        WriteToPipe = False
        Exit Function
    End If
    
    Dim bytesWritten As Long
    Dim success As Long
    
    success = WriteFile(hPipe, ByVal data, Len(data), bytesWritten, ByVal 0&)
    
    If success <> 0 And bytesWritten = Len(data) Then
        WriteToPipe = True
    Else
        WriteToPipe = False
        Debug.Print "写入管道失败，错误代码: " & Err.LastDllError
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "写入管道时出错: " & Err.Description
    WriteToPipe = False
End Function


