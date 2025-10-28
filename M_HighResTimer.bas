Attribute VB_Name = "M_HighResTimer"
Option Explicit

Private m_performanceFrequency As Currency
Private m_timerInitialized As Boolean

Public Sub InitializeHighResTimer()
    If m_timerInitialized Then Exit Sub
    
    Dim success As Long
    success = QueryPerformanceFrequency(m_performanceFrequency)
    
    If success = 0 Then
        Debug.Print "警告: 无法获取高精度计时器频率，使用低精度计时"
        m_performanceFrequency = 1000 ' 默认1kHz
    Else
        Debug.Print "高精度计时器频率: " & m_performanceFrequency & " Hz"
    End If
    
    m_timerInitialized = True
End Sub

Public Function GetHighResTimeMs() As Currency
    If Not m_timerInitialized Then
        InitializeHighResTimer
    End If
    
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

Public Function GetFullHighResTimestamp() As String
    Dim ms As Currency
    Dim currentTime As Date
    Dim milliseconds As Long
    
    ms = GetHighResTimeMs()
    currentTime = Now ' 获取当前完整时间
    
    ' 提取毫秒部分并组合
    milliseconds = ms Mod 1000
    currentTime = DateAdd("s", -Second(currentTime), currentTime) ' 移除秒数
    currentTime = DateAdd("s", ms / 1000, currentTime) ' 添加高精度秒数
    
    GetFullHighResTimestamp = Format(currentTime, "yyyy-mm-dd hh:mm:ss") & _
                              "." & Format(milliseconds, "000")
End Function

Public Function GetTimeDiffMs(ByVal startTime As Currency, ByVal endTime As Currency) As Currency
    GetTimeDiffMs = endTime - startTime
End Function

Public Function IsTimeIntervalPassed( _
    ByVal lastTime As Currency, _
    ByVal intervalMs As Long, _
    Optional ByRef currentTime As Currency = 0) As Boolean
    
    If currentTime = 0 Then
        currentTime = GetHighResTimeMs()
    End If
    
    IsTimeIntervalPassed = (currentTime - lastTime) >= intervalMs
End Function

