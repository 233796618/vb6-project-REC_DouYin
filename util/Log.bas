' Log.bas - 简单日志模块
Option Explicit

Public Enum LogLevel
    LOG_DEBUG = 0
    LOG_INFO = 1
    LOG_WARN = 2
    LOG_ERROR = 3
End Enum

Public g_LogLevel As LogLevel

Public Sub LogMessage(level As LogLevel, ByVal msg As String)
    If level < g_LogLevel Then Exit Sub
    Dim prefix As String
    Select Case level
        Case LOG_DEBUG: prefix = "DEBUG"
        Case LOG_INFO: prefix = "INFO"
        Case LOG_WARN: prefix = "WARN"
        Case LOG_ERROR: prefix = "ERROR"
        Case Else: prefix = "LOG"
    End Select
    Debug.Print "[" & prefix & "] " & msg
End Sub

Public Sub InitLog()
    g_LogLevel = LOG_DEBUG
End Sub

' End of Log.bas