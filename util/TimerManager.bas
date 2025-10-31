' TimerManager.bas - 统一定时器接口（使用 SetTimer/WM_TIMER 或高精度 Timer）
Option Explicit

Public Sub StartHighResTimer(ByVal intervalMs As Long)
    On Error Resume Next
    ' TODO: 实现高精度定时器或用现有 SetTimer
    Debug.Print "[TimerManager] StartHighResTimer: " & intervalMs
End Sub

Public Sub StopHighResTimer()
    On Error Resume Next
    Debug.Print "[TimerManager] StopHighResTimer"
End Sub

' End of TimerManager.bas