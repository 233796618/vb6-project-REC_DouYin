Attribute VB_Name = "M_LiveroomManager"
Option Explicit

' LiveroomManager 定时器相关

Private g_liveroomManagerTimerID As Long
Private g_liveroomManagerInstance As LiveroomManager

' LiveroomManager 定时器回调
Public Sub LiveroomManagerTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error Resume Next
    'Debug.Print ">>> LiveroomManager 定时器触发: " & Format(Now, "hh:mm:ss")
    
    If Not g_liveroomManagerInstance Is Nothing Then
        g_liveroomManagerInstance.ProcessPolling
    End If
End Sub

' 设置 LiveroomManager 实例
Public Sub SetLiveroomManagerInstance(ByRef manager As LiveroomManager)
    Set g_liveroomManagerInstance = manager
    Debug.Print "LiveroomManager 实例已设置"
End Sub

' 启动 LiveroomManager 定时器
Public Sub StartLiveroomManagerTimer()
    If g_liveroomManagerTimerID = 0 Then
        g_liveroomManagerTimerID = SetTimer(0, 0, 1000, AddressOf LiveroomManagerTimerProc) ' 1秒间隔
        If g_liveroomManagerTimerID <> 0 Then
            Debug.Print "LiveroomManager 定时器启动成功 (ID: " & g_liveroomManagerTimerID & ")"
        Else
            Debug.Print "错误：无法启动 LiveroomManager 定时器"
        End If
    End If
End Sub

' 停止 LiveroomManager 定时器
Public Sub StopLiveroomManagerTimer()
    If g_liveroomManagerTimerID <> 0 Then
        KillTimer 0, g_liveroomManagerTimerID
        g_liveroomManagerTimerID = 0
        Debug.Print "LiveroomManager 定时器已停止"
    End If
End Sub

' 清理 LiveroomManager 定时器
Public Sub CleanupLiveroomManagerTimer()
    StopLiveroomManagerTimer
    Set g_liveroomManagerInstance = Nothing
    Debug.Print "LiveroomManager 定时器已清理"
End Sub