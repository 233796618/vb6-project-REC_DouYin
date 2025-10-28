VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LiveroomManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 直播间集合
Private m_allLiverooms As Collection      ' 所有直播间
Private m_enabledLiverooms As Collection  ' 启用中的直播间
Private m_disabledLiverooms As Collection ' 禁用中的直播间

' Timer相关
Private m_timerID As Long
Private m_timerRunning As Boolean
Private Const TIMER_INTERVAL As Long = 5000 ' 1秒轮询间隔

' 调度器引用
Private m_scheduler As TaskScheduler

' 事件定义
Public Event LiveroomAdded(ByRef room As LiveRoom)
Public Event LiveroomRemoved(ByRef room As LiveRoom)
Public Event LiveroomEnabled(ByRef room As LiveRoom)
Public Event LiveroomDisabled(ByRef room As LiveRoom)
Public Event ParseTaskTriggered(ByRef room As LiveRoom)
Public Event RecordTaskTriggered(ByRef room As LiveRoom)
Public Event PollingCompleted()  ' 新增：轮询完成事件

Private Sub Class_Initialize()
    Set m_allLiverooms = New Collection
    Set m_enabledLiverooms = New Collection
    Set m_disabledLiverooms = New Collection
    
    m_timerID = 0
    m_timerRunning = False
    
    Debug.Print "LiveroomManager 初始化完成"
End Sub

Private Sub Class_Terminate()
    StopTimer
    Set m_allLiverooms = Nothing
    Set m_enabledLiverooms = Nothing
    Set m_disabledLiverooms = Nothing
    Set m_scheduler = Nothing
    Debug.Print "LiveroomManager 已清理"
End Sub

' 从数组添加直播间
Public Sub AddLiveroomsFromArray(ByRef web_ridArray() As String, _
                                Optional ByVal checkIntervalMinutes As Long = 5)
    On Error Resume Next
    
    Dim i As Long
    For i = LBound(web_ridArray) To UBound(web_ridArray)
        AddLiveroomSimple web_ridArray(i), "", checkIntervalMinutes
    Next i
    
    Debug.Print "LiveroomManager: 从数组添加了 " & (UBound(web_ridArray) - LBound(web_ridArray) + 1) & " 个直播间"
End Sub

' 修改：添加直播间（简化版本）- 修复递归调用
Public Function AddLiveroomSimple(ByVal web_rid As String, _
                                  Optional ByVal AnchorName As String = "", _
                                  Optional ByVal checkIntervalMinutes As Long = 5, _
                                  Optional ByVal checkAllTime As Boolean = True, _
                                  Optional ByVal watchTimeStart As String = "", _
                                  Optional ByVal watchTimeEnd As String = "", _
                                  Optional ByVal taskPriority As taskPriority = tpMedium, _
                                  Optional ByVal enabled As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    ' 检查是否已存在
    If LiveroomExists(web_rid) Then
        Debug.Print "LiveroomManager: 直播间已存在 - " & web_rid
        AddLiveroomSimple = True
        Exit Function
    End If
    
    ' 创建新的直播间对象
    Dim room As New LiveRoom
    room.web_rid = web_rid
    room.AnchorName = IIf(AnchorName = "", "主播_" & web_rid, AnchorName)
    room.checkIntervalMinutes = checkIntervalMinutes
    room.checkAllTime = checkAllTime
    room.watchTimeStart = watchTimeStart
    room.watchTimeEnd = watchTimeEnd
    room.taskPriority = taskPriority
    room.enabled = enabled
    
    ' 调用原有的添加方法，而不是递归调用 AddLiveroomSimple
    Call AddLiveroom(room)  ' 明确使用 Call 关键字
    
    AddLiveroomSimple = True
    Exit Function
    
ErrorHandler:
    Debug.Print "AddLiveroomSimple 错误: " & Err.Description & " (Web_rid: " & web_rid & ")"
    AddLiveroomSimple = False
End Function

' 确保原有的 AddLiveroom 方法正常工作
Public Sub AddLiveroom(ByRef room As LiveRoom)
    On Error GoTo ErrorHandler
    
    ' 检查是否已存在
    If LiveroomExists(room.web_rid) Then
        Debug.Print "LiveroomManager: 直播间已存在 - " & room.web_rid
        Exit Sub
    End If
    
    ' 添加到所有直播间集合
    m_allLiverooms.Add room, room.web_rid
    
    ' 根据启用状态添加到相应集合
    If room.enabled Then
        m_enabledLiverooms.Add room, room.web_rid
    Else
        m_disabledLiverooms.Add room, room.web_rid
    End If
    
    ' 初始化直播间属性
    room.LastCheckTime = 0
    If room.checkIntervalMinutes = 0 Then
        room.checkIntervalMinutes = 5 ' 默认5分钟检查间隔
    End If
    
    ' 启动定时器（如果还没有运行且有启用的直播间）
    If m_enabledLiverooms.count > 0 And Not m_timerRunning Then
        StartTimer
    End If
    
    Debug.Print "LiveroomManager: 添加直播间 - " & room.web_rid & " (主播: " & room.AnchorName & ", 优先级: " & GetPriorityName(room.taskPriority) & ", 启用: " & room.enabled & ")"
    RaiseEvent LiveroomAdded(room)
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "LiveroomManager.AddLiveroom 错误: " & Err.Description & " (Web_rid: " & room.web_rid & ")"
End Sub

' 添加辅助函数：获取优先级名称
Private Function GetPriorityName(ByVal priority As taskPriority) As String
    Select Case priority
        Case tpHigh: GetPriorityName = "高"
        Case tpMedium: GetPriorityName = "中"
        Case tpLow: GetPriorityName = "低"
        Case Else: GetPriorityName = "未知"
    End Select
End Function

' ========== 集合返回方法 ==========

' 获取所有直播间集合
Public Function GetAllLiverooms() As Collection
    Set GetAllLiverooms = m_allLiverooms
End Function

' 获取启用中的直播间集合
Public Function GetEnabledLiverooms() As Collection
    Set GetEnabledLiverooms = m_enabledLiverooms
End Function

' 获取禁用的直播间集合
Public Function GetDisabledLiverooms() As Collection
    Set GetDisabledLiverooms = m_disabledLiverooms
End Function

' 获取按状态筛选的直播间集合
Public Function GetLiveroomsByStatus(Optional ByVal enabledOnly As Boolean = True) As Collection
    Dim result As New Collection
    Dim room As LiveRoom
    
    If enabledOnly Then
        For Each room In m_enabledLiverooms
            result.Add room, room.web_rid
        Next room
    Else
        For Each room In m_allLiverooms
            result.Add room, room.web_rid
        Next room
    End If
    
    Set GetLiveroomsByStatus = result
End Function

' 获取按任务状态筛选的直播间集合
Public Function GetLiveroomsByTaskStatus(Optional ByVal taskType As String = "", _
                                         Optional ByVal isRecording As Boolean = False) As Collection
    Dim result As New Collection
    Dim room As LiveRoom
    
    For Each room In m_allLiverooms
        Select Case UCase(taskType)
            Case "PARSE"
                If room.ParseTaskID > 0 Then
                    result.Add room, room.web_rid
                End If
            Case "RECORD"
                If room.isRecording = isRecording Then
                    result.Add room, room.web_rid
                End If
            Case Else
                ' 返回所有直播间
                result.Add room, room.web_rid
        End Select
    Next room
    
    Set GetLiveroomsByTaskStatus = result
End Function

' 获取按房间状态筛选的直播间集合
Public Function GetLiveroomsByRoomStatus(Optional ByVal roomStatus As String = "") As Collection
    Dim result As New Collection
    Dim room As LiveRoom
    
    For Each room In m_allLiverooms
        If roomStatus = "" Or UCase(room.roomStatus) = UCase(roomStatus) Then
            result.Add room, room.web_rid
        End If
    Next room
    
    Set GetLiveroomsByRoomStatus = result
End Function

' 获取在监控时段的直播间集合
Public Function GetLiveroomsInWatchTime() As Collection
    Dim result As New Collection
    Dim room As LiveRoom
    
    For Each room In m_enabledLiverooms
        If room.IsInWatchTime() Then
            result.Add room, room.web_rid
        End If
    Next room
    
    Set GetLiveroomsInWatchTime = result
End Function

' 获取需要解析的直播间集合
Public Function GetLiveroomsNeedParse() As Collection
    Dim result As New Collection
    Dim room As LiveRoom
    
    For Each room In m_enabledLiverooms
        If room.NeedTriggerParseTask() Then
            result.Add room, room.web_rid
        End If
    Next room
    
    Set GetLiveroomsNeedParse = result
End Function

' 获取需要录制的直播间集合
Public Function GetLiveroomsNeedRecord() As Collection
    Dim result As New Collection
    Dim room As LiveRoom
    
    For Each room In m_enabledLiverooms
        If room.NeedTriggerRecordTask() Then
            result.Add room, room.web_rid
        End If
    Next room
    
    Set GetLiveroomsNeedRecord = result
End Function

' 获取活跃的直播间集合（正在解析或录制）
Public Function GetActiveLiverooms() As Collection
    Dim result As New Collection
    Dim room As LiveRoom
    
    For Each room In m_allLiverooms
        If room.ParseTaskID > 0 Or room.isRecording Then
            result.Add room, room.web_rid
        End If
    Next room
    
    Set GetActiveLiverooms = result
End Function

' ========== 统计信息方法 ==========

' 获取统计信息
Public Sub GetStats(ByRef TotalCount As Long, ByRef enabledCount As Long, ByRef disabledCount As Long)
    TotalCount = m_allLiverooms.count
    enabledCount = m_enabledLiverooms.count
    disabledCount = m_disabledLiverooms.count
End Sub

' 获取详细统计信息
Public Function GetDetailedStats() As String
    Dim TotalCount As Long, enabledCount As Long, disabledCount As Long
    GetStats TotalCount, enabledCount, disabledCount
    
    Dim parseCount As Long, recordCount As Long, activeCount As Long
    parseCount = GetLiveroomsByTaskStatus("PARSE").count
    recordCount = GetLiveroomsByTaskStatus("RECORD", True).count
    activeCount = GetActiveLiverooms().count
    
    Dim liveCount As Long, offlineCount As Long
    liveCount = GetLiveroomsByRoomStatus("直播中").count
    offlineCount = GetLiveroomsByRoomStatus("离线").count
    
    GetDetailedStats = "总计:" & TotalCount & _
                      " 启用:" & enabledCount & _
                      " 禁用:" & disabledCount & _
                      " 解析中:" & parseCount & _
                      " 录制中:" & recordCount & _
                      " 活跃:" & activeCount & _
                      " 直播中:" & liveCount & _
                      " 离线:" & offlineCount
End Function

' ========== 基本管理方法 ==========

' 设置调度器
Public Sub SetScheduler(ByRef scheduler As TaskScheduler)
    Set m_scheduler = scheduler
    Debug.Print "LiveroomManager: 调度器已设置"
End Sub

' 移除直播间
Public Sub RemoveLiveroom(ByVal web_rid As String)
    On Error GoTo ErrorHandler
    
    Dim room As LiveRoom
    Set room = GetLiveroom(web_rid)
    
    If room Is Nothing Then
        Debug.Print "LiveroomManager: 要移除的直播间不存在 - " & web_rid
        Exit Sub
    End If
    
    ' 从所有集合中移除
    RemoveFromCollection m_allLiverooms, web_rid
    RemoveFromCollection m_enabledLiverooms, web_rid
    RemoveFromCollection m_disabledLiverooms, web_rid
    
    ' 如果启用集合为空，停止定时器
    If m_enabledLiverooms.count = 0 And m_timerRunning Then
        StopTimer
    End If
    
    Debug.Print "LiveroomManager: 移除直播间 - " & web_rid
    RaiseEvent LiveroomRemoved(room)
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "LiveroomManager.RemoveLiveroom 错误: " & Err.Description & " (Web_rid: " & web_rid & ")"
End Sub

' 启用直播间
Public Sub EnableLiveroom(ByVal web_rid As String)
    On Error GoTo ErrorHandler
    
    Dim room As LiveRoom
    Set room = GetLiveroom(web_rid)
    
    If room Is Nothing Then
        Debug.Print "LiveroomManager: 要启用的直播间不存在 - " & web_rid
        Exit Sub
    End If
    
    ' 如果已经在启用集合中，直接返回
    If CollectionContains(m_enabledLiverooms, web_rid) Then
        Exit Sub
    End If
    
    ' 从禁用集合移到启用集合
    RemoveFromCollection m_disabledLiverooms, web_rid
    m_enabledLiverooms.Add room, web_rid
    
    ' 启动定时器（如果还没有运行）
    If m_enabledLiverooms.count > 0 And Not m_timerRunning Then
        StartTimer
    End If
    
    Debug.Print "LiveroomManager: 启用直播间 - " & web_rid
    RaiseEvent LiveroomEnabled(room)
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "LiveroomManager.EnableLiveroom 错误: " & Err.Description & " (Web_rid: " & web_rid & ")"
End Sub

' 禁用直播间
Public Sub DisableLiveroom(ByVal web_rid As String)
    On Error GoTo ErrorHandler
    
    Dim room As LiveRoom
    Set room = GetLiveroom(web_rid)
    
    If room Is Nothing Then
        Debug.Print "LiveroomManager: 要禁用的直播间不存在 - " & web_rid
        Exit Sub
    End If
    
    ' 如果已经在禁用集合中，直接返回
    If CollectionContains(m_disabledLiverooms, web_rid) Then
        Exit Sub
    End If
    
    ' 从启用集合移到禁用集合
    RemoveFromCollection m_enabledLiverooms, web_rid
    m_disabledLiverooms.Add room, web_rid
    
    ' 如果启用集合为空，停止定时器
    If m_enabledLiverooms.count = 0 And m_timerRunning Then
        StopTimer
    End If
    
    Debug.Print "LiveroomManager: 禁用直播间 - " & web_rid
    RaiseEvent LiveroomDisabled(room)
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "LiveroomManager.DisableLiveroom 错误: " & Err.Description & " (Web_rid: " & web_rid & ")"
End Sub

' 获取直播间
Public Function GetLiveroom(ByVal web_rid As String) As LiveRoom
    On Error Resume Next
    Set GetLiveroom = m_allLiverooms(web_rid)
End Function

' 检查直播间是否存在
Public Function LiveroomExists(ByVal web_rid As String) As Boolean
    On Error Resume Next
    Dim room As LiveRoom
    Set room = m_allLiverooms(web_rid)
    LiveroomExists = (Not room Is Nothing)
End Function

' ========== Timer 管理 ==========
Private Sub StartTimer()
    If m_timerID = 0 Then
        m_timerID = SetTimer(0, 0, TIMER_INTERVAL, AddressOf LiveroomManagerTimerProc)
        If m_timerID <> 0 Then
            m_timerRunning = True
            Debug.Print "LiveroomManager: 定时器启动成功 (ID: " & m_timerID & ")"
        Else
            Debug.Print "LiveroomManager: 错误 - 无法启动定时器"
        End If
    End If
End Sub

Public Sub StopTimer()
    On Error Resume Next
    If m_timerID <> 0 Then
        KillTimer 0, m_timerID
        m_timerID = 0
        m_timerRunning = False
        Debug.Print "LiveroomManager: 定时器已停止"
    End If
End Sub

' ?? 修复：简化的轮询方法，确保每次只更新一次表格
Public Sub ProcessPolling()
    On Error Resume Next
    
    If m_timerRunning And g_appShuttingDown Then
        StopTimer
        Exit Sub
    End If
    
    If Not m_timerRunning Then
        Exit Sub
    End If
    
    Static lastProcessTime As Date
    Dim currentTime As Date
    currentTime = Now
    
    ' 基本的轮询频率控制
    If (currentTime - lastProcessTime) * 24 * 60 * 60 < 0.25 Then
        Exit Sub
    End If
    
    lastProcessTime = currentTime
    
    ' ?? 处理任务逻辑
    Dim room As LiveRoom
    For Each room In m_enabledLiverooms
        CheckAndTriggerTasks room
    Next room
    
    ' ?? 修复：每次轮询只更新一次表格
    If Not frmMain Is Nothing Then
        frmMain.SingleUpdateFromManager Me.GetAllLiverooms()
    End If
    
    Debug.Print ">>> 轮询完成并更新显示: " & Format(Now, "hh:mm:ss") & " - 直播间数: " & m_allLiverooms.count
End Sub

' ========== 任务触发逻辑 ==========

' 检查并触发任务
Private Sub CheckAndTriggerTasks(ByRef room As LiveRoom)
    On Error Resume Next
    
    ' 检查解析任务触发条件
    If room.NeedTriggerParseTask() Then
        TriggerParseTask room
    End If
    
    ' 检查录制任务触发条件
    If room.NeedTriggerRecordTask() Then
        TriggerRecordTask room
    End If
End Sub

' 触发解析任务
Private Sub TriggerParseTask(ByRef room As LiveRoom)
    On Error Resume Next
    
    If m_scheduler Is Nothing Then
        Debug.Print "LiveroomManager: 错误 - 调度器未设置，无法触发解析任务"
        Exit Sub
    End If
    
    ' 提交解析任务
    room.SubmitParseTask m_scheduler
    
    Debug.Print "LiveroomManager: 触发解析任务 - " & room.web_rid & " (任务ID: " & room.ParseTaskID & ")"
    RaiseEvent ParseTaskTriggered(room)
End Sub

' 触发录制任务
Private Sub TriggerRecordTask(ByRef room As LiveRoom)
    On Error Resume Next
    
    If m_scheduler Is Nothing Then
        Debug.Print "LiveroomManager: 错误 - 调度器未设置，无法触发录制任务"
        Exit Sub
    End If
    
    ' 提交录制任务
    room.SubmitRecordTask m_scheduler
    
    Debug.Print "LiveroomManager: 触发录制任务 - " & room.web_rid & " (任务ID: " & room.RecordTaskID & ")"
    RaiseEvent RecordTaskTriggered(room)
End Sub

' ========== 任务控制方法 ==========

' 停止指定直播间的录制任务
Public Function StopRecordTask(ByVal web_rid As String) As Boolean
    On Error GoTo ErrorHandler
    
    If m_scheduler Is Nothing Then
        Debug.Print "LiveroomManager: 错误 - 调度器未设置"
        StopRecordTask = False
        Exit Function
    End If
    
    ' 通过调度器停止录制任务
    StopRecordTask = m_scheduler.StopRecordTask(web_rid)
    
    If StopRecordTask Then
        Debug.Print "LiveroomManager: 已停止录制任务 - " & web_rid
    Else
        Debug.Print "LiveroomManager: 停止录制任务失败 - " & web_rid
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "LiveroomManager.StopRecordTask 错误: " & Err.Description
    StopRecordTask = False
End Function

' 停止所有录制任务
Public Function StopAllRecordTasks() As Boolean
    On Error GoTo ErrorHandler
    
    If m_scheduler Is Nothing Then
        Debug.Print "LiveroomManager: 错误 - 调度器未设置"
        StopAllRecordTasks = False
        Exit Function
    End If
    
    StopAllRecordTasks = m_scheduler.StopAllRecordTasks()
    
    Exit Function
    
ErrorHandler:
    Debug.Print "LiveroomManager.StopAllRecordTasks 错误: " & Err.Description
    StopAllRecordTasks = False
End Function

' 获取录制任务状态
Public Function GetRecordTaskStatus(ByVal web_rid As String) As String
    On Error Resume Next
    
    Dim room As LiveRoom
    Set room = GetLiveroom(web_rid)
    
    If room Is Nothing Then
        GetRecordTaskStatus = "直播间不存在"
        Exit Function
    End If
    
    If room.isRecording Then
        GetRecordTaskStatus = "录制中"
    ElseIf room.RecordTaskStatus = "排队中" Then
        GetRecordTaskStatus = "排队中"
    ElseIf room.RecordTaskStatus = "运行中" Then
        GetRecordTaskStatus = "运行中"
    Else
        GetRecordTaskStatus = "未录制"
    End If
End Function

' 获取解析任务状态
Public Function GetParseTaskStatus(ByVal web_rid As String) As String
    On Error Resume Next
    
    Dim room As LiveRoom
    Set room = GetLiveroom(web_rid)
    
    If room Is Nothing Then
        GetParseTaskStatus = "直播间不存在"
        Exit Function
    End If
    
    GetParseTaskStatus = room.ParseTaskStatus
End Function

' ========== 批量操作方法 ==========

' 启用所有直播间
Public Sub EnableAllLiverooms()
    On Error Resume Next
    
    Dim room As LiveRoom
    For Each room In m_disabledLiverooms
        EnableLiveroom room.web_rid
    Next room
    
    Debug.Print "LiveroomManager: 已启用所有直播间"
End Sub

' 禁用所有直播间
Public Sub DisableAllLiverooms()
    On Error Resume Next
    
    Dim room As LiveRoom
    For Each room In m_enabledLiverooms
        DisableLiveroom room.web_rid
    Next room
    
    Debug.Print "LiveroomManager: 已禁用所有直播间"
End Sub

' 清除所有直播间
Public Sub ClearAllLiverooms()
    On Error Resume Next
    
    StopTimer
    
    Set m_allLiverooms = New Collection
    Set m_enabledLiverooms = New Collection
    Set m_disabledLiverooms = New Collection
    
    Debug.Print "LiveroomManager: 已清除所有直播间"
End Sub

' ========== 辅助函数 ==========
Private Sub RemoveFromCollection(ByRef col As Collection, ByVal key As String)
    On Error Resume Next
    col.Remove key
End Sub

Private Function CollectionContains(ByRef col As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Dim item As Object
    Set item = col(key)
    CollectionContains = (Not item Is Nothing)
End Function

' 获取直播间数量统计
Public Sub GetCountStats(ByRef total As Long, ByRef enabled As Long, ByRef disabled As Long, _
                        ByRef parsing As Long, ByRef recording As Long)
    On Error Resume Next
    
    total = m_allLiverooms.count
    enabled = m_enabledLiverooms.count
    disabled = m_disabledLiverooms.count
    
    parsing = 0
    recording = 0
    
    Dim room As LiveRoom
    For Each room In m_allLiverooms
        If room.ParseTaskID > 0 Then
            parsing = parsing + 1
        End If
        If room.isRecording Then
            recording = recording + 1
        End If
    Next room
End Sub

' 导出直播间列表到字符串
Public Function ExportLiveroomsToString() As String
    On Error Resume Next
    
    Dim result As String
    Dim room As LiveRoom
    
    For Each room In m_allLiverooms
        If result <> "" Then
            result = result & ","
        End If
        result = result & room.web_rid
    Next room
    
    ExportLiveroomsToString = result
End Function

' 从字符串导入直播间列表
Public Sub ImportLiveroomsFromString(ByVal liveroomsString As String, _
                                    Optional ByVal checkIntervalMinutes As Long = 5)
    On Error Resume Next
    
    If liveroomsString = "" Then Exit Sub
    
    Dim web_rids() As String
    web_rids = Split(liveroomsString, ",")
    
    Dim i As Long
    For i = LBound(web_rids) To UBound(web_rids)
        Dim web_rid As String
        web_rid = Trim(web_rids(i))
        If web_rid <> "" And Not LiveroomExists(web_rid) Then
            AddLiveroomSimple web_rid, "", checkIntervalMinutes
        End If
    Next i
    
    Debug.Print "LiveroomManager: 从字符串导入了 " & (UBound(web_rids) - LBound(web_rids) + 1) & " 个直播间"
End Sub

