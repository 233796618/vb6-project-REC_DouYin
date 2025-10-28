VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LiveRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 新增：Tag 属性用于标记备注信息
Public Tag As String

' 监控时段相关属性
Public checkAllTime As Boolean          ' 24小时不停监控
Public watchTimeStart As String         ' 监控开始时刻 (格式: "HH:MM" 或 "H")
Public watchTimeEnd As String           ' 监控结束时刻 (格式: "HH:MM" 或 "H")

Public parseTime As String

' 录制任务相关属性
Public RecordTaskID As Long             ' 录制任务ID
Public isRecording As Boolean           ' 是否正在录制
Public LastRecordTime As Date           ' 上次录制时间

Public ParseTaskID As Long

' 任务状态跟踪
Public ParseTaskStatus As String        ' 解析任务状态: "未提交", "排队中", "运行中", "已完成"
Public RecordTaskStatus As String       ' 录制任务状态: "未提交", "排队中", "运行中", "停止中", "已完成"
Public ParseTaskSubmitTime As Date      ' 解析任务提交时间
Public RecordTaskSubmitTime As Date     ' 录制任务提交时间
Public ParseTaskStartTime As Date       ' 解析任务开始时间
Public RecordTaskStartTime As Date      ' 录制任务开始时间

' 新增：启用/禁用状态
Public enabled As Boolean               ' 直播间是否启用

' 新增：任务优先级
Public taskPriority As taskPriority     ' 任务优先级

'不落地的一组变量
Public NotRecUntilTime As String '计算到某时刻之前，都不需要录制。但是如果CheckAllTime，仍然检查

Public StepStatus As String '任务步骤和状态，准备，录制结束，准备
'第一组是解析检查相关的基本属性
Public LastCheckTime As Date
Public AnchorName As String
Public web_rid As String
Public room_title As String
Public roomStatus As String
Public StreamGrade As String
Public streamUrl As String
Public streamWidth As Long
Public streamHeight As Long

'第二组是程序设置相关的
Public checkIntervalMinutes As Long
Public SelQuality As String
Public TaskTag As String
Public TaskStatus As String
Public RepeatTimesLimitEachDay As String
Public ValidEndureMinutes As String
Public RangeRec As Boolean
Public StartRecTime As String
Public EndRecTime As String
'Public CheckAllTime As Boolean
Public ForceStopRecordInIddleTime As Boolean
Public CropVideo As Boolean
Public CropLeft As Long
Public CropTop As Long
Public CropWidth As Long
Public CropHeight As Long

'第三组是录制相关的
Public RecStatus As String
Public CurRecordStartTime As Date
Public RecordMinutes As Long
Public FileSize As Long
Public IsCropped As Boolean ' 应该是等同于 CropVideo，但是实际并不等同。因为 如果分辨率不对，就不能实际crop
Public IsForcedStopped As Boolean

'这一组变量用来辅助解析直播流
Private m_Streams() As StreamInfo
Private m_StreamCount As Long
Private Type StreamInfo
    Url As String
    QualityDescription As String
    QualityScore As Long
End Type

' FFmpeg信息变量
Public currentFrame As String
Public currentFPS As String
Public currentQuality As String
Public currentFileSize As String
Public currentTime As String
Public currentBitrate As String
Public currentSpeed As String
Public currentElapsed As String
Public FFmepgExitCode As String

' 数据变化检测相关
Private m_lastUpdateHash As String
Private m_HasChanges As Boolean
Private m_forceChanges As Boolean ' 新增：强制变化标志

' 用于优化的静态变量
Private m_lastRoomStatus As String
Private m_lastParseStatus As String
Private m_lastRecordStatus As String
Private m_lastBitrate As String
Private m_lastEnabled As Boolean

' ?? 新增：重要的变化检测静态变量
Private m_significantRoomStatus As String
Private m_significantParseStatus As String
Private m_significantRecordStatus As String
Private m_significantBitrate As String
Private m_significantEnabled As Boolean

Private Sub Class_Initialize()
    ' 初始化 Tag 属性
    Tag = ""
    
    LastCheckTime = 0 '初始化变量
    ParseTaskStatus = "未提交"
    RecordTaskStatus = "未提交"
    ParseTaskID = 0
    RecordTaskID = 0
    isRecording = False
    enabled = True  ' 默认启用
    taskPriority = tpMedium  ' 默认中等优先级
    
    ' 初始化变化检测
    m_lastUpdateHash = ""
    m_HasChanges = True  ' 初始化为有变化，确保首次显示
    m_forceChanges = True  ' 初始时强制更新
    
    ' 初始化静态变量
    m_lastRoomStatus = ""
    m_lastParseStatus = ""
    m_lastRecordStatus = ""
    m_lastBitrate = ""
    m_lastEnabled = enabled
    
    ' 初始化重要变化检测变量
    m_significantRoomStatus = ""
    m_significantParseStatus = ""
    m_significantRecordStatus = ""
    m_significantBitrate = ""
    m_significantEnabled = enabled
End Sub

' 检查是否需要触发解析任务
Public Function NeedTriggerParseTask() As Boolean
    ' 如果被禁用，则不触发任何任务
    If Not enabled Then
        NeedTriggerParseTask = False
        Exit Function
    End If
    
    ' 如果正在录制，则不需要触发解析任务
    If isRecording Then
        NeedTriggerParseTask = False
        Exit Function
    End If
    
    ' 如果解析任务已经在排队或运行中，则不需要触发
    If ParseTaskStatus = "排队中" Or ParseTaskStatus = "运行中" Then
        NeedTriggerParseTask = False
        Exit Function
    End If
    
    ' 检查时间间隔
    If LastCheckTime = 0 Then
        ' 从未检查过，需要立即检查
        NeedTriggerParseTask = True
    Else
        Dim minutesSinceLastCheck As Double
        minutesSinceLastCheck = (Now - LastCheckTime) * 24 * 60
        
        If minutesSinceLastCheck >= checkIntervalMinutes Then
            NeedTriggerParseTask = True
        Else
            NeedTriggerParseTask = False
        End If
    End If
End Function

' 检查是否需要触发录制任务
Public Function NeedTriggerRecordTask() As Boolean
    ' 如果被禁用，则不触发任何任务
    If Not enabled Then
        NeedTriggerRecordTask = False
        Exit Function
    End If
    
    ' 如果已经在录制，则不需要触发
    If isRecording Then
        NeedTriggerRecordTask = False
        Exit Function
    End If
    
    ' 如果录制任务已经在排队或运行中，则不需要触发
    If RecordTaskStatus = "排队中" Or RecordTaskStatus = "运行中" Then
        NeedTriggerRecordTask = False
        Exit Function
    End If
    
    ' 检查直播间状态和监控时段
    If roomStatus = "直播中" And IsInWatchTime() Then
        NeedTriggerRecordTask = True
    Else
        NeedTriggerRecordTask = False
    End If
End Function

' 检查是否在监控时段内
Public Function IsInWatchTime() As Boolean
    On Error Resume Next
    
    ' 如果24小时监控，始终返回True
    If checkAllTime Then
        IsInWatchTime = True
        Exit Function
    End If
    
    ' 如果没有设置监控时段，返回False
    If watchTimeStart = "" Or watchTimeEnd = "" Then
        IsInWatchTime = False
        Exit Function
    End If
    
    Dim currentTime As Date
    currentTime = Now
    
    Dim currentHour As Integer
    Dim currentMinute As Integer
    currentHour = Hour(currentTime)
    currentMinute = Minute(currentTime)
    
    Dim startHour As Integer
    Dim startMinute As Integer
    Dim endHour As Integer
    Dim endMinute As Integer
    
    ' 解析开始时间
    If InStr(watchTimeStart, ":") > 0 Then
        startHour = Hour(CDate(watchTimeStart))
        startMinute = Minute(CDate(watchTimeStart))
    Else
        startHour = CInt(watchTimeStart)
        startMinute = 0
    End If
    
    ' 解析结束时间
    If InStr(watchTimeEnd, ":") > 0 Then
        endHour = Hour(CDate(watchTimeEnd))
        endMinute = Minute(CDate(watchTimeEnd))
    Else
        endHour = CInt(watchTimeEnd)
        endMinute = 0
    End If
    
    ' 转换为分钟数便于比较
    Dim currentTotalMinutes As Integer
    Dim startTotalMinutes As Integer
    Dim endTotalMinutes As Integer
    
    currentTotalMinutes = currentHour * 60 + currentMinute
    startTotalMinutes = startHour * 60 + startMinute
    endTotalMinutes = endHour * 60 + endMinute
    
    ' 判断当前时间是否在监控时段内
    If startTotalMinutes <= endTotalMinutes Then
        ' 正常时段：0点不在监控时段内
        IsInWatchTime = (currentTotalMinutes >= startTotalMinutes And currentTotalMinutes < endTotalMinutes)
    Else
        ' 跨天时段：0点在监控时段内
        IsInWatchTime = (currentTotalMinutes >= startTotalMinutes Or currentTotalMinutes < endTotalMinutes)
    End If
End Function

' 标记解析任务开始运行
Public Sub MarkParseTaskRunning()
    ParseTaskStatus = "运行中"
    ParseTaskStartTime = Now
    SmartMarkAsChanged
    Debug.Print "直播间 " & web_rid & ": 解析任务开始运行"
End Sub

' 标记解析任务完成
Public Sub MarkParseTaskCompleted()
    ParseTaskStatus = "已完成"
    ParseTaskID = 0
    SmartMarkAsChanged
    Debug.Print "直播间 " & web_rid & ": 解析任务完成"
End Sub

Public Sub SubmitParseTask(TS As TaskScheduler)
    ' 明确设置任务名称
    Dim taskName As String
    taskName = "解析任务_" & web_rid
    
    ParseTaskID = TS.AddTask("parser", App.Path & "\parser.exe", " -h " & frm_WMCopyData.hWnd & " -id " & web_rid, taskName, taskPriority, , , web_rid, App.Path)
    
    ' 更新状态
    ParseTaskStatus = "排队中"
    ParseTaskSubmitTime = Now
    LastCheckTime = Now
    SmartMarkAsChanged
    
    ' 直接调用全局函数获取优先级名称
    Dim priorityName As String
    priorityName = GetPriorityText(taskPriority)
    
    Debug.Print "直播间 " & web_rid & ": 提交解析任务 (任务ID: " & ParseTaskID & ", 任务名称: " & taskName & ", 优先级: " & priorityName & ")"
End Sub

Public Sub SubmitRecordTask(TS As TaskScheduler)
    ' 明确设置任务名称
    Dim taskName As String
    taskName = "录制任务_" & web_rid
    
    Dim cmdLine As String
    cmdLine = PrepareCmdLine(streamUrl, AnchorName, web_rid)
    Debug.Print "生成ffmpeg命令:" & cmdLine
    
    RecordTaskID = TS.AddTask("recorder", cmdLine, , taskName, taskPriority, , True, web_rid, App.Path)
    
    ' 更新状态
    RecordTaskStatus = "排队中"
    RecordTaskSubmitTime = Now
    isRecording = True
    SmartMarkAsChanged
    
    ' 直接调用全局函数获取优先级名称
    Dim priorityName As String
    priorityName = GetPriorityText(taskPriority)
    
    Debug.Print "直播间 " & web_rid & ": 提交录制任务 (任务ID: " & RecordTaskID & ", 任务名称: " & taskName & ", 优先级: " & priorityName & ")"
End Sub

' 标记录制任务开始运行
Public Sub MarkRecordTaskRunning()
    RecordTaskStatus = "运行中"
    RecordTaskStartTime = Now
    SmartMarkAsChanged
    Debug.Print "直播间 " & web_rid & ": 录制任务开始运行"
End Sub

' 新增：标记录制任务为停止中状态
Public Sub MarkRecordTaskStopping()
    RecordTaskStatus = "停止中"
    SmartMarkAsChanged
    Debug.Print "直播间 " & web_rid & ": 录制任务正在停止"
End Sub

' 标记录制任务完成
Public Sub MarkRecordTaskCompleted()
    RecordTaskStatus = "已完成"
    RecordTaskID = 0
    isRecording = False
    
    ' 重置 FFmpeg 相关信息
    ResetFFmpegInfo
    
    SmartMarkAsChanged
    Debug.Print "直播间 " & web_rid & ": 录制任务完成"
End Sub

' 重置解析任务状态（用于清理）
Public Sub ResetParseTask()
    ParseTaskID = 0
    ParseTaskStatus = "未提交"
    ParseTaskSubmitTime = 0
    ParseTaskStartTime = 0
    SmartMarkAsChanged
End Sub

' 重置录制任务状态（用于清理）
Public Sub ResetRecordTask()
    RecordTaskID = 0
    RecordTaskStatus = "未提交"
    RecordTaskSubmitTime = 0
    RecordTaskStartTime = 0
    isRecording = False
    
    ' 重置 FFmpeg 信息
    ResetFFmpegInfo
    
    SmartMarkAsChanged
End Sub

' 启用直播间
Public Sub EnableRoom()
    If Not enabled Then
        enabled = True
        SmartMarkAsChanged
        Debug.Print "直播间 " & web_rid & ": 已启用"
    Else
        Debug.Print "直播间 " & web_rid & ": 已经处于启用状态"
    End If
End Sub

' 禁用直播间
Public Sub DisableRoom()
    If enabled Then
        enabled = False
        SmartMarkAsChanged
        Debug.Print "直播间 " & web_rid & ": 已禁用"
        
        ' 如果正在录制，停止录制任务
        If isRecording Then
            StopRecordTask web_rid
        End If
    End If
End Sub

' ========== 优化的数据变化检测方法 ==========

' ?? 新增：优化的变化检测 - 只检测真正重要的变化
Public Function HasSignificantChanges() As Boolean
    On Error Resume Next
    
    ' 检查关键字段的变化
    If m_significantRoomStatus <> roomStatus Or _
       m_significantParseStatus <> ParseTaskStatus Or _
       m_significantRecordStatus <> IIf(isRecording, "录制中", "未录制") Or _
       m_significantBitrate <> currentBitrate Or _
       m_significantEnabled <> enabled Then
        
        ' 更新静态变量
        m_significantRoomStatus = roomStatus
        m_significantParseStatus = ParseTaskStatus
        m_significantRecordStatus = IIf(isRecording, "录制中", "未录制")
        m_significantBitrate = currentBitrate
        m_significantEnabled = enabled
        
        HasSignificantChanges = True
        
        ' 调试输出重要变化
        Debug.Print ">>> 重要变化检测: " & web_rid & _
                    " 状态:" & roomStatus & _
                    " 解析:" & ParseTaskStatus & _
                    " 录制:" & IIf(isRecording, "是", "否") & _
                    " 码率:" & currentBitrate & _
                    " 启用:" & IIf(enabled, "是", "否")
    Else
        HasSignificantChanges = False
    End If
End Function

' ?? 新增：智能标记变化
Public Sub SmartMarkAsChanged()
    On Error Resume Next
    If HasSignificantChanges() Then
        MarkAsChanged
        Debug.Print ">>> 智能标记变化: " & web_rid & " - 检测到重要变化"
    End If
End Sub

' 优化的数据变化检测
Public Function HasChanges() As Boolean
    On Error Resume Next
    
    ' 如果强制更新标志为真，总是返回有变化
    If m_forceChanges Then
        m_forceChanges = False
        HasChanges = True
        Exit Function
    End If
    
    ' 检查关键字段的变化
    If m_lastRoomStatus <> roomStatus Or _
       m_lastParseStatus <> ParseTaskStatus Or _
       m_lastRecordStatus <> IIf(isRecording, "录制中", "未录制") Or _
       m_lastBitrate <> currentBitrate Or _
       m_lastEnabled <> enabled Then
        
        m_lastRoomStatus = roomStatus
        m_lastParseStatus = ParseTaskStatus
        m_lastRecordStatus = IIf(isRecording, "录制中", "未录制")
        m_lastBitrate = currentBitrate
        m_lastEnabled = enabled
        
        HasChanges = True
    Else
        HasChanges = False
    End If
End Function

' 强制标记变化
Public Sub ForceMarkAsChanged()
    m_forceChanges = True
    m_HasChanges = True
    Debug.Print ">>> 强制标记变化: " & web_rid
End Sub

' 获取数据哈希值，用于检测变化
Public Function GetDataHash() As String
    Dim hashString As String
    hashString = roomStatus & "|" & ParseTaskStatus & "|" & IIf(isRecording, "1", "0") & "|" & _
                currentBitrate & "|" & currentFPS & "|" & currentFileSize & "|" & currentTime & "|" & _
                IIf(enabled, "1", "0") & "|" & GetPriorityDisplayName(taskPriority) & "|" & _
                IIf(LastCheckTime <> 0, Format(LastCheckTime, "hh:mm:ss"), "")
    GetDataHash = CStr(Len(hashString)) & ":" & hashString
End Function

' 重置变化标志
Public Sub ResetChangeFlag()
    m_HasChanges = False
End Sub

' 标记数据已变化
Public Sub MarkAsChanged()
    m_HasChanges = True
    m_forceChanges = True ' 同时设置强制变化标志
End Sub

' 获取优先级显示名称
Private Function GetPriorityDisplayName(ByVal priority As taskPriority) As String
    Select Case priority
        Case tpHigh: GetPriorityDisplayName = "高"
        Case tpMedium: GetPriorityDisplayName = "中"
        Case tpLow: GetPriorityDisplayName = "低"
        Case Else: GetPriorityDisplayName = "未知"
    End Select
End Function

' 新增：重置 FFmpeg 相关信息
Public Sub ResetFFmpegInfo()
    currentFrame = ""
    currentFPS = ""
    currentQuality = ""
    currentFileSize = ""
    currentTime = ""
    currentBitrate = ""
    currentSpeed = ""
    currentElapsed = ""
    FFmepgExitCode = ""
    
    Debug.Print "直播间 " & web_rid & ": 已重置 FFmpeg 信息"
End Sub

' 新增：获取任务状态显示文本
Public Function GetTaskStatusDisplay() As String
    Dim statusText As String
    
    ' 录制状态优先显示
    If isRecording Then
        If RecordTaskStatus = "停止中" Then
            statusText = "录制停止中"
        Else
            statusText = "录制中"
        End If
    ElseIf RecordTaskStatus = "排队中" Then
        statusText = "录制排队中"
    ElseIf RecordTaskStatus = "运行中" Then
        statusText = "录制运行中"
    ElseIf ParseTaskStatus = "运行中" Then
        statusText = "解析中"
    ElseIf ParseTaskStatus = "排队中" Then
        statusText = "解析排队中"
    ElseIf roomStatus = "直播中" Then
        statusText = "直播中"
    Else
        statusText = roomStatus
    End If
    
    GetTaskStatusDisplay = statusText
End Function

' 新增：获取监控时段显示文本
Public Function GetWatchTimeDisplay() As String
    If checkAllTime Then
        GetWatchTimeDisplay = "全天"
    ElseIf watchTimeStart <> "" And watchTimeEnd <> "" Then
        GetWatchTimeDisplay = watchTimeStart & "-" & watchTimeEnd
    Else
        GetWatchTimeDisplay = "未设置"
    End If
End Function

' 新增：获取优先级显示文本
Public Function GetPriorityDisplay() As String
    Select Case taskPriority
        Case tpHigh: GetPriorityDisplay = "高"
        Case tpMedium: GetPriorityDisplay = "中"
        Case tpLow: GetPriorityDisplay = "低"
        Case Else: GetPriorityDisplay = "未知"
    End Select
End Function

' 新增：检查是否处于活跃状态（解析中或录制中）
Public Function IsActive() As Boolean
    IsActive = (ParseTaskStatus = "运行中" Or ParseTaskStatus = "排队中" Or _
                isRecording Or RecordTaskStatus = "排队中" Or RecordTaskStatus = "运行中")
End Function

' 新增：检查是否可以接受新任务
Public Function CanAcceptNewTask() As Boolean
    ' 如果被禁用，不能接受新任务
    If Not enabled Then
        CanAcceptNewTask = False
        Exit Function
    End If
    
    ' 如果正在录制，不能接受解析任务
    If isRecording Then
        CanAcceptNewTask = False
        Exit Function
    End If
    
    ' 如果解析任务正在运行或排队中，不能接受新的解析任务
    If ParseTaskStatus = "运行中" Or ParseTaskStatus = "排队中" Then
        CanAcceptNewTask = False
        Exit Function
    End If
    
    ' 如果录制任务正在运行或排队中，不能接受新的录制任务
    If RecordTaskStatus = "运行中" Or RecordTaskStatus = "排队中" Then
        CanAcceptNewTask = False
        Exit Function
    End If
    
    CanAcceptNewTask = True
End Function

' 新增：获取任务状态详细信息
Public Function GetTaskStatusDetails() As String
    Dim details As String
    
    details = "直播间状态: " & roomStatus & vbCrLf & _
              "解析任务: " & ParseTaskStatus & vbCrLf & _
              "录制任务: " & RecordTaskStatus & vbCrLf & _
              "启用状态: " & IIf(enabled, "已启用", "已禁用") & vbCrLf & _
              "优先级: " & GetPriorityDisplay() & vbCrLf & _
              "Tag: " & Tag
    
    GetTaskStatusDetails = details
End Function

