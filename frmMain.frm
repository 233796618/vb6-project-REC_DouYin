VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMain 
   Caption         =   "抖音直播录制工具"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   25740
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   25740
   StartUpPosition =   3  '窗口缺省
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid 
      Height          =   8655
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   25695
      _ExtentX        =   45323
      _ExtentY        =   15266
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      GridLines       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdStopRecordTask 
      Caption         =   "停止选中录制"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdStopAllRecords 
      Caption         =   "停止所有录制"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "添加直播间"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Menu mnuContextMenu 
      Caption         =   "ContextMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuEnableRoom 
         Caption         =   "启用(&E)"
      End
      Begin VB.Menu mnuDisableRoom 
         Caption         =   "禁用(&D)"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRoom 
         Caption         =   "编辑(&E)"
      End
      Begin VB.Menu mnuRemoveRoom 
         Caption         =   "删除(&R)"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIncreasePriority 
         Caption         =   "提高优先级(&H)"
      End
      Begin VB.Menu mnuDecreasePriority 
         Caption         =   "降低优先级(&L)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 在 frmMain.frm 的声明部分添加这些常量
Private Const flexHighlightAlways As Long = 2
Private Const flexSelectionByRow As Long = 1
Private Const flexFocusNone As Long = 0

Private Const GWL_STYLE As Long = (-16)
Private Const WS_CLIPCHILDREN As Long = &H2000000

' 事件声明
Private WithEvents m_scheduler As TaskScheduler
Attribute m_scheduler.VB_VarHelpID = -1
'Attribute m_scheduler.VB_VarHelpID = -1
Private WithEvents M_LiveroomManager As LiveroomManager
Attribute M_LiveroomManager.VB_VarHelpID = -1
'Attribute M_LiveroomManager.VB_VarHelpID = -1

' 助手对象
Private m_gridHelper As CMSHFlexGridHelper
Private m_dataBuffer As CGridDataBuffer

' 状态变量
Private m_lastRoomCount As Long
Private m_sortColumn As Long
Private m_sortAscending As Boolean
Private m_lastSortColumn As Long

' 新增：排序状态管理
Private m_isSorted As Boolean
Private m_currentSortedRooms As Collection

' ========== 窗体生命周期事件 ==========

Private Sub Form_Load()
    ' 初始化助手对象
    Set m_gridHelper = New CMSHFlexGridHelper
    Set m_dataBuffer = New CGridDataBuffer
    m_lastRoomCount = 0
    
    ' 初始化排序变量
    m_sortColumn = -1
    m_sortAscending = True
    m_lastSortColumn = -1
    
    ' 初始化排序状态
    m_isSorted = False
    Set m_currentSortedRooms = New Collection
    
    ' 初始化全局变量
    Debug.Print ">>> 开始初始化全局变量..."
    InitializeGlobalVariables
    
    ' 检查初始化结果
    If g_liveroomManager Is Nothing Then
        Debug.Print ">>> 错误: g_liveroomManager 初始化失败"
    Else
        Debug.Print ">>> g_liveroomManager 初始化成功"
    End If
    
    If g_scheduler Is Nothing Then
        Debug.Print ">>> 错误: g_scheduler 初始化失败"
    Else
        Debug.Print ">>> g_scheduler 初始化成功"
    End If
    
    ' 设置事件处理
    Set m_scheduler = g_scheduler
    Set M_LiveroomManager = g_liveroomManager
    
    ' 初始化界面组件
    InitializeMSHFlexGrid
    InitializeStatsLabel
    
    ' 设置网格选择模式
    m_gridHelper.ForceFullRowSelection MSHFlexGrid
    SetSingleRowSelection MSHFlexGrid
    
    ' 加载消息窗体
    LogMessage "调度器已初始化"
    frm_WMCopyData.Visible = False
    Load frm_WMCopyData
    
    ' 初始数据更新测试
    If Not M_LiveroomManager Is Nothing Then
        Dim testRooms As Collection
        Set testRooms = M_LiveroomManager.GetAllLiverooms()
        Debug.Print ">>> 立即测试更新: " & testRooms.count & " 个直播间"
        DirectUpdateFromManager testRooms
    End If
    
    Debug.Print ">>> 窗体加载完成"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frm_WMCopyData
    SafeFormUnload Me, Cancel
    End
End Sub

' ========== 界面初始化方法 ==========

Private Sub InitializeMSHFlexGrid()
    On Error Resume Next
    m_gridHelper.InitializeGrid MSHFlexGrid
    
    With MSHFlexGrid
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .ScrollBars = flexScrollBarBoth
        .FocusRect = flexFocusNone
        
        If .rows > 1 Then
            .row = 1
            .col = 0
            .ColSel = .cols - 1
        End If
    End With
    
    Debug.Print "MSHFlexGrid 初始化完成"
End Sub

Private Sub InitializeStatsLabel()
    ' 修复：使用新的方法名避免递归调用
    InitializeStatsLabelOnForm Me
End Sub

Private Sub Form_Resize()
    ' 修复：使用正确的参数
    HandleFormResize Me, MSHFlexGrid
End Sub

' ========== 数据更新方法 ==========

Public Sub DirectUpdateFromManager(ByRef rooms As Collection)
    On Error Resume Next
    
    If m_isSorted Then
        ' 如果当前处于排序状态，保持排序
        MaintainSortedDisplay rooms
    Else
        ' 正常更新
        UpdateGridWithBuffer MSHFlexGrid, rooms
    End If
    
    UpdateStatsLabel
End Sub

Public Sub SingleUpdateFromManager(ByRef rooms As Collection)
    On Error Resume Next
    DirectUpdateFromManager rooms
    Debug.Print ">>> 单次整体更新完成: " & rooms.count & " 个直播间"
End Sub

Private Sub UpdateGridWithBuffer(ByRef grid As MSHFlexGrid, ByRef rooms As Collection)
    On Error Resume Next
    
    Dim roomCount As Long
    roomCount = rooms.count
    
    ' 调整缓冲区大小
    If roomCount <> m_lastRoomCount Then
        m_dataBuffer.Reinitialize roomCount + 1, 15
        m_lastRoomCount = roomCount
    End If
    
    Dim totalUpdates As Long
    totalUpdates = 0
    
    ' 处理空数据情况
    If roomCount = 0 Then
        While grid.rows > 1
            grid.RemoveItem grid.rows - 1
        Wend
        Exit Sub
    End If
    
    ' 禁止重绘并保存选择状态
    grid.Redraw = False
    Dim savedRow As Long, savedCol As Long, savedRowSel As Long, savedColSel As Long
    savedRow = grid.row
    savedCol = grid.col
    savedRowSel = grid.RowSel
    savedColSel = grid.ColSel
    
    ' 调整网格行数
    Dim requiredRows As Long
    requiredRows = roomCount + 1
    
    If grid.rows < requiredRows Then
        grid.rows = requiredRows
    ElseIf grid.rows > requiredRows Then
        While grid.rows > requiredRows
            grid.RemoveItem grid.rows - 1
        Wend
    End If
    
    ' 更新单元格数据
    Dim i As Long
    Dim room As LiveRoom
    i = 1
    
    For Each room In rooms
        If i >= grid.rows Then
            grid.rows = grid.rows + 1
        End If
        
        Dim cellData() As String
        Dim rowColor As Long
        m_gridHelper.CreateRowDataFromRoom room, cellData, rowColor
        
        Dim col As Long
        For col = 0 To 14
            Dim cellText As String
            cellText = cellData(col)
            
            ' 修正：使用正确的参数数量（5个参数）
            If m_dataBuffer.UpdateCellIfChanged(i, col + 1, cellText, vbBlack, rowColor) Then
                grid.TextMatrix(i, col) = cellText
                totalUpdates = totalUpdates + 1
            End If
            
            grid.row = i
            grid.col = col
            grid.CellBackColor = rowColor
        Next col
        
        i = i + 1
    Next room
    
    ' 恢复选择状态和重绘
    grid.row = savedRow
    grid.col = savedCol
    grid.RowSel = savedRowSel
    grid.ColSel = savedColSel
    grid.Redraw = True
    
    ' 调试信息
    If totalUpdates > 0 Then
        Debug.Print ">>> 网格更新完成: " & totalUpdates & " 个单元格已更新"
    End If
End Sub

' ========== 排序相关方法 ==========

' 保持排序显示的辅助方法
Private Sub MaintainSortedDisplay(ByRef rooms As Collection)
    If m_currentSortedRooms Is Nothing Then
        UpdateGridWithBuffer MSHFlexGrid, rooms
        Exit Sub
    End If
    
    ' 根据当前的排序顺序重新构建集合
    Dim maintainedRooms As New Collection
    Dim i As Long
    For i = 1 To m_currentSortedRooms.count
        Dim web_rid As String
        web_rid = m_currentSortedRooms(i).web_rid
        
        Dim room As LiveRoom
        Set room = rooms(web_rid)
        If Not room Is Nothing Then
            maintainedRooms.Add room
        End If
    Next i
    
    ' 使用维护的顺序更新显示
    UpdateGridWithSortedData maintainedRooms
    Set m_currentSortedRooms = maintainedRooms
End Sub

' 专门用于排序数据的更新方法
Private Sub UpdateGridWithSortedData(ByRef sortedRooms As Collection)
    On Error Resume Next
    
    Dim roomCount As Long
    roomCount = sortedRooms.count
    
    ' 调整缓冲区大小（确保一致）
    If roomCount <> m_lastRoomCount Then
        m_dataBuffer.Reinitialize roomCount + 1, 15
        m_lastRoomCount = roomCount
    End If
    
    ' 禁止重绘
    MSHFlexGrid.Redraw = False
    
    ' 保存选择状态
    Dim savedRow As Long, savedCol As Long
    savedRow = MSHFlexGrid.row
    savedCol = MSHFlexGrid.col
    
    ' 调整网格行数
    Dim requiredRows As Long
    requiredRows = roomCount + 1
    
    If MSHFlexGrid.rows < requiredRows Then
        MSHFlexGrid.rows = requiredRows
    ElseIf MSHFlexGrid.rows > requiredRows Then
        While MSHFlexGrid.rows > requiredRows
            MSHFlexGrid.RemoveItem MSHFlexGrid.rows - 1
        Wend
    End If
    
    ' 更新单元格数据
    Dim i As Long
    Dim room As LiveRoom
    i = 1
    
    For Each room In sortedRooms
        If i >= MSHFlexGrid.rows Then
            MSHFlexGrid.rows = MSHFlexGrid.rows + 1
        End If
        
        Dim cellData() As String
        Dim rowColor As Long
        m_gridHelper.CreateRowDataFromRoom room, cellData, rowColor
        
        Dim col As Long
        For col = 0 To 14
            Dim cellText As String
            cellText = cellData(col)
            
            ' 强制更新所有单元格（排序后）
            MSHFlexGrid.TextMatrix(i, col) = cellText
            MSHFlexGrid.row = i
            MSHFlexGrid.col = col
            MSHFlexGrid.CellBackColor = rowColor
            
            ' 更新缓冲区
            m_dataBuffer.UpdateCellIfChanged i, col + 1, cellText, vbBlack, rowColor
        Next col
        
        i = i + 1
    Next room
    
    ' 恢复选择状态
    If savedRow < MSHFlexGrid.rows Then
        MSHFlexGrid.row = savedRow
        MSHFlexGrid.col = savedCol
    End If
    
    MSHFlexGrid.Redraw = True
    UpdateColumnHeaders
End Sub

' 重置排序状态
Public Sub ResetSort()
    m_isSorted = False
    Set m_currentSortedRooms = Nothing
    m_sortColumn = -1
    m_sortAscending = True
    m_lastSortColumn = -1
End Sub

' ========== 网格事件处理 ==========



Private Sub MSHFlexGrid_Click()
    EnsureFullRowSelection MSHFlexGrid
End Sub

Private Sub MSHFlexGrid_RowColChange()
    EnsureFullRowSelection MSHFlexGrid
End Sub

Private Sub MSHFlexGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim row As Long, col As Long
    row = MSHFlexGrid.MouseRow
    col = MSHFlexGrid.MouseCol
    
    If row = 0 And col >= 0 Then
        MSHFlexGrid_ClickHeader col
    End If
End Sub

Private Sub MSHFlexGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim currentRow As Long
    currentRow = MSHFlexGrid.MouseRow
    
    If currentRow < MSHFlexGrid.FixedRows Or currentRow >= MSHFlexGrid.rows Then
        Exit Sub
    End If
    
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    
    If MSHFlexGrid.row <> currentRow Then
        With MSHFlexGrid
            .SelectionMode = flexSelectionByRow
            .HighLight = flexHighlightAlways
            .row = currentRow
            .RowSel = currentRow
            .col = 0
            .ColSel = .cols - 1
        End With
    End If
End Sub

' ========== 右键菜单事件处理 ==========

Private Sub mnuEnableRoom_Click()
    ' 修改：使用菜单选中的直播间ID
    Dim web_rid As String
    web_rid = GetMenuSelectedWebRid()
    If web_rid = "" Then
        web_rid = GetSelectedWebRid(MSHFlexGrid) ' 备用方案
    End If
    HandleEnableRoom web_rid
End Sub

Private Sub mnuDisableRoom_Click()
    ' 修改：使用菜单选中的直播间ID
    Dim web_rid As String
    web_rid = GetMenuSelectedWebRid()
    If web_rid = "" Then
        web_rid = GetSelectedWebRid(MSHFlexGrid) ' 备用方案
    End If
    HandleDisableRoom web_rid
End Sub

Private Sub mnuEditRoom_Click()
    ' 修改：使用菜单选中的直播间ID
    Dim web_rid As String
    web_rid = GetMenuSelectedWebRid()
    If web_rid = "" Then
        web_rid = GetSelectedWebRid(MSHFlexGrid) ' 备用方案
    End If
    HandleEditRoom web_rid
End Sub

Private Sub mnuRemoveRoom_Click()
    ' 修改：使用菜单选中的直播间ID
    Dim web_rid As String
    web_rid = GetMenuSelectedWebRid()
    If web_rid = "" Then
        web_rid = GetSelectedWebRid(MSHFlexGrid) ' 备用方案
    End If
    HandleRemoveRoom web_rid
End Sub

Private Sub mnuIncreasePriority_Click()
    ' 修改：使用菜单选中的直播间ID
    Dim web_rid As String
    web_rid = GetMenuSelectedWebRid()
    If web_rid = "" Then
        web_rid = GetSelectedWebRid(MSHFlexGrid) ' 备用方案
    End If
    HandleIncreasePriority web_rid
End Sub

Private Sub mnuDecreasePriority_Click()
    ' 修改：使用菜单选中的直播间ID
    Dim web_rid As String
    web_rid = GetMenuSelectedWebRid()
    If web_rid = "" Then
        web_rid = GetSelectedWebRid(MSHFlexGrid) ' 备用方案
    End If
    HandleDecreasePriority web_rid
End Sub
' ========== 按钮事件处理 ==========

Private Sub Command2_Click()
    With g_liveroomManager
    
        .AddLiveroomSimple "400994521664", , 5, False, "8:00", "23:00", tpMedium, True
        .AddLiveroomSimple "503356611126", , 3, False, "10:00", "22:00", tpLow, True
        .AddLiveroomSimple "418403004842", , 3, True, , , tpHigh, True
        .AddLiveroomSimple "36947836004", , 3, True, , , tpMedium, True
        .AddLiveroomSimple "973322214270", , 3, True, , , tpLow, True
        
    End With
    
    LogMessage "已使用简化方法添加直播间到全局集合"
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
End Sub

Private Sub cmdStopRecordTask_Click()
    On Error Resume Next
    
    If MSHFlexGrid.row < 1 Then
        LogMessage "请先选择一个直播间"
        Exit Sub
    End If
    
    Dim web_rid As String
    web_rid = GetSelectedWebRid(MSHFlexGrid)
    
    If web_rid = "" Then
        LogMessage "未获取到有效的直播间ID"
        Exit Sub
    End If
    
    Dim room As LiveRoom
    Set room = g_liveroomManager.GetLiveroom(web_rid)
    
    If room Is Nothing Then
        LogMessage "未找到对应的直播间: " & web_rid
        Exit Sub
    End If
    
    room.DisableRoom
    
    If StopRecordTask(web_rid) Then
        LogMessage "已发送停止命令到录制任务并禁用直播间: " & web_rid
    Else
        LogMessage "停止录制任务失败: " & web_rid
    End If
    
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
End Sub

Private Sub cmdStopAllRecords_Click()
    On Error Resume Next
    
    Dim recordingRooms As Collection
    Set recordingRooms = g_liveroomManager.GetLiveroomsByTaskStatus("RECORD", True)
    
    Dim room As LiveRoom
    For Each room In recordingRooms
        room.DisableRoom
    Next room
    
    If StopAllRecordTasks() Then
        LogMessage "已向所有录制任务发送停止命令并禁用对应的直播间"
    Else
        LogMessage "停止所有录制任务失败"
    End If
    
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
End Sub

' ========== 排序功能 ==========

Private Sub MSHFlexGrid_ClickHeader(ByVal col As Long)
    On Error Resume Next
    If MSHFlexGrid.MouseRow <> 0 Then Exit Sub
    
    If col = m_lastSortColumn Then
        m_sortAscending = Not m_sortAscending
    Else
        m_lastSortColumn = col
        m_sortAscending = True
    End If
    
    SortGridByColumn col, m_sortAscending
    UpdateColumnHeaders
End Sub

Private Sub SortGridByColumn(ByVal col As Long, ByVal ascending As Boolean)
    On Error Resume Next
    
    Dim rooms As Collection
    Set rooms = M_LiveroomManager.GetAllLiverooms()
    
    If rooms.count = 0 Then Exit Sub
    
    ' 创建 web_rid 数组用于排序
    Dim webRidArray() As String
    ReDim webRidArray(1 To rooms.count)
    
    Dim i As Long
    i = 1
    Dim room As LiveRoom
    For Each room In rooms
        webRidArray(i) = room.web_rid
        i = i + 1
    Next room
    
    ' 根据列选择排序字段
    Select Case col
        Case 0: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "AnchorName", ascending
        Case 1: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "web_rid", ascending
        Case 2: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "room_title", ascending
        Case 3: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "roomStatus", ascending
        Case 4: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "enabled", ascending
        Case 5: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "ParseTaskStatus", ascending
        Case 6: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "isRecording", ascending
        Case 7: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "taskPriority", ascending
        Case 8: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "resolution", ascending
        Case 9: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "currentBitrate", ascending
        Case 10: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "currentFPS", ascending
        Case 11: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "currentFileSize", ascending
        Case 12: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "currentTime", ascending
        Case 13: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "LastCheckTime", ascending
        Case 14: QuickSortWebRidArray webRidArray, 1, UBound(webRidArray), "watchTime", ascending
    End Select
    
    ' 创建排序后的直播间集合
    Dim sortedRooms As New Collection
    For i = 1 To UBound(webRidArray)
        Set room = M_LiveroomManager.GetLiveroom(webRidArray(i))
        If Not room Is Nothing Then
            sortedRooms.Add room
        End If
    Next i
    
    ' 关键修复：完全重置数据缓冲区
    m_dataBuffer.Reinitialize sortedRooms.count + 1, 15
    m_lastRoomCount = sortedRooms.count
    
    ' 设置排序状态标志
    m_isSorted = True
    Set m_currentSortedRooms = sortedRooms
    
    ' 更新网格显示
    UpdateGridWithSortedData sortedRooms
    LogMessage "按列 " & col & " 排序 (" & IIf(ascending, "升序", "降序") & ")"
End Sub

Private Sub QuickSortRoomArray(ByRef roomArray() As Long, ByVal low As Long, ByVal high As Long, _
                              ByVal sortField As String, ByVal ascending As Boolean)
    On Error Resume Next
    If low < high Then
        Dim pivot As Long
        pivot = PartitionRoomArray(roomArray, low, high, sortField, ascending)
        QuickSortRoomArray roomArray, low, pivot - 1, sortField, ascending
        QuickSortRoomArray roomArray, pivot + 1, high, sortField, ascending
    End If
End Sub

Private Function PartitionRoomArray(ByRef roomArray() As Long, ByVal low As Long, ByVal high As Long, _
                                   ByVal sortField As String, ByVal ascending As Boolean) As Long
    On Error Resume Next
    
    Dim pivotWebRid As Long
    pivotWebRid = roomArray(high)
    Dim pivotRoom As LiveRoom
    Set pivotRoom = M_LiveroomManager.GetLiveroom(pivotWebRid)
    
    Dim i As Long
    i = low - 1
    
    Dim j As Long
    For j = low To high - 1
        Dim currentRoom As LiveRoom
        Set currentRoom = M_LiveroomManager.GetLiveroom(roomArray(j))
        
        If CompareRooms(currentRoom, pivotRoom, sortField, ascending) <= 0 Then
            i = i + 1
            SwapRoomArrayItems roomArray, i, j
        End If
    Next j
    
    SwapRoomArrayItems roomArray, i + 1, high
    PartitionRoomArray = i + 1
End Function

Private Sub SwapRoomArrayItems(ByRef roomArray() As Long, ByVal i As Long, ByVal j As Long)
    On Error Resume Next
    Dim temp As Long
    temp = roomArray(i)
    roomArray(i) = roomArray(j)
    roomArray(j) = temp
End Sub

Private Function CompareRooms(ByRef room1 As LiveRoom, ByRef room2 As LiveRoom, _
                             ByVal sortField As String, ByVal ascending As Boolean) As Integer
    On Error Resume Next
    
    Dim result As Integer
    Dim value1 As String, value2 As String
    Dim num1 As Double, num2 As Double
    
    If room1 Is Nothing Or room2 Is Nothing Then
        CompareRooms = 0
        Exit Function
    End If
    
    Select Case sortField
        Case "AnchorName"
            value1 = room1.AnchorName
            value2 = room2.AnchorName
        Case "web_rid"
            value1 = room1.web_rid
            value2 = room2.web_rid
        Case "room_title"
            value1 = room1.room_title
            value2 = room2.room_title
        Case "roomStatus"
            value1 = room1.roomStatus
            value2 = room2.roomStatus
        Case "enabled"
            value1 = IIf(room1.enabled, "1", "0")
            value2 = IIf(room2.enabled, "1", "0")
        Case "ParseTaskStatus"
            value1 = room1.ParseTaskStatus
            value2 = room2.ParseTaskStatus
        Case "isRecording"
            value1 = IIf(room1.isRecording, "1", "0")
            value2 = IIf(room2.isRecording, "1", "0")
        Case "taskPriority"
            num1 = room1.taskPriority
            num2 = room2.taskPriority
            If num1 < num2 Then result = -1 Else If num1 > num2 Then result = 1 Else result = 0
            If Not ascending Then result = -result
            CompareRooms = result
            Exit Function
        Case "resolution"
            value1 = CStr(room1.streamWidth) & "x" & CStr(room1.streamHeight)
            value2 = CStr(room2.streamWidth) & "x" & CStr(room2.streamHeight)
        Case "currentBitrate"
            value1 = room1.currentBitrate
            value2 = room2.currentBitrate
        Case "currentFPS"
            value1 = room1.currentFPS
            value2 = room2.currentFPS
        Case "currentFileSize"
            value1 = room1.currentFileSize
            value2 = room2.currentFileSize
        Case "currentTime"
            value1 = room1.currentTime
            value2 = room2.currentTime
        Case "LastCheckTime"
            If room1.LastCheckTime < room2.LastCheckTime Then result = -1 Else If room1.LastCheckTime > room2.LastCheckTime Then result = 1 Else result = 0
            If Not ascending Then result = -result
            CompareRooms = result
            Exit Function
        Case "watchTime"
            value1 = GetWatchTimeDisplay(room1)
            value2 = GetWatchTimeDisplay(room2)
        Case Else
            value1 = ""
            value2 = ""
    End Select
    
    If value1 < value2 Then
        result = -1
    ElseIf value1 > value2 Then
        result = 1
    Else
        result = 0
    End If
    
    If Not ascending Then
        result = -result
    End If
    
    CompareRooms = result
End Function

Private Sub UpdateColumnHeaders()
    On Error Resume Next
    Dim i As Long
    For i = 0 To MSHFlexGrid.cols - 1
        Dim headerText As String
        headerText = MSHFlexGrid.TextMatrix(0, i)
        
        If Right(headerText, 2) = " ▲" Or Right(headerText, 2) = " ▼" Then
            headerText = Left(headerText, Len(headerText) - 2)
        End If
        
        If i = m_lastSortColumn Then
            If m_sortAscending Then
                headerText = headerText & " ▲"
            Else
                headerText = headerText & " ▼"
            End If
        End If
        
        MSHFlexGrid.TextMatrix(0, i) = headerText
    Next i
End Sub

Private Sub QuickSortWebRidArray(ByRef webRidArray() As String, ByVal low As Long, ByVal high As Long, _
                              ByVal sortField As String, ByVal ascending As Boolean)
    On Error Resume Next
    If low < high Then
        Dim pivot As Long
        pivot = PartitionWebRidArray(webRidArray, low, high, sortField, ascending)
        QuickSortWebRidArray webRidArray, low, pivot - 1, sortField, ascending
        QuickSortWebRidArray webRidArray, pivot + 1, high, sortField, ascending
    End If
End Sub

Private Function PartitionWebRidArray(ByRef webRidArray() As String, ByVal low As Long, ByVal high As Long, _
                                   ByVal sortField As String, ByVal ascending As Boolean) As Long
    On Error Resume Next
    
    Dim pivotWebRid As String
    pivotWebRid = webRidArray(high)
    Dim pivotRoom As LiveRoom
    Set pivotRoom = M_LiveroomManager.GetLiveroom(pivotWebRid)
    
    Dim i As Long
    i = low - 1
    
    Dim j As Long
    For j = low To high - 1
        Dim currentRoom As LiveRoom
        Set currentRoom = M_LiveroomManager.GetLiveroom(webRidArray(j))
        
        If CompareRooms(currentRoom, pivotRoom, sortField, ascending) <= 0 Then
            i = i + 1
            SwapWebRidArrayItems webRidArray, i, j
        End If
    Next j
    
    SwapWebRidArrayItems webRidArray, i + 1, high
    PartitionWebRidArray = i + 1
End Function

Private Sub SwapWebRidArrayItems(ByRef webRidArray() As String, ByVal i As Long, ByVal j As Long)
    On Error Resume Next
    Dim temp As String
    temp = webRidArray(i)
    webRidArray(i) = webRidArray(j)
    webRidArray(j) = temp
End Sub

' ========== 事件处理 ==========

Private Sub M_LiveroomManager_LiveroomAdded(ByRef room As LiveRoom)
    room.MarkAsChanged
    ResetSort ' 重置排序
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    UpdateStatsLabel
    LogMessage "直播间添加: " & room.AnchorName & " (" & room.web_rid & ")"
End Sub

Private Sub M_LiveroomManager_LiveroomRemoved(ByRef room As LiveRoom)
    ResetSort ' 重置排序
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    UpdateStatsLabel
    LogMessage "直播间移除: " & room.AnchorName & " (" & room.web_rid & ")"
End Sub

Private Sub M_LiveroomManager_LiveroomEnabled(ByRef room As LiveRoom)
    room.MarkAsChanged
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    LogMessage "直播间启用: " & room.AnchorName
End Sub

Private Sub M_LiveroomManager_LiveroomDisabled(ByRef room As LiveRoom)
    room.MarkAsChanged
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    LogMessage "直播间禁用: " & room.AnchorName
End Sub

Private Sub M_LiveroomManager_ParseTaskTriggered(ByRef room As LiveRoom)
    room.MarkAsChanged
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    LogMessage "解析任务触发: " & room.AnchorName
End Sub

Private Sub M_LiveroomManager_RecordTaskTriggered(ByRef room As LiveRoom)
    room.MarkAsChanged
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    LogMessage "录制任务触发: " & room.AnchorName
End Sub

Private Sub M_LiveroomManager_PollingCompleted()
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
End Sub

Private Sub m_scheduler_TaskAdded(ByRef task As task)
    MarkTaskRoomChanged task
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    LogMessage "任务添加: " & task.taskName & " (ID: " & task.taskID & ")"
End Sub

Private Sub m_scheduler_TaskStarted(ByRef task As task)
    MarkTaskRoomChanged task
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    LogMessage "任务开始: " & task.taskName & " (ID: " & task.taskID & ")"
End Sub

Private Sub m_scheduler_TaskCompleted(ByRef task As task)
    MarkTaskRoomChanged task
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    LogMessage "任务完成: " & task.taskName
End Sub

Private Sub m_scheduler_TaskFailed(ByRef task As task, ByVal errorMsg As String)
    MarkTaskRoomChanged task
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    LogMessage "任务失败: " & task.taskName & " - " & errorMsg
End Sub

Private Sub m_scheduler_TaskOutput(ByRef task As task, ByVal outputData As String, ByVal isErrorOutput As Boolean)
    Dim outputType As String
    If isErrorOutput Then
        outputType = "错误输出"
    Else
        outputType = "标准输出"
    End If
    
    MarkTaskRoomChanged task
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    
    LogMessage "任务输出 [" & outputType & "]: " & task.taskName & " - " & g_liveroomManager.GetLiveroom(task.web_rid).AnchorName & " " & g_liveroomManager.GetLiveroom(task.web_rid).currentBitrate
End Sub

Private Sub m_scheduler_TaskTimeout(ByRef task As task)
    MarkTaskRoomChanged task
    DirectUpdateFromManager M_LiveroomManager.GetAllLiverooms()
    LogMessage "任务超时: " & task.taskName & " (ID: " & task.taskID & ") - 进程已被终止"
    Debug.Print "任务超时: " & task.taskName & " (ID: " & task.taskID & ") - 进程已被终止"
End Sub

' ========== 辅助方法 ==========

Private Sub MarkTaskRoomChanged(ByRef task As task)
    On Error Resume Next
    If task.web_rid <> "" Then
        Dim room As LiveRoom
        Set room = g_liveroomManager.GetLiveroom(task.web_rid)
        If Not room Is Nothing Then
            room.MarkAsChanged
        End If
    End If
End Sub

Private Sub LogMessage(ByVal message As String)
    Debug.Print ">>> " & Format(Now, "hh:mm:ss") & " - " & message
End Sub

' 兼容性方法
Public Sub UpdateLiveroomListView()
    On Error Resume Next
    m_gridHelper.ForceFullRowSelection MSHFlexGrid
    Dim rooms As Collection
    Set rooms = M_LiveroomManager.GetAllLiverooms()
    DirectUpdateFromManager rooms
    EnsureFullRowSelection MSHFlexGrid
    UpdateStatsLabel
End Sub

Private Sub OptimizedUpdateLiveroomListView()
    On Error Resume Next
    Dim rooms As Collection
    Set rooms = M_LiveroomManager.GetAllLiverooms()
    m_gridHelper.ForceFullRowSelection MSHFlexGrid
    DirectUpdateFromManager rooms
    EnsureFullRowSelection MSHFlexGrid
    UpdateStatsLabel
End Sub


Private Sub MSHFlexGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = vbRightButton Then
        ' 获取鼠标位置对应的行
        Dim hitRow As Long
        hitRow = MSHFlexGrid.MouseRow
        
        If hitRow >= 1 Then ' 确保点击的是数据行（不是表头）
            ' 选中该行
            MSHFlexGrid.row = hitRow
            MSHFlexGrid.col = 0
            MSHFlexGrid.RowSel = hitRow
            MSHFlexGrid.ColSel = MSHFlexGrid.cols - 1
            
            ' 获取该行的直播间ID并存储到全局变量
            Dim web_rid As String
            web_rid = MSHFlexGrid.TextMatrix(hitRow, 1)
            SetMenuSelectedWebRid web_rid
            
            Debug.Print ">>> 右键菜单选中行: " & hitRow & ", 直播间ID: " & web_rid
            
            ' 直接弹出菜单
            PopupMenu mnuContextMenu
        Else
            Debug.Print ">>> 右键点击位置无效，行号: " & hitRow
        End If
    End If
End Sub
