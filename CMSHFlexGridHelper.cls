VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMSHFlexGridHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' API 声明
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

' 常量
Private Const WM_SETREDRAW As Long = &HB
Private Const flexSelectionByRow As Long = 1
Private Const flexResizeColumns As Long = 1
Private Const flexScrollBarBoth As Long = 3
Private Const flexFocusNone As Long = 0

' 状态变量
Private m_isUpdating As Boolean

' ========== 网格管理方法 ==========

Public Sub InitializeGrid(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    
    With grid
        .Clear
        .rows = 1
        .cols = 15
        
        ' 设置列标题
        .TextMatrix(0, 0) = "主播名"
        .TextMatrix(0, 1) = "房间ID"
        .TextMatrix(0, 2) = "房间标题"
        .TextMatrix(0, 3) = "房间状态"
        .TextMatrix(0, 4) = "启用状态"
        .TextMatrix(0, 5) = "解析状态"
        .TextMatrix(0, 6) = "录制状态"
        .TextMatrix(0, 7) = "优先级"
        .TextMatrix(0, 8) = "分辨率"
        .TextMatrix(0, 9) = "码率"
        .TextMatrix(0, 10) = "FPS"
        .TextMatrix(0, 11) = "文件大小"
        .TextMatrix(0, 12) = "录制时间"
        .TextMatrix(0, 13) = "最后检查"
        .TextMatrix(0, 14) = "监控时段"
        
        .FixedRows = 1
        .FixedCols = 0
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarBoth
        .FocusRect = flexFocusNone
    End With
    
    AutoSizeColumns grid
    Debug.Print "MSHFlexGrid 初始化完成"
End Sub

Public Sub AutoSizeColumns(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    
    If m_isUpdating Then Exit Sub
    
    With grid
        Dim totalWidth As Long
        totalWidth = .Width - 350
        
        .ColWidth(0) = totalWidth * 0.1   ' 主播名 10%
        .ColWidth(1) = totalWidth * 0.08  ' 房间ID 8%
        .ColWidth(2) = totalWidth * 0.1   ' 房间标题 15%
        .ColWidth(3) = totalWidth * 0.05  ' 房间状态 7%
        .ColWidth(4) = totalWidth * 0.06  ' 启用状态 6%
        .ColWidth(5) = totalWidth * 0.06  ' 解析状态 7%
        .ColWidth(6) = totalWidth * 0.06  ' 录制状态 7%
        .ColWidth(7) = totalWidth * 0.05  ' 优先级 5%
        .ColWidth(8) = totalWidth * 0.06  ' 分辨率 6%
        .ColWidth(9) = totalWidth * 0.08  ' 码率 5%
        .ColWidth(10) = totalWidth * 0.04 ' FPS 4%
        .ColWidth(11) = totalWidth * 0.06 ' 文件大小 6%
        .ColWidth(12) = totalWidth * 0.06 ' 录制时间 6%
        .ColWidth(13) = totalWidth * 0.06 ' 最后检查 6%
        .ColWidth(14) = totalWidth * 0.06 ' 监控时段 6%
    End With
End Sub

' ========== 更新控制方法 ==========

Public Sub BeginSimpleUpdate(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    If m_isUpdating Then Exit Sub
    
    m_isUpdating = True
    SendMessage grid.hWnd, WM_SETREDRAW, 0, 0
    LockWindowUpdate grid.hWnd
    grid.Redraw = False
End Sub

Public Sub EndSimpleUpdate(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    If Not m_isUpdating Then Exit Sub
    
    grid.Redraw = True
    LockWindowUpdate 0
    SendMessage grid.hWnd, WM_SETREDRAW, 1, 0
    m_isUpdating = False
End Sub

' ========== 数据行处理方法 ==========
Public Sub CreateRowDataFromRoom(ByRef room As LiveRoom, ByRef cellData() As String, ByRef rowColor As Long)
    ' 重新定义数组大小
    ReDim cellData(0 To 14) As String
    
    ' 填充单元格数据
    cellData(0) = room.AnchorName
    cellData(1) = room.web_rid
    cellData(2) = room.room_title
    cellData(3) = room.roomStatus
    cellData(4) = IIf(room.enabled, "启用", "禁用")
    cellData(5) = room.ParseTaskStatus
    cellData(6) = IIf(room.isRecording, "录制中", "未录制")
    cellData(7) = GetPriorityText(room.taskPriority)  ' 调用全局函数
    cellData(8) = IIf(room.streamWidth > 0, room.streamWidth & "x" & room.streamHeight, "")
    cellData(9) = room.currentBitrate
    cellData(10) = room.currentFPS
    cellData(11) = room.currentFileSize
    cellData(12) = room.currentTime
    cellData(13) = IIf(room.LastCheckTime <> 0, Format(room.LastCheckTime, "hh:mm:ss"), "")
    cellData(14) = GetWatchTimeDisplay(room)  ' 调用全局函数
    
    rowColor = CalculateRowBackgroundColor(room)  ' 调用全局函数
End Sub


' ========== 选择状态管理 ==========

Public Function GetSelectedWebRid(ByRef grid As MSHFlexGrid) As String
    On Error Resume Next
    
    If grid.row < 1 Then
        GetSelectedWebRid = ""
    Else
        GetSelectedWebRid = grid.TextMatrix(grid.row, 1)
    End If
End Function

Public Sub ForceFullRowSelection(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    If grid.SelectionMode <> flexSelectionByRow Then
        grid.SelectionMode = flexSelectionByRow
    End If
    
    If grid.row >= 1 Then
        grid.col = 0
        grid.ColSel = grid.cols - 1
    End If
End Sub





Public Property Get IsUpdating() As Boolean
    IsUpdating = m_isUpdating
End Property

