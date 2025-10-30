' GridHelper.bas
' 说明：
'   - 该模块整合并精简原工程中关于 MSHFlexGrid 的各种 helper（CMSHFlexGridHelper、M_GridUtils、M_GridManagement 等）
'   - 提供：初始化列、自动列宽、批量更新（Begin/EndBatchUpdate）、缓冲更新（配合 CGridDataBuffer）、若干安全选择函数
'   - 目标：将所有 Grid 相关维护集中，减少重复代码，提升可维护性与性能。
' 使用说明：
'   - 将此文件放入工程的 .bas 模块（保存为 ANSI/GBK）
'   - 在 frmMain.Form_Load 中使用 InitializeGrid(MSHFlexGrid)
'   - 在刷新大量行前使用 BeginBatchUpdate / EndBatchUpdate 或使用 UpdateGridWithBuffer
'   - 依赖：全局函数 CalculateRowBackgroundColor、GetPriorityText、GetWatchTimeDisplay（可放在 M__Global 或 SchedulerUtils）
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Const WM_SETREDRAW As Long = &HB

Public Const flexSelectionByRow As Long = 1
Public Const flexFocusNone As Long = 0
Public Const flexResizeColumns As Long = 1
Public Const flexScrollBarVertical As Long = 2

Private ratios() As Double

Private Sub InitRatios()
    On Error Resume Next
    If Not (Not ratios) Then Exit Sub
    ReDim ratios(0 To 15)
    ratios(0) = 0.10
    ratios(1) = 0.03
    ratios(2) = 0.06
    ratios(3) = 0.10
    ratios(4) = 0.05
    ratios(5) = 0.06
    ratios(6) = 0.06
    ratios(7) = 0.06
    ratios(8) = 0.05
    ratios(9) = 0.06
    ratios(10) = 0.08
    ratios(11) = 0.04
    ratios(12) = 0.06
    ratios(13) = 0.06
    ratios(14) = 0.06
    ratios(15) = 0.06
End Sub

Public Sub InitializeGrid(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    With grid
        .Clear
        .rows = 1
        .cols = 16
        .FixedRows = 1
        .FixedCols = 0

        .TextMatrix(0, 0) = "主播"
        .TextMatrix(0, 1) = "直播ID"
        .TextMatrix(0, 2) = "Tag"
        .TextMatrix(0, 3) = "房间标题"
        .TextMatrix(0, 4) = "房间状态"
        .TextMatrix(0, 5) = "是否启用"
        .TextMatrix(0, 6) = "解析状态"
        .TextMatrix(0, 7) = "录制状态"
        .TextMatrix(0, 8) = "优先级"
        .TextMatrix(0, 9) = "分辨率"
        .TextMatrix(0, 10) = "码率"
        .TextMatrix(0, 11) = "FPS"
        .TextMatrix(0, 12) = "文件大小"
        .TextMatrix(0, 13) = "时长"
        .TextMatrix(0, 14) = "最后检查"
        .TextMatrix(0, 15) = "可观看时间"

        .HighLight = 2
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ScrollBars = flexScrollBarVertical
        .FocusRect = flexFocusNone
    End With

    AutoSizeColumns grid
End Sub

Public Sub AutoSizeColumns(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    InitRatios
    Dim totalWidth As Long
    totalWidth = grid.Width - 60
    If totalWidth < 500 Then totalWidth = 500

    Dim i As Long
    For i = 0 To grid.cols - 1
        If i <= UBound(ratios) Then
            grid.ColWidth(i) = totalWidth * ratios(i)
        Else
            grid.ColWidth(i) = 50
        End If
    Next i

    Dim diff As Long
    diff = totalWidth - GetTotalColumnsWidth(grid)
    If grid.cols > 0 Then grid.ColWidth(grid.cols - 1) = grid.ColWidth(grid.cols - 1) + diff
End Sub

Public Function GetTotalColumnsWidth(ByRef grid As MSHFlexGrid) As Long
    On Error Resume Next
    Dim i As Long, s As Long
    For i = 0 To grid.cols - 1
        s = s + grid.ColWidth(i)
    Next i
    GetTotalColumnsWidth = s
End Function

Public Sub BeginBatchUpdate(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    SendMessage grid.hwnd, WM_SETREDRAW, 0, 0
    LockWindowUpdate grid.hwnd
    grid.Redraw = False
End Sub

Public Sub EndBatchUpdate(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    grid.Redraw = True
    LockWindowUpdate 0
    SendMessage grid.hwnd, WM_SETREDRAW, 1, 0
    grid.Refresh
End Sub

Public Sub UpdateGridWithBuffer(ByRef grid As MSHFlexGrid, ByRef rooms As Collection, ByRef buffer As CGridDataBuffer)
    On Error Resume Next
    If rooms Is Nothing Then Exit Sub

    Dim roomCount As Long
    roomCount = rooms.count

    BeginBatchUpdate grid

    If grid.rows <> roomCount + 1 Then
        AdjustGridRowCount grid, roomCount + 1
    End If

    Dim i As Long
    i = 1
    Dim room As LiveRoom
    For Each room In rooms
        If i >= grid.rows Then grid.rows = grid.rows + 1

        Dim cellData(0 To 15) As String
        Dim rowColor As Long
        CreateRowDataFromRoom room, cellData, rowColor

        Dim col As Long
        For col = 0 To 15
            If buffer.UpdateCellIfChanged(i, col + 1, cellData(col), vbBlack, rowColor) Then
                grid.TextMatrix(i, col) = cellData(col)
            End If
            grid.row = i
            grid.col = col
            grid.CellBackColor = rowColor
        Next col

        i = i + 1
    Next room

    EndBatchUpdate grid
End Sub

Public Sub AdjustGridRowCount(ByRef grid As MSHFlexGrid, ByVal requiredRows As Long)
    On Error Resume Next
    If grid.rows < requiredRows Then
        grid.rows = requiredRows
    ElseIf grid.rows > requiredRows Then
        grid.Redraw = False
        While grid.rows > requiredRows
            grid.RemoveItem grid.rows - 1
        Wend
        grid.Redraw = True
    End If
End Sub

Public Sub CreateRowDataFromRoom(ByRef room As LiveRoom, ByRef cellData() As String, ByRef rowColor As Long)
    On Error Resume Next
    ReDim cellData(0 To 15)

    cellData(0) = room.AnchorName
    cellData(1) = room.web_rid
    cellData(2) = room.Tag
    cellData(3) = room.room_title
    cellData(4) = room.roomStatus
    cellData(5) = IIf(room.enabled, "启用", "禁用")
    cellData(6) = room.ParseTaskStatus
    cellData(7) = room.RecordTaskStatus
    cellData(8) = GetPriorityText(room.taskPriority)
    cellData(9) = IIf(room.streamWidth > 0, room.streamWidth & "x" & room.streamHeight, "")
    cellData(10) = room.currentBitrate
    cellData(11) = room.currentFPS
    cellData(12) = room.currentFileSize
    cellData(13) = room.currentTime
    cellData(14) = IIf(room.LastCheckTime <> 0, Format(room.LastCheckTime, "hh:mm:ss"), "")
    cellData(15) = GetWatchTimeDisplay(room)

    rowColor = CalculateRowBackgroundColor(room)
End Sub

' End of GridHelper.bas

