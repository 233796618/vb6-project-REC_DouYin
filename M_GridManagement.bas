Attribute VB_Name = "M_GridManagement"
Option Explicit

' API 声明
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

' ListView 消息常量
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVS_EX_MULTISELECT As Long = &H4

' FlexGrid 常量
Private Const flexSelectionByRow As Long = 1
Private Const flexFocusNone As Long = 0

' 全局变量存储当前选中的直播间ID
Public g_SelectedWebRidForMenu As String

' 网格选择管理
Public Sub SetSingleRowSelection(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    
    ' 设置 FlexGrid 自身的属性
    grid.SelectionMode = flexSelectionByRow
    grid.AllowBigSelection = False
    grid.FocusRect = flexFocusNone
    
    ' 使用 API 方法来确保单行选择
    Dim style As Long
    style = SendMessage(grid.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    ' 移除多选标志
    style = style And Not LVS_EX_MULTISELECT
    ' 设置新的扩展样式
    SendMessage grid.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal style
    
    Debug.Print ">>> 设置单行选择模式完成"
End Sub

Public Sub EnsureFullRowSelection(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    If grid.row >= 1 Then
        grid.col = 0
        grid.ColSel = grid.cols - 1
    End If
End Sub

Public Sub ClearSelectionFinal(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    
    With grid
        .Redraw = False
        .row = 0
        .col = 0
        .RowSel = 0
        .ColSel = .cols - 1
        .SetFocus
        .Redraw = True
    End With
    
    Debug.Print ">>> 选择已清除（最终方法）"
End Sub

' 右键菜单处理 - 修改：添加鼠标位置行选择
Public Sub HandleGridContextMenu(ByRef grid As MSHFlexGrid, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    
    Dim hitRow As Long
    hitRow = GetRowAtPosition(grid, x, y)
    
    If hitRow >= 1 Then
        ' 选中该行
        grid.row = hitRow
        grid.col = 0
        grid.RowSel = hitRow
        grid.ColSel = grid.cols - 1
        
        ' 获取该行的直播间ID并存储到全局变量
        Dim web_rid As String
        web_rid = grid.TextMatrix(hitRow, 1)
        SetMenuSelectedWebRid web_rid
        
        Debug.Print ">>> 右键菜单选中行: " & hitRow & ", 直播间ID: " & web_rid
    End If
End Sub

' 获取鼠标位置对应的行
Private Function GetRowAtPosition(ByRef grid As MSHFlexGrid, ByVal x As Single, ByVal y As Single) As Long
    On Error Resume Next
    
    ' 保存当前选择状态
    Dim oldRow As Long, oldCol As Long
    oldRow = grid.row
    oldCol = grid.col
    
    ' 临时设置鼠标位置的行
    grid.row = grid.MouseRow
    grid.col = grid.MouseCol
    
    ' 获取鼠标所在行
    Dim hitRow As Long
    hitRow = grid.row
    
    ' 恢复原来的选择状态
    grid.row = oldRow
    grid.col = oldCol
    
    GetRowAtPosition = hitRow
End Function

' 获取选中行的 web_rid
Public Function GetSelectedWebRid(ByRef grid As MSHFlexGrid) As String
    On Error Resume Next
    If grid.row < 1 Then
        GetSelectedWebRid = ""
    Else
        GetSelectedWebRid = grid.TextMatrix(grid.row, 1)
    End If
End Function

' 新增：获取菜单选中的直播间ID
Public Function GetMenuSelectedWebRid() As String
    GetMenuSelectedWebRid = g_SelectedWebRidForMenu
End Function

' 新增：设置菜单选中的直播间ID
Public Sub SetMenuSelectedWebRid(ByVal web_rid As String)
    g_SelectedWebRidForMenu = web_rid
End Sub

' 新增：重置排序状态
Public Sub ResetGridSort(ByRef grid As MSHFlexGrid)
    On Error Resume Next
    ' 调用主窗体的重置排序方法
    If Not frmMain Is Nothing Then
        frmMain.ResetSort
    End If
End Sub

