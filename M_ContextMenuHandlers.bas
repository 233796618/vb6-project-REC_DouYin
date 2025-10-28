Attribute VB_Name = "M_ContextMenuHandlers"
Option Explicit

' 启用直播间
Public Sub HandleEnableRoom(ByVal web_rid As String)
    On Error GoTo ErrorHandler
    
    Debug.Print ">>> 开始处理启用直播间请求，传入web_rid: " & web_rid
    
    If web_rid = "" Then
        ' 尝试从网格获取选中的直播间ID
        web_rid = GetMenuSelectedWebRid()
        If web_rid = "" Then
            web_rid = GetSelectedWebRidFromGrid()
        End If
        
        If web_rid = "" Then
            MsgBox "请先选择一个直播间", vbInformation
            Exit Sub
        End If
    End If
    
    ' 获取直播间对象
    Dim room As LiveRoom
    Set room = g_liveroomManager.GetLiveroom(web_rid)
    
    If room Is Nothing Then
        MsgBox "未找到对应的直播间: " & web_rid, vbExclamation
        Exit Sub
    End If
    
    ' 如果已经启用，显示提示信息
    If room.enabled Then
        MsgBox "直播间 " & web_rid & " 已经处于启用状态", vbInformation
        Exit Sub
    End If
    
    ' 启用直播间
    room.EnableRoom
    
    ' 同时调用 LiveroomManager 的启用方法确保状态同步
    g_liveroomManager.EnableLiveroom web_rid
    
    LogMessage "启用直播间: " & web_rid & " - " & room.AnchorName
    ForceUpdate
    
    Debug.Print ">>> 直播间启用成功: " & web_rid
    Exit Sub
    
ErrorHandler:
    Debug.Print ">>> 启用直播间时出错: " & Err.Description & " (错误号: " & Err.Number & ")"
End Sub

' 禁用直播间
Public Sub HandleDisableRoom(ByVal web_rid As String)
    On Error Resume Next
    If web_rid = "" Then
        ' 尝试从菜单选中的直播间ID获取
        web_rid = GetMenuSelectedWebRid()
        If web_rid = "" Then
            web_rid = GetSelectedWebRidFromGrid()
        End If
        
        If web_rid = "" Then
            MsgBox "请先选择一个直播间", vbInformation
            Exit Sub
        End If
    End If
    
    Dim room As LiveRoom
    Set room = g_liveroomManager.GetLiveroom(web_rid)
    
    If Not room Is Nothing Then
        room.DisableRoom
        If StopRecordTask(web_rid) Then
            LogMessage "已发送停止命令到录制任务并禁用直播间: " & web_rid & " - " & room.AnchorName
        Else
            LogMessage "停止录制任务失败: " & web_rid
        End If
    Else
        MsgBox "未找到对应的直播间: " & web_rid, vbExclamation
    End If
    
    ForceUpdate
End Sub

' 编辑直播间
Public Sub HandleEditRoom(ByVal web_rid As String)
    On Error Resume Next
    If web_rid = "" Then
        web_rid = GetMenuSelectedWebRid()
        If web_rid = "" Then
            web_rid = GetSelectedWebRidFromGrid()
        End If
        
        If web_rid = "" Then
            MsgBox "请先选择一个直播间", vbInformation
            Exit Sub
        End If
    End If
    
    ' 这里可以调用编辑对话框
    MsgBox "编辑直播间: " & web_rid, vbInformation
    LogMessage "编辑直播间: " & web_rid
End Sub

' 删除直播间
Public Sub HandleRemoveRoom(ByVal web_rid As String)
    On Error Resume Next
    If web_rid = "" Then
        web_rid = GetMenuSelectedWebRid()
        If web_rid = "" Then
            web_rid = GetSelectedWebRidFromGrid()
        End If
        
        If web_rid = "" Then
            MsgBox "请先选择一个直播间", vbInformation
            Exit Sub
        End If
    End If
    
    If MsgBox("确定要删除直播间 " & web_rid & " 吗？", vbYesNo + vbQuestion) = vbYes Then
        g_liveroomManager.RemoveLiveroom web_rid
        LogMessage "删除直播间: " & web_rid
        ForceUpdate
    End If
End Sub

' 提高优先级
Public Sub HandleIncreasePriority(ByVal web_rid As String)
    On Error Resume Next
    If web_rid = "" Then
        web_rid = GetMenuSelectedWebRid()
        If web_rid = "" Then
            web_rid = GetSelectedWebRidFromGrid()
        End If
        
        If web_rid = "" Then
            MsgBox "请先选择一个直播间", vbInformation
            Exit Sub
        End If
    End If
    
    Dim room As LiveRoom
    Set room = g_liveroomManager.GetLiveroom(web_rid)
    
    If Not room Is Nothing And room.taskPriority > tpHigh Then
        room.taskPriority = room.taskPriority - 1
        room.MarkAsChanged
        ForceUpdate
        
        Dim priorityName As String
        priorityName = GetPriorityText(room.taskPriority)
        LogMessage "提高直播间优先级: " & web_rid & " -> " & priorityName & " - " & room.AnchorName
    ElseIf room.taskPriority = tpHigh Then
        MsgBox "已经是最高优先级", vbInformation
    End If
End Sub

' 降低优先级
Public Sub HandleDecreasePriority(ByVal web_rid As String)
    On Error Resume Next
    If web_rid = "" Then
        web_rid = GetMenuSelectedWebRid()
        If web_rid = "" Then
            web_rid = GetSelectedWebRidFromGrid()
        End If
        
        If web_rid = "" Then
            MsgBox "请先选择一个直播间", vbInformation
            Exit Sub
        End If
    End If
    
    Dim room As LiveRoom
    Set room = g_liveroomManager.GetLiveroom(web_rid)
    
    If Not room Is Nothing And room.taskPriority < tpLow Then
        room.taskPriority = room.taskPriority + 1
        room.MarkAsChanged
        ForceUpdate
        
        Dim priorityName As String
        priorityName = GetPriorityText(room.taskPriority)
        LogMessage "降低直播间优先级: " & web_rid & " -> " & priorityName & " - " & room.AnchorName
    ElseIf room.taskPriority = tpLow Then
        MsgBox "已经是最低优先级", vbInformation
    End If
End Sub
