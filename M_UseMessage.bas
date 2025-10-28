Attribute VB_Name = "M_UseMessage"
Option Explicit

Function UseJsonMessage(JSON As String)
    On Error GoTo ErrorHandler
    
    Dim JB As New JsonBag
    JB.JSON = JSON
    Dim web_rid As String
    web_rid = JB.item("Web_rid")
    
    '发现被限制IP，用cookie可能会好一些
    If web_rid <> "" Then
        '很可能是频率太高，被限IP了
        Debug.Print "解析不到数据 被限IP了"
    End If
    
    Dim lr As LiveRoom
    
    Set lr = g_liveroomManager.GetLiveroom(web_rid)
    If Not lr Is Nothing Then
        ' 更新直播间信息
        lr.AnchorName = JB.item("AnchorName")
        lr.room_title = JB.item("Room_Title")
        lr.streamUrl = JB.item("streamUrl")
        lr.roomStatus = JB.item("RoomStatus")
        lr.streamWidth = JB.item("streamWidth")
        lr.streamHeight = JB.item("streamHeight")
        lr.parseTime = JB.item("parseTime")
        

        
        ' 标记数据变化
        lr.MarkAsChanged
        
        ' 标记解析任务完成
        lr.MarkParseTaskCompleted
        
        Debug.Print "直播间 " & web_rid & ": 收到解析结果，标记解析任务完成"
        
        ' 记录调试信息
        Debug.Print "========================================"
        Debug.Print "收到解析结果 - 直播间: " & web_rid
        Debug.Print "主播: " & lr.AnchorName
        Debug.Print "房间标题: " & lr.room_title
        Debug.Print "房间状态: " & lr.roomStatus
        Debug.Print "流地址: " & lr.streamUrl
        Debug.Print "分辨率: " & lr.streamWidth & "x" & lr.streamHeight
        Debug.Print "解析时间: " & lr.parseTime
        Debug.Print "========================================"
    Else
        Debug.Print "错误: 未找到对应的直播间 - Web_rid: " & web_rid
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "UseJsonMessage 错误: " & Err.Description & " (错误号: " & Err.Number & ")"
End Function
