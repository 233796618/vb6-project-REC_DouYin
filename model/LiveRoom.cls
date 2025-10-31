' LiveRoom.cls - 房间业务对象（简化版）
Option Explicit

Public web_rid As String
Public AnchorName As String
Public Tag As String
Public room_title As String
Public roomStatus As String
Public enabled As Boolean
Public taskPriority As Integer

Public Sub InitializeMock()
    web_rid = "0"
    AnchorName = "示例主播"
    Tag = "测试"
    room_title = "示例房间"
    roomStatus = "离线"
    enabled = True
    taskPriority = 1
End Sub

' End of LiveRoom.cls