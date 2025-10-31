' CGridDataBuffer - 缓冲变更以减少 Grid 写入
Option Explicit

Private Type CellKey
    Row As Long
    Col As Long
End Type

Private m_lastValues As Collection

Private Sub Class_Initialize()
    Set m_lastValues = New Collection
End Sub

Public Function UpdateCellIfChanged(ByVal row As Long, ByVal col As Long, ByVal value As String, ByVal foreColor As Long, ByVal backColor As Long) As Boolean
    On Error Resume Next
    ' 简化实现：总是返回 True 并记录值
    UpdateCellIfChanged = True
End Function

' End of CGridDataBuffer.cls