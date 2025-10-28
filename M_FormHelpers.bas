Attribute VB_Name = "M_FormHelpers"
Option Explicit

' 统计标签实例
Private m_lblStats As Label

' 初始化统计标签 - 修改方法签名
Public Sub InitializeStatsLabelOnForm(ByRef parentForm As Form)
    On Error Resume Next
    Set m_lblStats = parentForm.Controls.Add("VB.Label", "lblStats")
    With m_lblStats
        .Caption = "统计: 录制中 0/启用 0/总共 0"
        .Alignment = vbRightJustify
        .BackStyle = 0
        .foreColor = vbBlue
        .Font.name = "微软雅黑"
        .Font.size = 9
        .AutoSize = True
        .Visible = True
    End With
    PositionStatsLabelOnForm parentForm
End Sub

' 定位统计标签 - 修改方法签名
Public Sub PositionStatsLabelOnForm(ByRef parentForm As Form)
    On Error Resume Next
    If Not m_lblStats Is Nothing Then
        With m_lblStats
            .Left = parentForm.ScaleWidth - .Width - 200
            .Top = 10
        End With
    End If
End Sub

' 更新统计标签
Public Sub UpdateStatsLabel()
    On Error Resume Next
    If Not m_lblStats Is Nothing Then
        Dim total As Long, enabled As Long, disabled As Long
        Dim parsing As Long, recording As Long
        GetLiveroomCounts total, enabled, disabled, parsing, recording
        
        m_lblStats.Caption = "统计: 录制中 " & recording & "/启用 " & enabled & "/总共 " & total
        ' 重新定位以确保在窗体调整大小时正确显示
        If Not frmMain Is Nothing Then
            PositionStatsLabelOnForm frmMain
        End If
    End If
End Sub

' 获取统计标签对象（供其他模块使用）
Public Function GetStatsLabel() As Label
    Set GetStatsLabel = m_lblStats
End Function

' 窗体关闭辅助
Public Sub SafeFormUnload(ByRef formToUnload As Form, Cancel As Integer)
    On Error Resume Next
    Cancel = True
    formToUnload.Visible = False
    DoEvents
    SafeShutdownApplication
End Sub

' 窗体调整大小辅助
Public Sub HandleFormResize(ByRef parentForm As Form, ByRef grid As MSHFlexGrid)
    On Error Resume Next
    
    If parentForm.WindowState <> vbMinimized Then
        grid.Width = parentForm.ScaleWidth
        grid.Height = parentForm.ScaleHeight - grid.Top - 200
        
        ' 重新定位统计标签
        PositionStatsLabelOnForm parentForm
        
        ' 调整列宽
        If TypeOf grid Is MSHFlexGrid Then
            Dim gridHelper As New CMSHFlexGridHelper
            gridHelper.AutoSizeColumns grid
        End If
    End If
End Sub
