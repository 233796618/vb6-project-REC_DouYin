Attribute VB_Name = "M_UpdateControl"
Option Explicit

' 更新控制变量
Private m_updateTimer As Boolean
Private m_lastUpdateTime As Date
Private m_updatePending As Boolean
Private m_forceUpdate As Boolean
Private Const MIN_UPDATE_INTERVAL As Double = 0.5

' 定时更新控制
Public Sub ScheduledUpdate()
    If m_updateTimer Then Exit Sub
    m_updateTimer = True
    
    If (Now - m_lastUpdateTime) * 24 * 60 * 60 < MIN_UPDATE_INTERVAL And Not m_forceUpdate Then
        m_updatePending = True
        m_updateTimer = False
        Exit Sub
    End If
    
    m_lastUpdateTime = Now
    m_updatePending = False
    
    ' 触发主窗体更新
    If Not frmMain Is Nothing Then
        frmMain.DirectUpdateFromManager g_liveroomManager.GetAllLiverooms()
    End If
    
    m_updateTimer = False
End Sub

Public Sub CheckPendingUpdate()
    If m_updatePending And Not m_updateTimer Then
        If (Now - m_lastUpdateTime) * 24 * 60 * 60 >= MIN_UPDATE_INTERVAL Then
            ScheduledUpdate
        End If
    End If
End Sub

Public Sub ForceUpdate()
    m_forceUpdate = True
    ScheduledUpdate
    m_forceUpdate = False
End Sub

' 重置更新状态
Public Sub ResetUpdateControl()
    m_updateTimer = False
    m_updatePending = False
    m_forceUpdate = False
    m_lastUpdateTime = Now
End Sub
