VERSION 5.00
Begin VB.Form frm_WMCopyData 
   Caption         =   "窗口消息"
   ClientHeight    =   12510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14415
   LinkTopic       =   "Form2"
   ScaleHeight     =   12510
   ScaleWidth      =   14415
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   12495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "frm_WMCopyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Debug.Print "frm_WMCopyData.Form_Load: 开始加载，hWnd = " & Me.hWnd
    
    ' 启动监听
    If AddCopyDataSubclass(Me.hWnd) Then
        Debug.Print "CopyData 子类化成功，窗口句柄: " & Me.hWnd
    Else
        Debug.Print "CopyData 子类化失败，窗口句柄: " & Me.hWnd
    End If
    
    ' 添加回调对象（需要实现 OnCopyDataReceived 方法）
    AddCopyDataCallback Me.hWnd, Me
    Debug.Print "CopyData 回调对象已注册，对象类型: " & TypeName(Me)
    
    Debug.Print "frm_WMCopyData.Form_Load: 加载完成"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' 停止监听
    RemoveCopyDataSubclass Me.hWnd
    Debug.Print "CopyData 子类化已移除"
End Sub



