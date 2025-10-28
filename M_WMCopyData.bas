Attribute VB_Name = "M_WMCopyData"
Option Explicit

' API 声明
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long) As Long

Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Any) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

' UTF-8 解码相关 API
Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long

' 常量
Public Const GWL_WNDPROC = -4
Public Const WM_COPYDATA = &H4A
Public Const WM_DESTROY = &H2
Public Const WM_NCDESTROY = &H82
Public Const WM_USER = &H400
Public Const WM_APP = &H8000
Private Const CP_UTF8 As Long = 65001

' 自定义消息
Public Const WM_UPDATEDISPLAY As Long = WM_APP + 100

' 结构定义
Public Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

' CopyData 子类化信息类型
Public Type COPYDATA_SUBCLASS_INFO
    hWnd As Long
    OldProc As Long
    enabled As Boolean
End Type

' 接收到的拷贝数据类型
Public Type COPYDATA_RECEIVED
    FromHwnd As Long
    data As String
    dwData As Long
    Timestamp As Date
    DataLength As Long
    encodingType As String
End Type

' 全局变量 - CopyData 专用
Private g_CopyDataSubclass() As COPYDATA_SUBCLASS_INFO
Private g_CopyDataSubclassCount As Long
Private g_LastCopyData As COPYDATA_RECEIVED
Private g_CopyDataReceived As Boolean
Private g_CopyDataCallback As Collection

' ========== 模块初始化 ==========
Private Sub InitializeCopyDataModule()
    g_CopyDataSubclassCount = 0
    g_CopyDataReceived = False
    
    With g_LastCopyData
        .FromHwnd = 0
        .data = ""
        .dwData = 0
        .Timestamp = Now
        .DataLength = 0
        .encodingType = ""
    End With
    
    Erase g_CopyDataSubclass
    Set g_CopyDataCallback = New Collection
End Sub

' ========== UTF-8 解码函数 ==========
Public Function UTF8ToUnicode(ByRef utf8Bytes() As Byte) As String
    On Error GoTo ErrorHandler
    
    Dim utf8Length As Long
    Dim unicodeLength As Long
    Dim unicodeBuffer() As Byte
    Dim result As Long
    
    utf8Length = UBound(utf8Bytes) - LBound(utf8Bytes) + 1
    If utf8Length <= 0 Then
        UTF8ToUnicode = ""
        Exit Function
    End If
    
    unicodeLength = MultiByteToWideChar(CP_UTF8, 0, _
        VarPtr(utf8Bytes(LBound(utf8Bytes))), utf8Length, 0, 0)
    
    If unicodeLength <= 0 Then
        UTF8ToUnicode = ""
        Exit Function
    End If
    
    ReDim unicodeBuffer(0 To (unicodeLength * 2) - 1) As Byte
    
    result = MultiByteToWideChar(CP_UTF8, 0, _
        VarPtr(utf8Bytes(LBound(utf8Bytes))), utf8Length, _
        VarPtr(unicodeBuffer(0)), unicodeLength)
    
    If result > 0 Then
        UTF8ToUnicode = unicodeBuffer
    Else
        UTF8ToUnicode = ""
    End If
    
    Exit Function
    
ErrorHandler:
    UTF8ToUnicode = ""
End Function

Public Function IsUTF8(ByRef dataBytes() As Byte) As Boolean
    On Error Resume Next
    
    Dim i As Long
    Dim currentByte As Long
    Dim byteCount As Long
    
    byteCount = UBound(dataBytes) - LBound(dataBytes) + 1
    If byteCount < 1 Then
        IsUTF8 = False
        Exit Function
    End If
    
    ' 检查 UTF-8 BOM
    If byteCount >= 3 Then
        If dataBytes(LBound(dataBytes)) = &HEF And _
           dataBytes(LBound(dataBytes) + 1) = &HBB And _
           dataBytes(LBound(dataBytes) + 2) = &HBF Then
            IsUTF8 = True
            Exit Function
        End If
    End If
    
    ' 简单的 UTF-8 模式检查
    i = LBound(dataBytes)
    Do While i <= UBound(dataBytes)
        currentByte = dataBytes(i)
        
        If (currentByte And &H80) = 0 Then
            ' ASCII 字符
            i = i + 1
        ElseIf (currentByte And &HE0) = &HC0 Then
            ' 2字节 UTF-8 字符
            If i + 1 > UBound(dataBytes) Then Exit Do
            If (dataBytes(i + 1) And &HC0) <> &H80 Then
                IsUTF8 = False
                Exit Function
            End If
            i = i + 2
        ElseIf (currentByte And &HF0) = &HE0 Then
            ' 3字节 UTF-8 字符
            If i + 2 > UBound(dataBytes) Then Exit Do
            If (dataBytes(i + 1) And &HC0) <> &H80 Or _
               (dataBytes(i + 2) And &HC0) <> &H80 Then
                IsUTF8 = False
                Exit Function
            End If
            i = i + 3
        ElseIf (currentByte And &HF8) = &HF0 Then
            ' 4字节 UTF-8 字符
            If i + 3 > UBound(dataBytes) Then Exit Do
            If (dataBytes(i + 1) And &HC0) <> &H80 Or _
               (dataBytes(i + 2) And &HC0) <> &H80 Or _
               (dataBytes(i + 3) And &HC0) <> &H80 Then
                IsUTF8 = False
                Exit Function
            End If
            i = i + 4
        Else
            IsUTF8 = False
            Exit Function
        End If
    Loop
    
    IsUTF8 = True
End Function

Public Function SmartDecodeText(ByRef dataBytes() As Byte) As String
    On Error GoTo ErrorHandler
    
    Dim DataLength As Long
    DataLength = UBound(dataBytes) - LBound(dataBytes) + 1
    
    If DataLength <= 0 Then
        SmartDecodeText = ""
        Exit Function
    End If
    
    ' 优先尝试 UTF-8 解码
    If IsUTF8(dataBytes) Then
        SmartDecodeText = UTF8ToUnicode(dataBytes)
        If Len(SmartDecodeText) > 0 Then
            Exit Function
        End If
    End If
    
    ' 如果 UTF-8 解码失败，尝试 ANSI 解码
    On Error Resume Next
    SmartDecodeText = StrConv(dataBytes, vbUnicode)
    
    ' 去除可能的 null 字符
    If InStr(SmartDecodeText, vbNullChar) > 0 Then
        SmartDecodeText = Left$(SmartDecodeText, InStr(SmartDecodeText, vbNullChar) - 1)
    End If
    
    Exit Function
    
ErrorHandler:
    SmartDecodeText = "[解码错误: " & Err.Description & "]"
End Function

' ========== CopyData 子类化核心函数 ==========
Public Function CopyDataSubclassProc( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim OldProc As Long
    Dim Handled As Boolean
    Dim ReturnValue As Long
    
    ' 查找对应的子类化信息
    For i = 0 To g_CopyDataSubclassCount - 1
        If g_CopyDataSubclass(i).hWnd = hWnd Then
            OldProc = g_CopyDataSubclass(i).OldProc
            
            ' 根据消息类型处理
            Select Case uMsg
                Case WM_COPYDATA
                    ReturnValue = HandleCopyDataMessage(wParam, lParam)
                    Handled = True
                    
                Case WM_DESTROY, WM_NCDESTROY
                    ' 处理销毁消息
                    Debug.Print "CopyDataSubclassProc: 收到销毁消息，窗口句柄: " & hWnd
                    RemoveCopyDataSubclass hWnd
                    
                Case WM_UPDATEDISPLAY
                    ' 自定义消息 - 通知窗体更新显示
                    UpdateMessageFormImmediate
                    Handled = True
                    ReturnValue = 1
            End Select
            
            ' 如果消息已处理，返回处理结果
            If Handled Then
                CopyDataSubclassProc = ReturnValue
                Exit Function
            End If
            
            ' 调用原窗口过程
            If OldProc <> 0 Then
                CopyDataSubclassProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
            Else
                CopyDataSubclassProc = 0
            End If
            
            Exit Function
        End If
    Next i
    
    ' 如果没有找到子类化信息，调用默认处理
    CopyDataSubclassProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
    Exit Function
    
ErrorHandler:
    Debug.Print "CopyDataSubclassProc 错误: " & Err.Description & ", 消息: " & uMsg
    If OldProc <> 0 Then
        CopyDataSubclassProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
    Else
        CopyDataSubclassProc = 0
    End If
End Function

' 默认窗口过程
Private Function DefWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    DefWindowProc = 0
End Function

Private Function HandleCopyDataMessage( _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
    On Error GoTo ErrorHandler
    
    Dim cds As COPYDATASTRUCT
    Dim byteData() As Byte
    Dim strData As String
    Dim DataLength As Long
    Dim encodingType As String
    
    ' 拷贝数据结构
    CopyMemory cds, ByVal lParam, Len(cds)
    
    DataLength = cds.cbData
    If DataLength > 0 And DataLength < 65536 Then
        
        ' 复制数据到字节数组
        ReDim byteData(0 To DataLength - 1) As Byte
        CopyMemory byteData(0), ByVal cds.lpData, DataLength
        
        ' 使用智能解码
        strData = SmartDecodeText(byteData)
        
        Debug.Print "HandleCopyDataMessage: 解码后的数据: " & strData
        
        ' 确定编码类型
        If IsUTF8(byteData) Then
            encodingType = "UTF-8"
        Else
            encodingType = "ANSI"
        End If
        
        ' 如果解码后为空，显示原始信息
        If Len(strData) = 0 Then
            strData = "[二进制数据，长度: " & DataLength & " 字节]"
        End If
        
        ' 存储接收到的数据
        With g_LastCopyData
            .FromHwnd = wParam
            .data = strData
            .dwData = cds.dwData
            .Timestamp = Now
            .DataLength = DataLength
            .encodingType = encodingType
        End With
        
        g_CopyDataReceived = True
        
        ' 直接处理数据（确保数据被处理）
        UseJsonMessage strData
        
        HandleCopyDataMessage = 1  ' 表示消息已处理
    Else
        HandleCopyDataMessage = 0  ' 表示消息未处理
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "HandleCopyDataMessage 错误: " & Err.Description
    HandleCopyDataMessage = 0
End Function

' ========== CopyData 子类化管理函数 ==========
Public Function AddCopyDataSubclass(ByVal hWnd As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' 初始化模块（如果尚未初始化）
    If g_CopyDataSubclassCount = 0 Then
        InitializeCopyDataModule
    End If
    
    ' 检查窗口是否有效
    If IsWindow(hWnd) = 0 Then
        Debug.Print "AddCopyDataSubclass: 无效的窗口句柄: " & hWnd
        AddCopyDataSubclass = False
        Exit Function
    End If
    
    ' 检查是否已子类化
    If FindCopyDataSubclassIndex(hWnd) >= 0 Then
        Debug.Print "AddCopyDataSubclass: 窗口已子类化: " & hWnd
        AddCopyDataSubclass = True
        Exit Function
    End If
    
    ' 创建新的子类化
    If g_CopyDataSubclassCount = 0 Then
        ReDim g_CopyDataSubclass(0) As COPYDATA_SUBCLASS_INFO
    Else
        ReDim Preserve g_CopyDataSubclass(g_CopyDataSubclassCount) As COPYDATA_SUBCLASS_INFO
    End If
    
    With g_CopyDataSubclass(g_CopyDataSubclassCount)
        .hWnd = hWnd
        .enabled = True
        
        ' 保存原窗口过程并设置新的窗口过程
        .OldProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf CopyDataSubclassProc)
        
        If .OldProc = 0 Then
            Debug.Print "AddCopyDataSubclass: SetWindowLong 失败，窗口句柄: " & hWnd
            AddCopyDataSubclass = False
            Exit Function
        End If
    End With
    
    Debug.Print "AddCopyDataSubclass: 创建 CopyData 子类化成功，窗口句柄: " & hWnd & ", 原过程: " & g_CopyDataSubclass(g_CopyDataSubclassCount).OldProc
    
    g_CopyDataSubclassCount = g_CopyDataSubclassCount + 1
    AddCopyDataSubclass = True
    
    Exit Function
    
ErrorHandler:
    Debug.Print "AddCopyDataSubclass 错误: " & Err.Description
    AddCopyDataSubclass = False
End Function

Public Sub RemoveCopyDataSubclass(ByVal hWnd As Long)
    On Error Resume Next
    
    Dim Index As Long
    Index = FindCopyDataSubclassIndex(hWnd)
    
    If Index >= 0 Then
        Debug.Print "RemoveCopyDataSubclass: 移除 CopyData 子类化，窗口句柄: " & hWnd
        
        ' 恢复原窗口过程
        If g_CopyDataSubclass(Index).OldProc <> 0 Then
            SetWindowLong hWnd, GWL_WNDPROC, g_CopyDataSubclass(Index).OldProc
        End If
        
        ' 从数组中移除
        If g_CopyDataSubclassCount > 1 Then
            Dim i As Long
            For i = Index To g_CopyDataSubclassCount - 2
                g_CopyDataSubclass(i) = g_CopyDataSubclass(i + 1)
            Next i
        End If
        
        g_CopyDataSubclassCount = g_CopyDataSubclassCount - 1
        
        If g_CopyDataSubclassCount > 0 Then
            ReDim Preserve g_CopyDataSubclass(g_CopyDataSubclassCount - 1) As COPYDATA_SUBCLASS_INFO
        Else
            Erase g_CopyDataSubclass
            InitializeCopyDataModule
        End If
        
        ' 移除回调
        RemoveCopyDataCallback hWnd
    End If
End Sub

Public Function FindCopyDataSubclassIndex(ByVal hWnd As Long) As Long
    Dim i As Long
    For i = 0 To g_CopyDataSubclassCount - 1
        If g_CopyDataSubclass(i).hWnd = hWnd Then
            FindCopyDataSubclassIndex = i
            Exit Function
        End If
    Next i
    FindCopyDataSubclassIndex = -1
End Function

Public Function IsCopyDataSubclassed(ByVal hWnd As Long) As Boolean
    IsCopyDataSubclassed = (FindCopyDataSubclassIndex(hWnd) >= 0)
End Function

Public Sub CleanupAllCopyDataSubclass()
    On Error Resume Next
    Dim i As Long
    
    Debug.Print "CleanupAllCopyDataSubclass: 开始清理所有 CopyData 子类化"
    
    For i = g_CopyDataSubclassCount - 1 To 0 Step -1
        If g_CopyDataSubclass(i).OldProc <> 0 And g_CopyDataSubclass(i).hWnd <> 0 Then
            If IsWindow(g_CopyDataSubclass(i).hWnd) Then
                SetWindowLong g_CopyDataSubclass(i).hWnd, GWL_WNDPROC, g_CopyDataSubclass(i).OldProc
                Debug.Print "CleanupAllCopyDataSubclass: 恢复窗口过程，句柄: " & g_CopyDataSubclass(i).hWnd
            End If
        End If
    Next i
    
    g_CopyDataSubclassCount = 0
    Erase g_CopyDataSubclass
    Set g_CopyDataCallback = Nothing
    InitializeCopyDataModule
    
    Debug.Print "CleanupAllCopyDataSubclass: 所有 CopyData 子类化已清理完成"
End Sub

' ========== 回调管理函数 ==========
Public Sub AddCopyDataCallback(ByVal hWnd As Long, ByRef CallbackObject As Object)
    On Error Resume Next
    g_CopyDataCallback.Add CallbackObject, CStr(hWnd)
    Debug.Print "AddCopyDataCallback: 添加回调对象成功，hWnd: " & hWnd & ", 对象类型: " & TypeName(CallbackObject)
End Sub

Public Sub RemoveCopyDataCallback(ByVal hWnd As Long)
    On Error Resume Next
    g_CopyDataCallback.Remove CStr(hWnd)
    Debug.Print "RemoveCopyDataCallback: 移除回调对象，hWnd: " & hWnd
End Sub

Private Sub CallCopyDataCallbacks(ByVal FromHwnd As Long, ByVal data As String, ByVal dwData As Long, ByVal DataLength As Long, ByVal encodingType As String)
    On Error Resume Next
    Dim CallbackObj As Object
    Dim key As Variant
    
    Debug.Print "CallCopyDataCallbacks: 开始调用回调，回调对象数量: " & g_CopyDataCallback.Count
    
    For Each key In g_CopyDataCallback
        Set CallbackObj = g_CopyDataCallback(key)
        If Not CallbackObj Is Nothing Then
            Debug.Print "CallCopyDataCallbacks: 调用回调对象，键: " & key & ", 类型: " & TypeName(CallbackObj)
            
            ' 直接调用 OnCopyDataReceived 方法
            CallbackObj.OnCopyDataReceived FromHwnd, data, dwData, DataLength, encodingType
            Debug.Print "CallCopyDataCallbacks: 回调调用完成"
        Else
            Debug.Print "CallCopyDataCallbacks: 回调对象为空，键: " & key
        End If
    Next key
    
    Debug.Print "CallCopyDataCallbacks: 所有回调调用完成"
End Sub

Private Function GetCallbackHwnd() As Long
    On Error Resume Next
    If g_CopyDataCallback.Count > 0 Then
        GetCallbackHwnd = CLng(g_CopyDataCallback(1).hWnd)
    Else
        GetCallbackHwnd = 0
    End If
End Function

' 立即更新消息窗体显示
Public Sub UpdateMessageFormImmediate()
    On Error Resume Next
    ' 这里需要根据具体窗体调整
    ' If Not Form2 Is Nothing Then
    '     Form2.UpdateDisplayImmediate
    ' End If
End Sub