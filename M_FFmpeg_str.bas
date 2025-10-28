Attribute VB_Name = "M_FFmpeg_str"

' 精确解析FFmpeg进度信息
Public Sub ParseFFmpegOutputExact(ByVal output As String)
    On Error Resume Next


End Sub

' 使用模式匹配提取值
Public Function ExtractByPattern(ByVal text As String, ByVal pattern As String) As String
    On Error Resume Next
    
    Dim key As String
    Dim keyPos As Long
    Dim valueStart As Long
    Dim valueEnd As Long
    Dim inValue As Boolean
    Dim result As String
    
    ' 提取关键字
    key = Left(pattern, InStr(pattern, "="))
    keyPos = InStr(text, key)
    If keyPos = 0 Then
        ExtractByPattern = ""
        Exit Function
    End If
    
    valueStart = keyPos + Len(key)
    inValue = False
    result = ""
    
    ' 遍历字符提取值
    Do While valueStart <= Len(text)
        Dim ch As String
        ch = Mid(text, valueStart, 1)
        
        If ch <> " " And ch <> vbTab Then
            ' 开始提取值
            If Not inValue Then
                inValue = True
            End If
            result = result & ch
        ElseIf inValue Then
            ' 遇到空格且已经在提取值中，结束提取
            Exit Do
        End If
        
        valueStart = valueStart + 1
    Loop
    
    ExtractByPattern = Trim(result)
End Function

' 格式化FFmpeg时间显示
Public Function FormatFFmpegTime(ByVal timeStr As String) As String
    On Error Resume Next
    
    ' FFmpeg时间格式: HH:MM:SS.mmm 或 H:MM:SS.mmm
    Dim parts() As String
    Dim timePart As String
    Dim msPart As String
    
    ' 分割毫秒部分
    parts = Split(timeStr, ".")
    If UBound(parts) >= 1 Then
        timePart = parts(0)
        msPart = Left(parts(1), 3) ' 只取前3位毫秒
        FormatFFmpegTime = timePart & "." & msPart
    Else
        FormatFFmpegTime = timeStr
    End If
End Function

' 格式化文件大小显示
Public Function FormatFileSizeDisplay(ByVal sizeStr As String) As String
    On Error Resume Next
    
    Dim sizeValue As Double
    Dim unit As String
    
    ' 提取数值和单位
    If InStr(sizeStr, "kB") > 0 Then
        sizeValue = Val(Replace(sizeStr, "kB", ""))
        unit = "KB"
    ElseIf InStr(sizeStr, "MB") > 0 Then
        sizeValue = Val(Replace(sizeStr, "MB", ""))
        unit = "MB"
    ElseIf InStr(sizeStr, "GB") > 0 Then
        sizeValue = Val(Replace(sizeStr, "GB", ""))
        unit = "GB"
    Else
        ' 假设是字节
        sizeValue = Val(sizeStr)
        If sizeValue > 0 Then
            If sizeValue < 1024 Then
                unit = "B"
            ElseIf sizeValue < 1024 * 1024 Then
                sizeValue = sizeValue / 1024
                unit = "KB"
            Else
                sizeValue = sizeValue / 1024 / 1024
                unit = "MB"
            End If
        Else
            unit = "B"
        End If
    End If
    
    FormatFileSizeDisplay = Format(sizeValue, "0.0") & " " & unit
End Function
