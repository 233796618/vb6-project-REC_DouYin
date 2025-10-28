Attribute VB_Name = "M_FFmpeg"
Option Explicit

Function ResetFFmpegPipe(web_rid As String, strInfo As String)
    Dim lr As LiveRoom
    lr.currentFrame = ""
    lr.currentFPS = ""
    lr.currentQuality = ""
    lr.currentFileSize = ""
    lr.currentTime = ""
    lr.currentBitrate = ""
    lr.currentSpeed = ""
    lr.currentElapsed = ""
End Function

Function UseFFmpegPipe(web_rid As String, strInfo As String)
    Dim lr As LiveRoom
    
    Set lr = g_liveroomManager.GetLiveroom(web_rid)
    If lr Is Nothing Then Exit Function
    
        
    Dim frame As String, fps As String, quality As String
    Dim size As String, fftime As String, bitrate As String
    Dim speed As String, elapsed As String
    
    frame = ExtractByPattern(strInfo, "frame=\s*(\S+)")
    fps = ExtractByPattern(strInfo, "fps=\s*(\S+)")
    quality = ExtractByPattern(strInfo, "q=\s*(\S+)")
    size = ExtractByPattern(strInfo, "size=\s*(\S+)")
    fftime = ExtractByPattern(strInfo, "time=\s*(\S+)")
    bitrate = ExtractByPattern(strInfo, "bitrate=\s*(\S+)")
    speed = ExtractByPattern(strInfo, "speed=\s*(\S+)")
    elapsed = ExtractByPattern(strInfo, "elapsed=\s*(\S+)")
    ' 特殊处理时间格式
    If fftime <> "" Then
        fftime = FormatFFmpegTime(fftime)
    End If
    
    ' 特殊处理文件大小
    If size <> "" Then
        size = FormatFileSizeDisplay(size)
    End If
    
lr.currentFrame = frame
lr.currentFPS = fps
lr.currentQuality = quality
lr.currentFileSize = size
lr.currentTime = fftime
lr.currentBitrate = bitrate
lr.currentSpeed = speed
lr.currentElapsed = elapsed
'lr.FFmepgExitCode
    
    Debug.Print lr.AnchorName & " " & lr.currentBitrate
End Function

Function PrepareCmdLine(streamUrl As String, AnchorName As String, web_rid As String) As String
'y:\ffmpeg.exe -hwaccel auto -i "$inputurl"-vf crop=1080:990:0:330 -c:v libx264 -crf 32 -preset veryfast -r 15 -g 300 -sc_threshold 500 -bf 4 -refs 6 -c:a copy -y "$output.flv"
Dim Template As String
Template = "y:\ffmpeg.exe -hwaccel auto -i $quote$inputurl$quote -c:v libx264 -crf 32 -preset veryfast -r 15 -g 300 -sc_threshold 500 -bf 4 -refs 6 -c:a copy -y $quote$output.flv$quote"

    Dim Timestamp As String
    Timestamp = Format(Now, "yyyymmddhhnn")
    Dim outputFile As String
    Dim RecFileFolder As String
    RecFileFolder = "e:\抖音直播录像"
    outputFile = RecFileFolder & "\抖音-" & AnchorName & "-" & web_rid & "\" & AnchorName & "_" & Timestamp & ".flv"
    CreateFolderFromPath outputFile
   
   
Template = Replace(Template, "$quote", """")
Template = Replace(Template, "$inputurl", streamUrl)
PrepareCmdLine = Replace(Template, "$output.flv", outputFile)

End Function


Public Function CreateFolderFromPath(ByVal sFilePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sFolderPath As String
    Dim iPos As Integer
    Dim i As Integer
    Dim sPath As String
    
    ' 提取文件夹路径（去掉文件名）
    iPos = InStrRev(sFilePath, "\")
    If iPos > 0 Then
        sFolderPath = Left(sFilePath, iPos - 1)
    Else
        ' 如果没有反斜杠，说明只有文件名，不需要创建文件夹
        CreateFolderFromPath = True
        Exit Function
    End If
    
    ' 如果文件夹已经存在，直接返回成功
    If Dir(sFolderPath, vbDirectory) <> "" Then
        CreateFolderFromPath = True
        Exit Function
    End If
    
    ' 逐层创建文件夹
    sPath = ""
    For i = 1 To Len(sFolderPath)
        Dim sChar As String
        sChar = Mid(sFolderPath, i, 1)
        sPath = sPath & sChar
        
        ' 当遇到反斜杠或者到达路径末尾时，检查并创建文件夹
        If sChar = "\" Or i = Len(sFolderPath) Then
            ' 跳过根目录（如 "C:\"）
            If Right(sPath, 2) = ":\" Then
                ' 根目录已存在，继续下一级
            Else
                ' 检查当前路径是否存在
                If Dir(sPath, vbDirectory) = "" Then
                    MkDir sPath
                End If
            End If
        End If
    Next i
    
    CreateFolderFromPath = True
    Exit Function
    
ErrorHandler:
    CreateFolderFromPath = False
End Function




