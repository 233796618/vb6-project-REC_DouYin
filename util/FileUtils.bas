' FileUtils.bas - 常用文件工具
Option Explicit

Public Function EnsureFolderExists(ByVal path As String) As Boolean
    On Error Resume Next
    If Dir(path, vbDirectory) = "" Then MkDir path
    EnsureFolderExists = True
End Function

Public Function FormatBytes(ByVal bytes As Currency) As String
    If bytes < 1024 Then
        FormatBytes = CStr(bytes) & " B"
    ElseIf bytes < 1024 * 1024 Then
        FormatBytes = Format$(bytes / 1024, "0.00") & " KB"
    Else
        FormatBytes = Format$(bytes / 1024 / 1024, "0.00") & " MB"
    End If
End Function

' End of FileUtils.bas