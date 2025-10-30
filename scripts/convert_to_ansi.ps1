# convert_to_ansi.ps1
# 将当前目录（递归）下的 .bas .cls .frm 文件从 UTF-8 转为 系统 ANSI (GBK/CP936)
$enc = [System.Text.Encoding]::GetEncoding(936)
$exts = @("*.bas", "*.cls", "*.frm")

Write-Output "开始转换所有 .bas .cls .frm 文件为 ANSI (CP936/GBK)..."

foreach ($ext in $exts) {
    Get-ChildItem -Recurse -Filter $ext -File | ForEach-Object {
        $path = $_.FullName
        try {
            $text = Get-Content -Raw -Encoding UTF8 -LiteralPath $path
            [System.IO.File]::WriteAllText($path, $text, $enc)
            Write-Output "已转换: $path"
        } catch {
            Write-Output "转换失败: $path - $_"
        }
    }
}

Write-Output "转换完成。请在 VB6 中打开以验证。"

