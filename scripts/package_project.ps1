# package_project.ps1
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$distDir = Join-Path -Path (Get-Location) -ChildPath "dist"
If (-Not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }
$zipPath = Join-Path -Path $distDir -ChildPath ("project_" + $timestamp + ".zip")

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Add-ToZip($sourceDir, $zipFile) {
    [System.IO.Compression.ZipFile]::CreateFromDirectory($sourceDir, $zipFile)
}

tmp = Join-Path -Path $env:TEMP -ChildPath ("project_pkg_" + $timestamp)
if (Test-Path $tmp) { Remove-Item -Recurse -Force $tmp }
New-Item -ItemType Directory -Path $tmp | Out-Null
Get-ChildItem -Force | Where-Object { $_.Name -ne ".git" -and $_.Name -ne "dist" } | ForEach-Object {
    $destination = Join-Path $tmp $_.Name
    if ($_.PSIsContainer) {
        Copy-Item -Recurse -Force -Path $_.FullName -Destination $destination
    } else {
        Copy-Item -Force -Path $_.FullName -Destination $destination
    }
}

Add-ToZip $tmp $zipPath
Remove-Item -Recurse -Force $tmp

Write-Output "打包完成： $zipPath"

