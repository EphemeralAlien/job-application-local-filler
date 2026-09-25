Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"
$toolDirectory = $PSScriptRoot
$vaultPath = Join-Path $toolDirectory "personal.vault"
$fillScript = Join-Path $toolDirectory "local_fill_form.py"

if (-not (Test-Path -LiteralPath $vaultPath)) {
    [System.Windows.Forms.MessageBox]::Show("找不到 personal.vault。请确认它仍在工具文件夹内。", "无法开始") | Out-Null
    exit 1
}

$python = Get-Command python.exe -ErrorAction SilentlyContinue
if (-not $python) {
    [System.Windows.Forms.MessageBox]::Show("找不到 Python。请先安装 Python，或从 PowerShell 运行 python --version 检查。", "无法开始") | Out-Null
    exit 1
}

$picker = New-Object System.Windows.Forms.OpenFileDialog
$picker.Title = "选择 AI 已处理、但仍包含占位符的申请表"
$picker.InitialDirectory = $toolDirectory
$picker.Filter = "Word 文件 (*.docx)|*.docx"
$picker.Multiselect = $false
if ($picker.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    exit 0
}

$inputPath = $picker.FileName
$outputDirectory = Split-Path -LiteralPath $inputPath -Parent
$stem = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
$outputPath = Join-Path $outputDirectory ("{0}-完整信息表.docx" -f $stem)
$number = 2
while (Test-Path -LiteralPath $outputPath) {
    $outputPath = Join-Path $outputDirectory ("{0}-完整信息表-{1}.docx" -f $stem, $number)
    $number++
}

Write-Host "已选择：$([System.IO.Path]::GetFileName($inputPath))"
Write-Host "正在本机替换占位符；接下来输入 Vault 密码时不会显示任何字符。"
& $python.Source $fillScript fill --input $inputPath --output $outputPath --vault $vaultPath
if ($LASTEXITCODE -ne 0) {
    Write-Host "未生成完整申请表。请根据上方提示补齐本地字段后重试。" -ForegroundColor Yellow
    Read-Host "按 Enter 关闭"
    exit $LASTEXITCODE
}

[System.Windows.Forms.MessageBox]::Show("已生成完整申请表：`n$outputPath", "完成") | Out-Null
Start-Process -FilePath $outputPath
