param(
    [string]$PythonPath = "",
    [string]$VaultPath = "",
    [string]$OutputDirectory = "",
    [string]$AutomationDirectory = ""
)

$ErrorActionPreference = "Stop"
$toolDirectory = $PSScriptRoot
if (-not $PythonPath) {
    $pythonCommand = Get-Command python.exe -ErrorAction Stop
    $PythonPath = $pythonCommand.Source
}
if (-not $VaultPath) { $VaultPath = Join-Path $toolDirectory "personal.vault" }
if (-not $AutomationDirectory) { $AutomationDirectory = Join-Path $toolDirectory "automation" }
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "网申本地完成表"
}

$inbox = Join-Path $AutomationDirectory "inbox"
$status = Join-Path $AutomationDirectory "status"
$failed = Join-Path $AutomationDirectory "failed"
$pidFile = Join-Path $AutomationDirectory "service.pid"
$serviceScript = Join-Path $toolDirectory "local_fill_service.py"
New-Item -ItemType Directory -Force -Path $inbox, $status, $failed, $OutputDirectory | Out-Null

if (-not (Test-Path -LiteralPath $VaultPath)) {
    throw "找不到 Vault：$VaultPath"
}

$password = Read-Host "输入一次 Vault 密码（不会显示；服务停止后需要重新输入）" -AsSecureString
$credential = New-Object System.Management.Automation.PSCredential("unused", $password)
$plainPassword = $credential.GetNetworkCredential().Password

$arguments = @(
    "`"$serviceScript`"",
    "--inbox", "`"$inbox`"",
    "--outbox", "`"$OutputDirectory`"",
    "--status", "`"$status`"",
    "--failed", "`"$failed`"",
    "--vault", "`"$VaultPath`"",
    "--fill-script", "`"$(Join-Path $toolDirectory 'local_fill_form.py')`"",
    "--password-stdin",
    "--pid-file", "`"$pidFile`""
) -join " "

$startInfo = New-Object System.Diagnostics.ProcessStartInfo
$startInfo.FileName = $PythonPath
$startInfo.Arguments = $arguments
$startInfo.UseShellExecute = $false
$startInfo.CreateNoWindow = $true
$startInfo.RedirectStandardInput = $true
$process = New-Object System.Diagnostics.Process
$process.StartInfo = $startInfo
[void]$process.Start()
$process.StandardInput.WriteLine($plainPassword)
$process.StandardInput.Close()
$plainPassword = $null
$password = $null
$credential = $null
Write-Host "本地填表服务已在后台运行。AI 只能投递占位符 DOCX，完整表格输出在：$OutputDirectory"
