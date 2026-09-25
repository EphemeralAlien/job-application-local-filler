$pidFile = Join-Path $PSScriptRoot "automation\service.pid"
if (-not (Test-Path -LiteralPath $pidFile)) {
    Write-Host "未发现正在运行的本地填表服务。"
    exit 0
}
$servicePid = [int](Get-Content -LiteralPath $pidFile -Raw)
$process = Get-Process -Id $servicePid -ErrorAction SilentlyContinue
if ($process) {
    Stop-Process -Id $servicePid
    Write-Host "本地填表服务已停止。"
} else {
    Write-Host "服务进程已经结束。"
}
Remove-Item -LiteralPath $pidFile -Force -ErrorAction SilentlyContinue
