# 本地一键运行：读取 .env → 调用脚本
# 用法：.\run.ps1            用 .env 里的配置
#      .\run.ps1 --days 30  额外参数会透传给脚本
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
$env:PYTHONIOENCODING = 'utf-8'

if (Test-Path .env) {
    Get-Content .env -Encoding utf8 | ForEach-Object {
        $line = $_.Trim()
        if (-not $line -or $line.StartsWith('#')) { return }
        $idx = $line.IndexOf('=')
        if ($idx -lt 1) { return }
        $key = $line.Substring(0, $idx).Trim()
        $val = $line.Substring($idx + 1).Trim().Trim('"').Trim("'")
        Set-Item -Path "env:$key" -Value $val
    }
}

$py = if (Test-Path '.venv\Scripts\python.exe') { '.venv\Scripts\python.exe' } else { 'python' }
$days = if ($env:DAYS) { $env:DAYS } else { '10' }

& $py 'scripts\baostock_daily_excel.py' '--days' $days @args
exit $LASTEXITCODE
