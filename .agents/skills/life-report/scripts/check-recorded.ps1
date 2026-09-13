# check-recorded.ps1 - 检查某日期是否已有事件记录（默认昨天）
# 输出 RECORDED <日期> 退出码 0 = 已记录
# 输出 MISSING  <日期> 退出码 1 = 缺失（需要发补发提醒）
param([string]$Date)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$Root = Split-Path -Parent $PSScriptRoot
if (-not $Date) { $Date = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd') }
if ($Date -notmatch '^\d{4}-\d{2}-\d{2}$') {
  Write-Output ('ERROR 日期格式应为 yyyy-MM-dd: ' + $Date)
  exit 4
}

$file = Join-Path (Join-Path $Root 'data\events') ($Date + '.md')
if (-not (Test-Path $file)) {
  Write-Output ('MISSING ' + $Date + ' （无记录文件）')
  exit 1
}

$lines = Get-Content $file -Encoding UTF8
$content = @($lines | Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*#' -and $_ -notmatch '^\s*>' })
if ($content.Count -eq 0) {
  Write-Output ('MISSING ' + $Date + ' （文件存在但没有实质内容）')
  exit 1
}

Write-Output ('RECORDED ' + $Date + ' （' + $content.Count + ' 行内容）')
exit 0
