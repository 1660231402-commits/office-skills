# pick-quote.ps1 - 从名言库随机选一条（可按主题过滤）
# 用法: pick-quote.ps1 [-theme 坚持|成长|挫折|行动|心态|自律|习惯|时间|目标|勇气|感恩|当下|复盘]
# 输出: TEXT / AUTHOR / THEMES 三行
param([string]$Theme)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$Root = Split-Path -Parent $PSScriptRoot
$qp = Join-Path $Root 'assets\quotes.json'
if (-not (Test-Path $qp)) {
  Write-Output 'ERROR assets/quotes.json 不存在'
  exit 4
}

$raw = Get-Content $qp -Raw -Encoding UTF8
$all = @((ConvertFrom-Json -InputObject $raw).quotes)
$pool = $all
if ($Theme) {
  $m = @($all | Where-Object { @($_.themes) -contains $Theme })
  if ($m.Count -gt 0) { $pool = $m }
  else { Write-Output ('NOTE 未找到主题「' + $Theme + '」，已从全部名言中随机选取') }
}

$q = Get-Random -InputObject $pool
Write-Output ('TEXT: ' + $q.text)
Write-Output ('AUTHOR: ' + $q.author)
Write-Output ('THEMES: ' + (@($q.themes) -join ','))
exit 0
