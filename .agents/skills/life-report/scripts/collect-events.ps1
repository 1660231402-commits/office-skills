# collect-events.ps1 - 汇总日期区间内的事件记录（供报告生成使用）
# 用法: -from 2026-09-01 -to 2026-09-30   或   -date 2026-09-12
# 输出人类可读的结构化文本: PERIOD / DAYS_WITH_RECORDS / STATS / 每天各小节要点
param([string]$From, [string]$To, [string]$Date)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$Root = Split-Path -Parent $PSScriptRoot
if ($Date) {
  if (-not $From) { $From = $Date }
  $To = $Date
}
if (-not $From -or -not $To) {
  Write-Output 'ERROR 用法: collect-events.ps1 -from YYYY-MM-DD -to YYYY-MM-DD  或  -date YYYY-MM-DD'
  exit 4
}

$eventsDir = Join-Path $Root 'data\events'
$days = @()
$sKey = 0; $sGood = 0; $sBad = 0; $sOther = 0

if (Test-Path $eventsDir) {
  $files = @(Get-ChildItem $eventsDir -Filter '*.md' | Where-Object { $_.BaseName -match '^\d{4}-\d{2}-\d{2}$' } | Sort-Object Name)
  foreach ($f in $files) {
    $d = $f.BaseName
    if ($d -lt $From -or $d -gt $To) { continue }
    $sections = [ordered]@{}
    $cur = '补充'
    foreach ($ln in (Get-Content $f.FullName -Encoding UTF8)) {
      if ($ln -match '^##\s*(.+?)\s*$') {
        $cur = $Matches[1]
        if (-not $sections.Contains($cur)) { $sections[$cur] = @() }
      }
      elseif ($ln -match '^\s*[-*]\s+(.+?)\s*$') {
        if (-not $sections.Contains($cur)) { $sections[$cur] = @() }
        $sections[$cur] = @($sections[$cur]) + $Matches[1]
      }
    }
    foreach ($k in @($sections.Keys)) {
      $n = @($sections[$k]).Count
      if ($k -match '重点') { $sKey += $n }
      elseif ($k -match '好') { $sGood += $n }
      elseif ($k -match '坏|问题|挑战') { $sBad += $n }
      else { $sOther += $n }
    }
    $days += $d
    Write-Output ('=== ' + $d + ' ===')
    if ($sections.Count -eq 0) {
      Write-Output '（无结构化内容）'
    }
    foreach ($k in $sections.Keys) {
      Write-Output ('[' + $k + ']')
      foreach ($item in @($sections[$k])) { Write-Output ('- ' + $item) }
    }
  }
}

Write-Output ('PERIOD ' + $From + ' ~ ' + $To)
Write-Output ('DAYS_WITH_RECORDS ' + $days.Count)
Write-Output ('STATS 重点事件=' + $sKey + ' 好事情=' + $sGood + ' 坏事情=' + $sBad + ' 补充=' + $sOther)
exit 0
