# send-mail.ps1 - life-report skill 邮件发送脚本（Windows PowerShell 5.1，零依赖）
# 用法:
#   -mode test                                    发送测试邮件
#   -mode reminder1 | reminder2                   每日记录提醒（当天幂等）
#   -mode missed [-date yyyy-MM-dd]               补发提醒（默认检查昨天，按日期幂等）
#   -mode report -file <html路径> -subject <主题> -tag <标识>    发送报告（HTML正文+附件，按tag幂等）
#   -dryrun                                       只构建邮件不发送（写入 data\last_dryrun.eml 供检查）
param(
  [Parameter(Mandatory = $true)]
  [ValidateSet('test', 'reminder1', 'reminder2', 'missed', 'report')]
  [string]$Mode,
  [string]$Date,
  [string]$File,
  [string]$Subject,
  [string]$Tag,
  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$Root = Split-Path -Parent $PSScriptRoot
$ConfigPath = Join-Path $Root 'config.json'
$StatePath = Join-Path $Root 'data\state.json'

function Fail([string]$Msg, [int]$Code = 3) {
  Write-Host ("[life-report] 发送未完成: " + $Msg)
  exit $Code
}

if (-not $Date) {
  if ($Mode -eq 'missed') { $Date = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd') }
  else { $Date = (Get-Date).ToString('yyyy-MM-dd') }
}

# ---------- 配置 ----------
if (-not (Test-Path $ConfigPath)) {
  Fail ("找不到配置文件 " + $ConfigPath + " ，请先按 README.md 配置邮箱。", 2)
}
$cfg = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$smtp = $cfg.smtp

if (-not $DryRun) {
  $missing = @()
  if (-not $smtp.username -or $smtp.username -match 'your_email') { $missing += 'smtp.username' }
  if (-not $smtp.password -or $smtp.password -match '在这里|授权码') { $missing += 'smtp.password' }
  if (-not $cfg.to -or (@($cfg.to)[0]) -match 'your_email') { $missing += 'to' }
  if ($missing.Count -gt 0) {
    Write-Host ("[life-report] 邮箱尚未配置，缺少: " + ($missing -join ', '))
    Write-Host ("请编辑 " + $ConfigPath)
    Write-Host "填入邮箱地址与 SMTP 授权码（授权码不是登录密码，获取方法见 README.md）。"
    Write-Host "QQ邮箱: 网页版 -> 设置 -> 账户 -> 开启 POP3/IMAP/SMTP 服务 -> 生成授权码"
    exit 2
  }
}

# ---------- 幂等 ----------
function Test-TagSent([string]$t) {
  if (-not (Test-Path $StatePath)) { return $false }
  try {
    $st = Get-Content $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($st.sent -and $st.sent.PSObject.Properties[$t]) { return $true }
  } catch {}
  return $false
}

function Set-TagSent([string]$t) {
  $obj = $null
  if (Test-Path $StatePath) {
    try { $obj = Get-Content $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
  }
  if (-not $obj) { $obj = New-Object System.Management.Automation.PSObject }
  $now = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
  if (-not ($obj.PSObject.Properties['sent'])) {
    $obj | Add-Member -NotePropertyName 'sent' -NotePropertyValue (New-Object System.Management.Automation.PSObject) -Force
  }
  if ($obj.sent.PSObject.Properties[$t]) { $obj.sent.$t = $now }
  else { $obj.sent | Add-Member -NotePropertyName $t -NotePropertyValue $now }
  New-Item -ItemType Directory -Force -Path (Split-Path -Parent $StatePath) | Out-Null
  $obj | ConvertTo-Json -Depth 8 | Set-Content -Path $StatePath -Encoding UTF8
}

if (-not $Tag) {
  switch ($Mode) {
    'reminder1' { $Tag = $Date + '_提醒2000' }
    'reminder2' { $Tag = $Date + '_提醒2200' }
    'missed'    { $Tag = $Date + '_补发提醒' }
  }
}
if ($Tag -and -not $DryRun -and (Test-TagSent $Tag)) {
  Write-Host ("ALREADY_SENT " + $Tag + " （此前已发送过，本次跳过）")
  exit 0
}

# ---------- 邮件内容 ----------
function Get-RandomQuote {
  $qp = Join-Path $Root 'assets\quotes.json'
  if (-not (Test-Path $qp)) { return $null }
  try {
    $all = @((ConvertFrom-Json -InputObject (Get-Content $qp -Raw -Encoding UTF8)).quotes)
    if ($all.Count -gt 0) { return (Get-Random -InputObject $all) }
  } catch {}
  return $null
}

$attachPath = $null
$textBody = ''
$htmlBody = ''

if ($Mode -eq 'report') {
  if (-not $File) { Fail 'report 模式需要 -file 指向报告 HTML 文件。' 4 }
  if (-not (Test-Path $File)) { Fail ("报告文件不存在: " + $File) 4 }
  if (-not $Tag) { Fail 'report 模式需要 -tag 用于防重复发送，例如 2026-09-12_日报。' 4 }
  $attachPath = (Resolve-Path $File).Path
  $htmlBody = Get-Content $attachPath -Raw -Encoding UTF8
  if (-not $Subject) { $Subject = [System.IO.Path]::GetFileNameWithoutExtension($attachPath) }
  $textBody = "【生活复盘报告】" + $Subject + "`r`n`r`n报告已生成为 HTML，请直接查看邮件正文，或下载附件用浏览器打开。"
}
else {
  $quote = Get-RandomQuote
  $qLine = ''
  $qHtml = ''
  if ($quote) {
    $qLine = '「' + $quote.text + '」 —— ' + $quote.author
    $qHtml = '<div style="margin-top:18px;padding-top:14px;border-top:1px solid #e4e4e7;font-size:13px;color:#71717a;"><i>' + $quote.text + '</i> —— ' + $quote.author + '</div>'
  }
  $head = ''
  $body = ''
  switch ($Mode) {
    'test' {
      $head = '邮箱配置成功 ✅'
      $body = '这是一封测试邮件。收到它说明 life-report 的邮箱配置正确。<br><br>之后每晚 20:00 / 22:00 的记录提醒，以及日报 / 周报 / 月报 / 半年报 / 年报，都会发送到这个邮箱。'
    }
    'reminder1' {
      $head = '晚上好，今天过得怎么样？'
      $body = '花三分钟，把今天的事情讲给我听吧。<br><br>在 ZCode 中对我说：<br><b>「记录事件：今天……」</b><br><br>我会帮你整理成 重点事件 / 好事情 / 坏事情，明早 8 点你会收到今天的日报。'
    }
    'reminder2' {
      $head = '今天的记录还没写'
      $body = '睡前两分钟，别让今天的故事溜走。<br><br>对我说：<br><b>「记录事件：今天……」</b><br><br>现在记下，明早 8 点就能收到今天的日报。'
    }
    'missed' {
      $head = ($Date + ' 的记录还空着')
      $body = '检测到 ' + $Date + ' 没有事件记录。<br><br>现在补发还来得及，对我说：<br><b>「补发 ' + $Date + ' 的事件：……」</b><br><br>补发的内容会并入之后的周报 / 月报分析。'
    }
  }
  $plainBody = ($body -replace '<br>', "`r`n") -replace '<[^>]+>', ''
  $textBody = "【生活复盘提醒】" + $head + "`r`n`r`n" + $plainBody + "`r`n`r`n" + $qLine
  $tpl = @'
<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:28px 12px;background:#f2f2f0;font-family:'Helvetica Neue','PingFang SC','Microsoft YaHei',sans-serif;color:#18181b;">
  <div style="max-width:560px;margin:0 auto;background:#ffffff;border:1px solid #e0e0dd;">
    <div style="background:#111111;padding:26px 30px 24px;color:#ffffff;">
      <div style="font-size:10px;letter-spacing:5px;color:#9ca3af;">LIFE REPORT</div>
      <div style="font-size:19px;font-weight:600;letter-spacing:2px;margin-top:10px;">@@HEAD@@</div>
    </div>
    <div style="padding:26px 30px;font-size:14px;line-height:2.1;color:#3f3f46;">@@BODY@@
    @@QUOTE@@</div>
    <div style="padding:14px 30px 18px;border-top:1px solid #eeeeee;background:#fafafa;font-size:11px;color:#a1a1aa;">由 ZCode · life-report 自动发送</div>
  </div>
</body></html>
'@
  $htmlBody = $tpl.Replace('@@HEAD@@', $head).Replace('@@BODY@@', $body).Replace('@@QUOTE@@', $qHtml)
}

if (-not $Subject) {
  switch ($Mode) {
    'test'      { $Subject = 'life-report 测试邮件 · ' + $Date }
    'reminder1' { $Subject = '晚间提醒（第一遍）| 记下今天的 3 分钟 · ' + $Date }
    'reminder2' { $Subject = '最后提醒 | 今天的记录还没写 · ' + $Date }
    'missed'    { $Subject = '补发提醒 | ' + $Date + ' 的记录还空着' }
  }
}

# ---------- 组装 MIME ----------
function B64s([string]$s) { [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($s)) }
function EncodeHdr([string]$s) { '=?utf-8?B?' + (B64s $s) + '?=' }
function UrlEnc([string]$s) {
  -join ([System.Text.Encoding]::UTF8.GetBytes($s) | ForEach-Object { '%{0:X2}' -f $_ })
}
function ChunkB64([string]$b64) {
  $sb2 = New-Object System.Text.StringBuilder
  for ($i = 0; $i -lt $b64.Length; $i += 76) {
    $n = [Math]::Min(76, $b64.Length - $i)
    [void]$sb2.Append($b64.Substring($i, $n))
    if ($i + 76 -lt $b64.Length) { [void]$sb2.Append("`r`n") }
  }
  $sb2.ToString()
}

$fromAddr = $smtp.username
$fromDisp = if ($smtp.sender_name) { $smtp.sender_name } else { 'life-report' }
$toList = @($cfg.to)

$bm = '----=_lr_mixed_' + [Guid]::NewGuid().ToString('N')
$ba = '----=_lr_alt_' + [Guid]::NewGuid().ToString('N')

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append('From: ' + (EncodeHdr $fromDisp) + ' <' + $fromAddr + '>' + "`r`n")
[void]$sb.Append('To: ' + ($toList -join ', ') + "`r`n")
[void]$sb.Append('Subject: ' + (EncodeHdr $Subject) + "`r`n")
[void]$sb.Append('Date: ' + [System.DateTime]::UtcNow.ToString('R') + "`r`n")
[void]$sb.Append('MIME-Version: 1.0' + "`r`n")
[void]$sb.Append('X-Mailer: life-report-skill' + "`r`n")
[void]$sb.Append('Content-Type: multipart/mixed; boundary="' + $bm + '"' + "`r`n`r`n")
[void]$sb.Append('--' + $bm + "`r`n")
[void]$sb.Append('Content-Type: multipart/alternative; boundary="' + $ba + '"' + "`r`n`r`n")
[void]$sb.Append('--' + $ba + "`r`n")
[void]$sb.Append('Content-Type: text/plain; charset="utf-8"' + "`r`n")
[void]$sb.Append('Content-Transfer-Encoding: base64' + "`r`n`r`n")
[void]$sb.Append((ChunkB64 (B64s $textBody)) + "`r`n")
[void]$sb.Append('--' + $ba + "`r`n")
[void]$sb.Append('Content-Type: text/html; charset="utf-8"' + "`r`n")
[void]$sb.Append('Content-Transfer-Encoding: base64' + "`r`n`r`n")
[void]$sb.Append((ChunkB64 (B64s $htmlBody)) + "`r`n")
[void]$sb.Append('--' + $ba + '--' + "`r`n")

if ($attachPath) {
  $asciiName = ($Tag -replace '_日报$', '_Daily' -replace '_周报$', '_Weekly' -replace '_月报$', '_Monthly' -replace '_半年报$', '_HalfYear' -replace '_年报$', '_Annual') + '.html'
  $attB64 = ChunkB64 ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($attachPath)))
  [void]$sb.Append('--' + $bm + "`r`n")
  [void]$sb.Append('Content-Type: text/html; name="' + $asciiName + '"' + "`r`n")
  [void]$sb.Append('Content-Disposition: attachment; filename="' + $asciiName + '"; filename*=utf-8' + "''" + (UrlEnc ([System.IO.Path]::GetFileName($attachPath))) + "`r`n")
  [void]$sb.Append('Content-Transfer-Encoding: base64' + "`r`n`r`n")
  [void]$sb.Append($attB64 + "`r`n")
}
[void]$sb.Append('--' + $bm + '--' + "`r`n")

$msgText = $sb.ToString()
if (-not $msgText.EndsWith("`r`n")) { $msgText += "`r`n" }
$msgText = $msgText -replace '(?m)^\.', '..'
$payload = [System.Text.Encoding]::UTF8.GetBytes($msgText)

# ---------- 发送 ----------
function Send-SmtpRaw {
  param(
    [string]$SmtpHost, [int]$Port, [bool]$ImplicitSsl, [bool]$UseStartTls,
    [string]$User, [string]$Pass, [string]$From, [string[]]$To, [byte[]]$Payload
  )
  $tcp = New-Object System.Net.Sockets.TcpClient
  $tcp.SendTimeout = 30000
  $tcp.ReceiveTimeout = 30000
  $tcp.Connect($SmtpHost, $Port)
  try {
    $stream = $tcp.GetStream()
    if ($ImplicitSsl) {
      $ssl = New-Object System.Net.Security.SslStream($stream)
      $ssl.AuthenticateAsClient($SmtpHost)
      $stream = $ssl
    }
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
    $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
    $writer.AutoFlush = $true
    $writer.NewLine = "`r`n"

    function Read-Resp {
      $lines = @()
      while ($true) {
        $line = $reader.ReadLine()
        if ($null -eq $line) { throw '连接被服务器关闭' }
        $lines += $line
        if ($line.Length -lt 4) { break }
        if ($line[3] -eq ' ') { break }
      }
      return ($lines -join ' | ')
    }
    function Confirm-Resp([string]$Prefix, [string]$Resp, [string]$Step) {
      if (-not $Resp.StartsWith($Prefix)) { throw ($Step + ' 失败，服务器返回: ' + $Resp) }
    }

    $resp = Read-Resp
    Confirm-Resp '220' $resp 'SMTP连接'
    $writer.WriteLine('EHLO [127.0.0.1]')
    $resp = Read-Resp
    if (-not $resp.StartsWith('250')) {
      $writer.WriteLine('HELO [127.0.0.1]')
      $resp = Read-Resp
      Confirm-Resp '250' $resp 'HELO'
    }

    if ($UseStartTls) {
      $writer.WriteLine('STARTTLS')
      $resp = Read-Resp
      Confirm-Resp '220' $resp 'STARTTLS'
      $ssl = New-Object System.Net.Security.SslStream($stream)
      $ssl.AuthenticateAsClient($SmtpHost)
      $stream = $ssl
      $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
      $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
      $writer.AutoFlush = $true
      $writer.NewLine = "`r`n"
      $writer.WriteLine('EHLO [127.0.0.1]')
      $resp = Read-Resp
      Confirm-Resp '250' $resp 'EHLO(TLS)'
    }

    $writer.WriteLine('AUTH LOGIN')
    $resp = Read-Resp
    Confirm-Resp '334' $resp 'AUTH'
    $writer.WriteLine([Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($User)))
    $resp = Read-Resp
    Confirm-Resp '334' $resp '用户名'
    $writer.WriteLine([Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Pass)))
    $resp = Read-Resp
    Confirm-Resp '235' $resp '登录（请检查授权码）'

    $writer.WriteLine('MAIL FROM:<' + $From + '>')
    $resp = Read-Resp
    Confirm-Resp '250' $resp 'MAIL FROM'
    foreach ($r in $To) {
      $writer.WriteLine('RCPT TO:<' + $r + '>')
      $resp = Read-Resp
      if (-not ($resp.StartsWith('250') -or $resp.StartsWith('251'))) { throw ('收件人被拒绝 ' + $r + ' : ' + $resp) }
    }
    $writer.WriteLine('DATA')
    $resp = Read-Resp
    Confirm-Resp '354' $resp 'DATA'
    $stream.Write($Payload, 0, $Payload.Length)
    $stream.Flush()
    $writer.WriteLine('.')
    $resp = Read-Resp
    Confirm-Resp '250' $resp '投递'
    $writer.WriteLine('QUIT')
    try { Read-Resp | Out-Null } catch {}
    return $true
  }
  finally {
    $tcp.Close()
  }
}

if ($DryRun) {
  $eml = Join-Path $Root 'data\last_dryrun.eml'
  [System.IO.File]::WriteAllBytes($eml, $payload)
  Write-Host ('DRYRUN OK - 邮件已构建（未发送），已写入 ' + $eml)
  Write-Host ('收件人: ' + ($toList -join ', ') + ' | 主题: ' + $Subject + ' | 大小: ' + $payload.Length + ' 字节 | 附件: ' + $(if ($attachPath) { '有' } else { '无' }))
  exit 0
}

try {
  Send-SmtpRaw -SmtpHost $smtp.host -Port ([int]$smtp.port) -ImplicitSsl ([bool]$smtp.ssl) -UseStartTls ([bool]$smtp.starttls) -User $fromAddr -Pass $smtp.password -From $fromAddr -To $toList -Payload $payload | Out-Null
  if ($Tag) { Set-TagSent $Tag }
  Write-Host ('发送成功 -> ' + ($toList -join ', ') + ' | ' + $Subject)
  if ($Tag) { Write-Host ('已记录 tag: ' + $Tag) }
  exit 0
}
catch {
  Fail ('邮件发送失败: ' + $_.Exception.Message + ' 。常见原因: 授权码错误 / 端口或SSL配置不匹配 / 收件人地址错误。', 3)
}
