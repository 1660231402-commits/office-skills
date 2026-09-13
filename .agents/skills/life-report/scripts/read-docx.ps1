# read-docx.ps1 - 提取 .docx 文档正文文本（纯 PowerShell，无外部依赖）
# 用法: read-docx.ps1 -File <文档.docx>
# 原理: docx 是 zip 包，读取 word/document.xml，按段落还原文本。
param([Parameter(Mandatory = $true)][string]$File)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

if (-not (Test-Path $File)) { Write-Output ('ERROR 文件不存在: ' + $File); exit 4 }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $File).Path)
try {
  $entry = $zip.GetEntry('word/document.xml')
  if (-not $entry) { Write-Output 'ERROR 不是有效的 docx（缺少 word/document.xml）'; exit 4 }
  $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
  $xml = $reader.ReadToEnd()
  $reader.Close()
}
finally { $zip.Dispose() }

$text = $xml -replace '<w:br ?[^>]*/>', "`n" `
             -replace '<w:tab ?[^>]*/>', "`t" `
             -replace '</w:p>', "`n" `
             -replace '<[^>]+>', ''
$text = [System.Net.WebUtility]::HtmlDecode($text)
$text = ($text -split "`n" | ForEach-Object { $_.TrimEnd() }) -join "`n"
if ([string]::IsNullOrWhiteSpace($text)) { Write-Output 'EMPTY 文档没有正文文本'; exit 1 }
Write-Output $text
exit 0
