# transcribe.ps1 - 语音转文字（离线，使用 Windows 自带中文语音识别引擎）
# 用法: transcribe.ps1 -File <音频路径.wav> [-MaxSeconds 120]
# 说明: 仅支持 PCM WAV 格式（16kHz/22kHz 16bit 单声道最佳）。
#       其他格式(mp3/m4a/amr)需先转成 wav；实时语音建议用 Win+H Windows 语音输入。
# 输出: TEXT: <识别出的文字>   （退出码 0 成功 / 1 未识别到 / 4 参数或引擎问题）
param(
  [Parameter(Mandatory = $true)][string]$File,
  [int]$MaxSeconds = 120
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

if (-not (Test-Path $File)) { Write-Output ('ERROR 文件不存在: ' + $File); exit 4 }

Add-Type -AssemblyName System.Speech
$ri = [System.Speech.Recognition.SpeechRecognitionEngine]::InstalledRecognizers() |
  Where-Object { $_.Culture.Name -like 'zh*' } | Select-Object -First 1
if (-not $ri) { Write-Output 'ERROR 本机未安装中文语音识别引擎'; exit 4 }

$engine = New-Object System.Speech.Recognition.SpeechRecognitionEngine($ri.ID)
$engine.LoadGrammar((New-Object System.Speech.Recognition.DictationGrammar))
$engine.BabbleTimeout = [TimeSpan]::FromSeconds(15)
$engine.InitialSilenceTimeout = [TimeSpan]::FromSeconds(15)
$engine.EndSilenceTimeout = [TimeSpan]::FromSeconds(0.6)
$engine.SetInputToWaveFile((Resolve-Path $File).Path)

$parts = @()
$deadline = (Get-Date).AddSeconds($MaxSeconds)
while ((Get-Date) -lt $deadline) {
  try { $r = $engine.Recognize([TimeSpan]::FromSeconds(30)) }
  catch { break }  # 音频读尽后部分引擎会抛"无音频输入"，视为正常结束
  if ($null -eq $r) { break }
  if ([string]::IsNullOrWhiteSpace($r.Text)) { continue }
  $parts += $r.Text
}
$engine.Dispose()

$text = ($parts -join '')
if ([string]::IsNullOrWhiteSpace($text)) {
  Write-Output 'EMPTY 未识别到语音内容（确认是 PCM 格式的 wav，且音量正常）'
  exit 1
}
Write-Output ('TEXT: ' + $text)
exit 0
