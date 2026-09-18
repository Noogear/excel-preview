#Requires -Version 5.1
<#
  一键运行脚本的辅助进程（脱离控制台、隐藏窗口运行）。两种用法：

  1) 看门狗：-ParentPid <pid> -Port <port>
     父脚本（start.ps1）消失后，清掉占用该端口的服务进程。
     为什么需要它：窗口被强杀时父脚本的 finally 来不及执行，光靠"前台跑 vite"会留下孤儿 node。

  2) 开浏览器：-OpenUrl <url>
     轮询等到该地址返回 200 再打开默认浏览器（避免比服务先打开导致"无法访问"）。
     最多等 90 秒，然后自己退出。
#>
[CmdletBinding()]
param(
  [int]$ParentPid = 0,
  [int]$Port = 0,
  [string]$OpenUrl = ''
)

$ErrorActionPreference = 'SilentlyContinue'

function Get-PortOwnerPid([int]$p) {
  try {
    $conn = Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction Stop | Select-Object -First 1
    if ($conn) { return [int]$conn.OwningProcess }
  } catch {
    # 老系统没有 Get-NetTCPConnection，退回 netstat
  }
  $line = netstat -ano | Select-String -Pattern (":$p\s+\S+\s+LISTENING") | Select-Object -First 1
  if ($line) {
    $parts = ($line.Line -split '\s+') | Where-Object { $_ -ne '' }
    return [int]$parts[-1]
  }
  return 0
}

function Test-PortListening([int]$p) {
  return (Get-PortOwnerPid $p) -gt 0
}

# ---- 用法 2：等服务起来再开浏览器 ----
if ($OpenUrl -ne '') {
  $deadline = (Get-Date).AddSeconds(90)
  while ((Get-Date) -lt $deadline) {
    try {
      $res = Invoke-WebRequest -Uri $OpenUrl -UseBasicParsing -TimeoutSec 3
      if ($res.StatusCode -eq 200) {
        Start-Process $OpenUrl
        exit 0
      }
    } catch {
      # 还没起来，继续等
    }
    Start-Sleep -Milliseconds 400
  }
  exit 0
}

# ---- 用法 1：看门狗 ----
if ($ParentPid -gt 0 -and $Port -gt 0) {
  # 父进程还在 → 每秒看一眼；父进程没了（窗口被关/被强杀）→ 清掉端口上的服务
  # 注意：**不能**一上来就"端口没在监听就退出"——看门狗比服务先启动，
  # 第一轮检查时服务通常还没起来，那样它会立刻自杀（这个坑实测踩过）。
  # 正确判据：父进程还活着就一直守着；只有"曾经监听过、后来不监听了"才说明服务已正常退出。
  $sawListening = $false
  while ($true) {
    Start-Sleep -Seconds 1
    $parent = Get-Process -Id $ParentPid -ErrorAction SilentlyContinue
    if (-not $parent) { break }
    if (Test-PortListening $Port) {
      $sawListening = $true
    } elseif ($sawListening) {
      exit 0   # 服务已正常停止，看门狗下班
    }
  }

  $ownerPid = Get-PortOwnerPid $Port
  if ($ownerPid -gt 0) {
    Stop-Process -Id $ownerPid -Force -ErrorAction SilentlyContinue
  }
  exit 0
}

exit 0
