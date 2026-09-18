#Requires -Version 5.1
<#
  excel-preview 一键启动脚本（由 run.cmd 调用）

  只做一件事：把工具跑起来。检查环境 → 首次装依赖 → 确保有示例 → 起服务 → 打开浏览器。
  测试 / 基准这类额外任务请用 npm 脚本（npm test / npm run e2e / npm run bench:parse）。

  设计要点：
  - **前台直接跑 vite**（不经 npm 包一层）：本窗口一关（或 Ctrl+C），vite 随之结束，
    不会留下"看不见还在跑"的 node 进程。
  - 额外挂一个**脱离控制台的看门狗**（tools/run-helper.ps1）：即使窗口被强杀、finally 来不及执行，
    它也会在本脚本进程消失后清掉占用端口的服务。
  - 端口被占时不硬闯：若占用者就是本项目已跑着的服务，直接开浏览器复用；否则自动换端口。
#>
[CmdletBinding()]
param(
  [int]$Port = 5273,
  [switch]$NoBrowser,
  [switch]$SkipInstall
)

$ErrorActionPreference = 'Stop'
$script:Root = Split-Path -Parent $PSScriptRoot   # excel-preview 根目录
$script:Helper = Join-Path $PSScriptRoot 'run-helper.ps1'
Set-Location $script:Root

# ---------------------------------------------------------------- 输出小工具
function Write-Title([string]$text) { Write-Host ''; Write-Host "  $text" -ForegroundColor Cyan }
function Write-Step([string]$text) { Write-Host "  → $text" -ForegroundColor Gray }
function Write-Ok([string]$text) { Write-Host "  ✓ $text" -ForegroundColor Green }
function Write-Warn2([string]$text) { Write-Host "  ! $text" -ForegroundColor Yellow }
function Write-Err([string]$text) { Write-Host "  ✗ $text" -ForegroundColor Red }

function Write-Banner {
  Write-Host ''
  Write-Host '  ╔══════════════════════════════════════════════════════════╗' -ForegroundColor DarkCyan
  Write-Host '  ║   excel-preview · 纯前端 Excel 高保真预览与编辑工具      ║' -ForegroundColor DarkCyan
  Write-Host '  ╚══════════════════════════════════════════════════════════╝' -ForegroundColor DarkCyan
}

# ---------------------------------------------------------------- 环境与端口
function Get-ShellExe {
  $pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
  if ($pwsh) { return $pwsh.Source }
  return (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe')
}

function Test-PortFree([int]$p) {
  try {
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $p)
    $listener.Start()
    $listener.Stop()
    return $true
  } catch {
    return $false
  }
}

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

function Stop-PortListener([int]$p) {
  $ownerPid = Get-PortOwnerPid $p
  if ($ownerPid -le 0) { return $false }
  try {
    $proc = Get-Process -Id $ownerPid -ErrorAction Stop
    Write-Step ("结束占用端口 {0} 的进程 {1}（{2}）" -f $p, $ownerPid, $proc.ProcessName)
    Stop-Process -Id $ownerPid -Force -ErrorAction Stop
    return $true
  } catch {
    Write-Warn2 ("无法结束进程 {0}：{1}" -f $ownerPid, $_.Exception.Message)
    return $false
  }
}

function Test-IsOurApp([int]$p) {
  try {
    $res = Invoke-WebRequest -Uri ("http://localhost:{0}/" -f $p) -UseBasicParsing -TimeoutSec 3
    return ($res.StatusCode -eq 200) -and ($res.Content -match 'id="root"')
  } catch {
    return $false
  }
}

function Start-Watchdog {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][int]$ParentProcessId,
    [Parameter(Mandatory)][int]$WatchPort
  )
  if (-not (Test-Path $script:Helper)) { Write-Warn2 '找不到 run-helper.ps1，跳过看门狗（窗口强杀时可能留下进程）'; return }
  # 两处坑记在这里：
  # ① 函数参数名必须与调用处一致——简单函数会把绑不上的参数悄悄塞进 $args，
  #    之前参数叫 $p 却用 -Port 调用，看门狗收到 -Port 0 就自杀了，表面看是"没启动"。
  # ② -ExecutionPolicy 是 pwsh.exe 的**命令行参数**，必须放进 -ArgumentList；
  #    当成 Start-Process 的参数写会直接报错（这个错误靠 [CmdletBinding()] 才暴露出来）。
  $proc = Start-Process -FilePath (Get-ShellExe) -WindowStyle Hidden -PassThru -ArgumentList @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script:Helper,
    '-ParentPid', $ParentProcessId, '-Port', $WatchPort
  )
  Write-Step ("看门狗已启动（PID {0}）：本窗口一旦消失，它会自动释放端口 {1}" -f $proc.Id, $WatchPort)
}

function Start-BrowserOpener {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Url)
  if (-not (Test-Path $script:Helper)) { return }
  Start-Process -FilePath (Get-ShellExe) -WindowStyle Hidden -ArgumentList @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script:Helper, '-OpenUrl', $Url
  ) | Out-Null
}

function Ensure-Node {
  $node = Get-Command node -ErrorAction SilentlyContinue
  if (-not $node) {
    Write-Err '找不到 node。请先安装 Node.js 20+（https://nodejs.org）后重试。'
    exit 2
  }
  $version = (& node -v).TrimStart('v')
  $major = [int]($version -split '\.')[0]
  Write-Ok ("Node {0}" -f $version)
  if ($major -lt 18) { Write-Warn2 'Node 版本偏低（项目按 Node 20+ 开发），可能启动失败。' }
}

function Ensure-Deps {
  if ($SkipInstall) { Write-Step '按参数跳过依赖检查'; return }
  $vitePkg = Join-Path $script:Root 'node_modules\vite\package.json'
  if (Test-Path $vitePkg) { Write-Ok '依赖已就绪'; return }
  Write-Step '首次运行：安装依赖（npm install，约 1–3 分钟）…'
  & npm install
  if ($LASTEXITCODE -ne 0) {
    Write-Err 'npm install 失败。若报缓存/权限错误，可试：npm install --cache .npm-cache'
    exit 3
  }
  Write-Ok '依赖安装完成'
}

function Ensure-Fixtures {
  # 只是"有现成示例可以先试"，不属于测试：缺了才生成，已有就跳过
  $manifest = Join-Path $script:Root 'fixtures\manifest.json'
  if (Test-Path $manifest) { return }
  Write-Step '生成示例表格（fixtures/*.xlsx）…'
  & node 'tools\make-fixtures.mjs' | Out-Host
  if ($LASTEXITCODE -ne 0) { Write-Warn2 '示例表格生成失败（不影响使用，可自己拖入 xlsx）'; return }
  Write-Ok '示例表格已就绪'
}

# ---------------------------------------------------------------- 主流程
Write-Banner

Write-Title '1/3 检查运行环境'
Ensure-Node
Ensure-Deps
Ensure-Fixtures

Write-Title '2/3 选择端口'
$usePort = $Port
if (-not (Test-PortFree $usePort)) {
  if (Test-IsOurApp $usePort) {
    Write-Ok ("端口 {0} 上已经跑着本工具，直接复用" -f $usePort)
    $existing = "http://localhost:$usePort/"
    Write-Host ''
    Write-Host ("  已打开：{0}" -f $existing) -ForegroundColor Green
    if (-not $NoBrowser) { Start-Process $existing }
    Write-Host '  （这个窗口可以直接关掉：服务由另一个窗口负责）' -ForegroundColor DarkGray
    exit 0
  }
  Write-Warn2 ("端口 {0} 被其它程序占用，自动顺延" -f $usePort)
  $found = $false
  for ($p = $Port + 1; $p -le $Port + 20; $p++) {
    if (Test-PortFree $p) { $usePort = $p; $found = $true; break }
  }
  if (-not $found) { Write-Err '连续 20 个端口都被占用，请先释放端口或用 -Port 指定'; exit 5 }
}
Write-Ok ("使用端口 {0}" -f $usePort)

$url = "http://localhost:$usePort/"
$vite = Join-Path $script:Root 'node_modules\vite\bin\vite.js'
if (-not (Test-Path $vite)) { Write-Err '找不到 vite，请先执行 npm install'; exit 3 }

Write-Title '3/3 启动服务'
Start-Watchdog -ParentProcessId $PID -WatchPort $usePort
if (-not $NoBrowser) { Start-BrowserOpener -Url $url }

Write-Host ''
Write-Host ("  地址：{0}" -f $url) -ForegroundColor Green
Write-Host '  导入 fixtures/ 下的示例即可查看效果；右侧「事件日志」显示导入报告。' -ForegroundColor DarkGray
Write-Host '  关闭本窗口（或 Ctrl+C）即结束服务，会自动清理进程。' -ForegroundColor Yellow
Write-Host ''

try {
  # 前台跑 vite：这个窗口就是服务进程，关窗即停
  & node $vite --port $usePort --strictPort
} finally {
  if (-not (Test-PortFree $usePort)) {
    Write-Host ''
    Write-Step '正在结束服务…'
    Stop-PortListener $usePort | Out-Null
  }
  Write-Ok '服务已停止，端口已释放'
}
