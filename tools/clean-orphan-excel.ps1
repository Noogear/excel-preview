<#
.SYNOPSIS
  收拾"本地转换桥"可能留下的、看不见的 Excel（COM 自动化实例）。

.DESCRIPTION
  为什么要单独有这么一个脚本：桥的正常路径与异常路径都会 Quit Excel（`tools/excel-bridge.ps1` 的
  `finally`），超时路径也会由中间件按 PID 精确补一刀。但仍有兜不住的情况：
    · 开发服务器被强杀（任务管理器结束进程 / 关窗口）时，脚本连同它的 finally 一起消失；
    · 更早版本的脚本在 `-Probe` 分支里 `exit 0`，**直接跳过 Quit** —— 每次健康探测漏一个
      （实测在本机攒到过 57 个无窗口的 EXCEL.EXE，把内存吃满、让后续转换从 5 秒退化到 60 秒）。
  这个脚本就是给这种情况用的**手工兜底**（也方便排查时确认"到底还有没有残留"）。

  安全边界（重要）：只杀**命令行里带 `-Embedding`（或 `/automation`）的 EXCEL.EXE**，
  也就是"由 COM 自动化启动"的那种；你自己打开的 Excel 命令行里没有这个标记，绝不会被误杀。
  拿不到命令行时（CIM 不可用）会退化为"没有可见窗口的 Excel"，并且**先打印再确认式地列出**，
  可以用 `-DryRun` 只看不杀。

.PARAMETER DryRun
  只列出将要结束的进程，不动手。

.EXAMPLE
  npm run bridge:clean          # 收拾残留
  pwsh -File tools/clean-orphan-excel.ps1 -DryRun   # 只看有哪些
#>
[CmdletBinding()]
param([switch]$DryRun)

$ErrorActionPreference = 'Continue'

function Get-AutomationExcel {
  try {
    $procs = Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction Stop
    return @($procs | Where-Object { $_.CommandLine -and ($_.CommandLine -match '-Embedding' -or $_.CommandLine -match '/automation') })
  }
  catch {
    Write-Warning "拿不到进程命令行（CIM/WMI 不可用），退化为"没有可见窗口的 Excel"判定 —— 这个口径更宽，请先看 -DryRun 的输出再决定"
    return @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -eq 0 } |
        ForEach-Object { [pscustomobject]@{ ProcessId = $_.Id; CommandLine = '(命令行不可用)' } })
  }
}

$targets = Get-AutomationExcel
if ($targets.Count -eq 0) {
  Write-Output '[bridge:clean] 没有发现转换桥留下的 Excel 残留'
  exit 0
}

Write-Output "[bridge:clean] 发现 $($targets.Count) 个 COM 自动化 Excel："
foreach ($p in $targets) { Write-Output ("  PID {0}  {1}" -f $p.ProcessId, $p.CommandLine) }

if ($DryRun) {
  Write-Output '[bridge:clean] -DryRun：只看不动手'
  exit 0
}

$killed = 0
foreach ($p in $targets) {
  try {
    Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
    $killed += 1
  }
  catch {
    Write-Warning ("PID {0} 结束失败：{1}" -f $p.ProcessId, $_.Exception.Message)
  }
}
Write-Output "[bridge:clean] 已结束 $killed / $($targets.Count) 个；剩余 EXCEL 进程 $((Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Measure-Object).Count) 个"
