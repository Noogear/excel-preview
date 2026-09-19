<#
.SYNOPSIS
  本地转换桥的"手脚"：用**本机 Excel** 把一份 xlsx 另存成别的格式，或把别的格式读成 xlsx。

.DESCRIPTION
  为什么用 Excel 自己来转（而不是第三方库）：轮子只认"值"，样式/条件格式/批注/图片全都丢；
  Excel 是这些格式的"原生实现"，它另存出来的文件与"人工打开再另存"完全等价 —— 保真的唯一可靠办法。
  见 `P5-导入导出多格式与库选型-方案.md` 第 1 节。

  由 Vite 中间件（`vite-plugins/vite-plugin-excel-bridge.ts`）调用：
    · 导出：先得到"当前内容的 xlsx 字节" → 本脚本 SaveAs 成 .ods/.xls/.xlsb
    · 导入：把 .xlsb 等交给 Excel 打开 → SaveAs 成 xlsx → 交回浏览器走自研解析器

  安全约定（与中间件一起构成边界）：
    · 只接受**文件路径**参数（路径由中间件在临时目录里生成，浏览器不能直接传路径）；
    · Excel 全程不可见、关闭所有警告与宏（AutomationSecurity=ForceDisable）；
    · 结束一定 Quit 并释放 COM；中间件另有超时兜底。

  **Excel 进程绝不留下**（这一条踩过坑，见下）：
    ① 脚本**只有一个出口**（末尾 `exit $exitCode`），Quit/Release 放在 `finally` 里 ——
       以前 `-Probe` 分支里是 `exit 0`，**直接跳过 Quit**，于是每一次"健康探测"都漏一个
       看不见的 EXCEL.EXE（实测本机攒到过 57 个）；
    ② 启动 Excel 之后**第一行 stdout 就把它的 PID 报出去**（`{"bridge":"excel-pid","pid":N}`）：
       中间件超时时会强杀本进程，`finally` 没机会跑，这时它按这个 PID 精确收拾那个 Excel；
    ③ 中间件那边还有"我们起过的 Excel"登记与退出清理（见 `tools/clean-orphan-excel.ps1`）。

.PARAMETER In
  输入文件（绝对路径）。

.PARAMETER Out
  输出文件（绝对路径）。

.PARAMETER FileFormat
  Excel 的 SaveAs 格式号：51=xlsx、52=xlsm、50=xlsb、56=xls、60=ods、6=GBK csv、62=UTF-8 csv。

.PARAMETER Probe
  只探测"本机 Excel 能不能用"，不转换。
#>
[CmdletBinding()]
param(
  [string]$In,
  [string]$Out,
  [int]$FileFormat = 51,
  [switch]$Probe
)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
$exitCode = 0
# 结果对象：全程只填这一份，最后统一输出（保证 stdout 的"最后一行"永远是结果 JSON）
$result = $null

function Release-Com([object]$obj) {
  if ($null -eq $obj) { return }
  try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
}

# 找到"这一次 New-Object 新起出来的那个 EXCEL.EXE"的 PID。
# 用"创建前后差集"判断：COM 不暴露 PID，而机器上可能还有别的 Excel（用户自己开的），
# 所以不能简单取第一个 EXCEL 进程。找不到就返回 $null（中间件会退化成"只杀本进程"）。
function Get-NewExcelPid([int[]]$before) {
  try {
    $after = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $fresh = @($after | Where-Object { $before -notcontains $_ })
    if ($fresh.Count -gt 0) { return [int]$fresh[0] }
  } catch { }
  return $null
}

try {
  # 找 Excel：注册过的 COM 组件最可靠；找不到就明确报"没装"
  $before = @()
  try { $before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch { }
  $excel = New-Object -ComObject Excel.Application
  $version = $null
  try { $version = [string]$excel.Version } catch { }

  # 先把 PID 报给中间件（它超时强杀时按这个 PID 收拾残留的 Excel）
  $excelPid = Get-NewExcelPid $before
  [pscustomobject]@{ bridge = 'excel-pid'; pid = $excelPid } | ConvertTo-Json -Compress

  if ($Probe) {
    $result = [pscustomobject]@{ available = $true; excel = "Excel $version" }
  }
  else {
    if ([string]::IsNullOrWhiteSpace($In) -or [string]::IsNullOrWhiteSpace($Out)) {
      throw 'In / Out 都是必填（绝对路径）'
    }
    if (-not (Test-Path -LiteralPath $In)) { throw "输入文件不存在：$In" }

    # 不可见 + 不弹窗 + 禁用宏（宏只保留、不执行，与产品策略一致）
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AutomationSecurity = 3 } catch { }  # msoAutomationSecurityForceDisable
    try { $excel.AskToUpdateLinks = $false } catch { }
    try { $excel.EnableEvents = $false } catch { }

    # UpdateLinks=0（不更新外部链接）、ReadOnly=$true（绝不动原文件）
    $workbook = $excel.Workbooks.Open($In, 0, $true)
    # 目标格式若与"宏格式"不同，Excel 会自己按目标格式降级（这正是我们要的：由 Excel 决定）
    $workbook.SaveAs($Out, $FileFormat)
    $workbook.Close($false)
    $workbook = $null

    if (-not (Test-Path -LiteralPath $Out)) { throw "Excel 没有产出文件：$Out" }
    $result = [pscustomobject]@{ ok = $true; out = $Out; bytes = (Get-Item -LiteralPath $Out).Length }
  }
}
catch {
  $exitCode = 1
  $result = [pscustomobject]@{ available = $false; ok = $false; reason = $_.Exception.Message }
}
finally {
  # 收摊：无论走哪条路（探测 / 转换 / 抛错）都要关掉工作簿、Quit、释放 COM
  if ($workbook) { try { $workbook.Close($false) } catch { } }
  if ($excel) {
    try { $excel.Quit() } catch { }
    # Quit 之后给 Excel 一点时间自己退；没退掉就在本进程内强杀（保证"脚本跑完 = 没有残留"）
    if ($excelPid) {
      for ($i = 0; $i -lt 20; $i++) {
        if (-not (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) { break }
        Start-Sleep -Milliseconds 100
      }
      if (Get-Process -Id $excelPid -ErrorAction SilentlyContinue) {
        try { Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue } catch { }
      }
    }
  }
  Release-Com $workbook
  Release-Com $excel
}

# 唯一的出口：结果 JSON 是 stdout 的最后一行（中间件取最后一行解析）
if ($result) { $result | ConvertTo-Json -Compress }
exit $exitCode
