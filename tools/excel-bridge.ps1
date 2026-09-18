<#
.SYNOPSIS
  本地转换桥的"手脚"：用**本机 Excel** 把一份 xlsx 另存成别的格式，或把别的格式读成 xlsx。

.DESCRIPTION
  为什么用 Excel 自己来转（而不是第三方库）：轮子只认"值"，样式/条件格式/批注/图片全都丢；
  Excel 是这些格式的"原生实现"，它另存出来的文件与"人工打开再另存"完全等价 —— 保真的唯一可靠办法。
  见 `P5-导入导出多格式与库选型-方案.md` 第 1 节。

  由 Vite 中间件（`build/vite-plugin-excel-bridge.ts`）调用：
    · 导出：先得到"当前内容的 xlsx 字节" → 本脚本 SaveAs 成 .ods/.xls/.xlsb
    · 导入：把 .xlsb 等交给 Excel 打开 → SaveAs 成 xlsx → 交回浏览器走自研解析器

  安全约定（与中间件一起构成边界）：
    · 只接受**文件路径**参数（路径由中间件在临时目录里生成，浏览器不能直接传路径）；
    · Excel 全程不可见、关闭所有警告与宏（AutomationSecurity=ForceDisable）；
    · 结束一定 Quit 并释放 COM；中间件另有超时兜底。

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

function Release-Com([object]$obj) {
  if ($null -eq $obj) { return }
  try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
}

try {
  # 找 Excel：注册过的 COM 组件最可靠；找不到就明确报"没装"
  $excel = New-Object -ComObject Excel.Application
  $version = $null
  try { $version = [string]$excel.Version } catch { }

  if ($Probe) {
    [pscustomobject]@{ available = $true; excel = "Excel $version" } | ConvertTo-Json -Compress
    exit 0
  }

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
  $excel.Quit()
  Release-Com $excel
  $excel = $null

  if (-not (Test-Path -LiteralPath $Out)) { throw "Excel 没有产出文件：$Out" }
  [pscustomobject]@{ ok = $true; out = $Out; bytes = (Get-Item -LiteralPath $Out).Length } | ConvertTo-Json -Compress
  exit 0
}
catch {
  $message = $_.Exception.Message
  # 收摊：能关就关，避免留下看不见的 EXCEL.EXE 进程
  if ($workbook) { try { $workbook.Close($false) } catch { } }
  if ($excel) { try { $excel.Quit() } catch { } }
  Release-Com $workbook
  Release-Com $excel
  [pscustomobject]@{ available = $false; ok = $false; reason = $message } | ConvertTo-Json -Compress
  exit 1
}
