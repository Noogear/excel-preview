<#
.SYNOPSIS
  用本机 Excel（COM）把 fixtures/*.xlsx 另存成**真实**的 .xls / .ods / .csv 样本。

.DESCRIPTION
  为什么要真机转换、而不是自己写一个：
    - `.xls` 是 Excel 97–2003 的 BIFF8 **二进制**格式，`.ods` 是 OpenDocument 的 zip+XML，
      两者与我们自研的 OOXML 解析链路完全不同。要支持它们就得自己写解析器，
      而"自己写个 writer 再造样本给自己测"是循环论证；用真 Excel 产出样本才是**真值**。
    - `.csv` 还要覆盖编码差异：中文 Windows 上 `xlCSV` 写出的是 **GBK（ANSI）**，
      `xlCSVUTF8` 写出的是 UTF-8（带 BOM）。两条都要测。

  没有装 Excel 的机器上这个脚本会打印提示并跳过；对应测试用例会用
  `test.skip(!existsSync(...))` 自动跳过（与既有"缺夹具就跳过"的约定一致）。
  个别文件 Excel 自己打不开（例如某些第三方库写出的 ListObject）会打印警告并继续，
  不影响其余样本。

  运行：npm run fixtures:legacy
#>
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$FixtureDir = Join-Path $Root 'fixtures'

# xlFileFormat 枚举里我们用到的几个
$FORMATS = @(
  @{ Ext = '.xls'; Format = 56 },        # xlExcel8（BIFF8）
  @{ Ext = '.ods'; Format = 60 },        # xlOpenDocumentSpreadsheet
  @{ Ext = '-gbk.csv'; Format = 6 },     # xlCSV（系统 ANSI，中文 Windows 上是 GBK）
  @{ Ext = '-utf8.csv'; Format = 62 }    # xlCSVUTF8
)

$sources = @(Get-ChildItem -Path $FixtureDir -Filter 'fixture-*.xlsx' -File)
if ($sources.Count -eq 0) {
  Write-Host '先运行 npm run fixtures 生成 xlsx 样本。' -ForegroundColor Yellow
  exit 0
}

try {
  $excel = New-Object -ComObject Excel.Application
} catch {
  Write-Host '本机没有可用的 Excel（COM 不可用），跳过 .xls/.ods/.csv 样本生成。' -ForegroundColor Yellow
  Write-Host '（这些样本不是测试的必需品：缺失时相关用例会自动跳过。）'
  exit 0
}

$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.AskToUpdateLinks = $false

$made = 0
$failed = @()
foreach ($src in $sources) {
  $stem = [System.IO.Path]::GetFileNameWithoutExtension($src.Name)
  $wb = $null
  try {
    # 只读打开；个别文件 Excel 拒绝打开时不中断整个脚本
    $wb = $excel.Workbooks.Open($src.FullName, 0, $true)
    foreach ($format in $FORMATS) {
      $target = Join-Path $FixtureDir ($stem + $format.Ext)
      if (Test-Path $target) { Remove-Item $target -Force }
      $wb.SaveAs($target, $format.Format)
      $made++
      Write-Host ("  ✓ " + [System.IO.Path]::GetFileName($target))
    }
  } catch {
    $failed += "$($src.Name)（$($_.Exception.Message)）"
    Write-Host ("  ! 跳过 " + $src.Name + "：" + $_.Exception.Message) -ForegroundColor Yellow
  } finally {
    if ($wb) { try { $wb.Close($false) } catch { } }
  }
}

try { $excel.Quit() } catch { }
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()

Write-Host "✓ 生成 $made 个 .xls / .ods / .csv 样本（真 Excel 产出，作为解析器真值）" -ForegroundColor Green
if ($failed.Count -gt 0) {
  Write-Host ("跳过 " + $failed.Count + " 个源文件：" + ($failed -join '；')) -ForegroundColor Yellow
}
