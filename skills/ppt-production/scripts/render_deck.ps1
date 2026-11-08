# 导出 PPTX 为 PDF 并栅格化为逐页 PNG，供逐页读图验收。
# 用法：
#   ./render_deck.ps1 -Pptx "deck.pptx"                    # 全本
#   ./render_deck.ps1 -Pptx "deck.pptx" -Pages 3,8         # 只重渲染改动页
#   ./render_deck.ps1 -Pptx "deck.pptx" -Engine soffice    # 无 COM 环境
# 说明：只把“改完能出图”变成一条命令，不替代人工读图。
param(
  [Parameter(Mandatory=$true)][string]$Pptx,
  [string]$OutDir,
  [int[]]$Pages,
  [int]$Dpi = 120,
  [ValidateSet('auto','powerpoint','wps','soffice')][string]$Engine = 'auto',
  [string]$Pdftoppm
)
$ErrorActionPreference = 'Stop'

$Pptx = (Resolve-Path -LiteralPath $Pptx).Path
if (-not $OutDir) { $OutDir = Join-Path (Split-Path -Parent $Pptx) 'render' }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$OutDir = [IO.Path]::GetFullPath($OutDir)
if ($OutDir.TrimEnd('\','/') -eq (Split-Path -Parent $Pptx).TrimEnd('\','/')) { throw '输出目录必须与源文件目录不同。' }
if ($Dpi -le 0 -or @($Pages | Where-Object { $_ -le 0 }).Count -gt 0) { throw 'Dpi 和页号必须为正数。' }

$pdf = Join-Path $OutDir ([IO.Path]::GetFileNameWithoutExtension($Pptx) + '.pdf')
if (Test-Path -LiteralPath $pdf) { Remove-Item -LiteralPath $pdf -Force }

function Convert-With-Com($progId) {
  $app = New-Object -ComObject $progId
  $pres = $null
  try {
    $app.Visible = $true
    if ($progId -eq 'PowerPoint.Application') {
      $pres = $app.Presentations.Open($Pptx, $true, $false, $false)
    } else {
      $pres = $app.Presentations.Open($Pptx)
    }
    $pres.SaveAs($pdf, 32)   # 32 = PDF
  } finally {
    if ($pres) {
      $pres.Close()
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres)
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
  }
}

function Convert-With-Soffice {
  $exe = (Get-Command soffice -ErrorAction SilentlyContinue).Source
  if (-not $exe -and $env:ProgramFiles) {
    $c = Join-Path $env:ProgramFiles 'LibreOffice\program\soffice.exe'
    if (Test-Path -LiteralPath $c) { $exe = $c }
  }
  if (-not $exe) { throw '找不到 soffice，请安装 LibreOffice 或改用 -Engine auto。' }
  & $exe --headless --norestore --convert-to pdf --outdir $OutDir $Pptx | Out-Host
  if ($LASTEXITCODE -ne 0) { throw "soffice 退出码 $LASTEXITCODE。" }
  $produced = Join-Path $OutDir ([IO.Path]::GetFileNameWithoutExtension($Pptx) + '.pdf')
  if ($produced -ne $pdf -and (Test-Path -LiteralPath $produced)) { Move-Item -LiteralPath $produced -Destination $pdf -Force }
}

$order = switch ($Engine) {
  'powerpoint' { @('PowerPoint.Application') }
  'wps'        { @('Kwpp.Application') }
  'soffice'    { @() }
  default      { @('PowerPoint.Application', 'Kwpp.Application') }
}
$used = $null
foreach ($prog in $order) {
  try { Convert-With-Com $prog; $used = $prog; break }
  catch { Write-Host "跳过 $prog : $($_.Exception.Message)" }
}
if (-not $used -and $Engine -in @('auto','soffice')) {
  Convert-With-Soffice; $used = 'soffice'
}
if (-not $used) { throw '无法导出 PDF：所选引擎均失败。按当前环境检查权限或采用可验证的演示软件导出入口。' }
if (-not (Test-Path -LiteralPath $pdf)) { throw "导出报告成功但未生成 PDF：$pdf" }
Write-Host "PDF 已生成（$used）：$pdf  $((Get-Item -LiteralPath $pdf).Length) B"

# 栅格化：-Pdftoppm 参数 > PATH > 常见 bundled 位置。
$candidates = @()
if ($Pdftoppm) { $candidates += $Pdftoppm }
$onPath = Get-Command pdftoppm -ErrorAction SilentlyContinue
if ($onPath) { $candidates += $onPath.Source }
if ($env:CODEX_HOME) {
  $candidates += Get-ChildItem -LiteralPath $env:CODEX_HOME -Filter 'pdftoppm.exe' -Recurse -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName }
}
$candidates += @('/usr/bin/pdftoppm','/usr/local/bin/pdftoppm','/opt/homebrew/bin/pdftoppm')
$tool = $candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
if (-not $tool) { throw '找不到 pdftoppm。安装 poppler 后用 -Pdftoppm 指定完整路径。' }

$prefixName = [IO.Path]::GetFileNameWithoutExtension($Pptx) + '-page'
$prefix = Join-Path $OutDir $prefixName
$previous = @(Get-ChildItem -LiteralPath $OutDir -File | Where-Object { $_.Name -match ('^' + [regex]::Escape($prefixName) + '-(\d+)\.png$') })
foreach ($image in $previous) {
  [void]($image.Name -match ('^' + [regex]::Escape($prefixName) + '-(\d+)\.png$'))
  if (-not $Pages -or $Pages -contains [int]$Matches[1]) { Remove-Item -LiteralPath $image.FullName -Force }
}
if ($Pages -and $Pages.Count -gt 0) {
  foreach ($n in $Pages) {
    & $tool -png -r $Dpi -f $n -l $n $pdf $prefix
    if ($LASTEXITCODE -ne 0) { throw "pdftoppm 在第 $n 页失败：$LASTEXITCODE。" }
  }
} else {
  & $tool -png -r $Dpi $pdf $prefix
}
if ($LASTEXITCODE -ne 0) { throw "pdftoppm 退出码 $LASTEXITCODE。" }
$made = Get-ChildItem -LiteralPath $OutDir -File | Where-Object {
  $_.Name -match ('^' + [regex]::Escape($prefixName) + '-(\d+)\.png$') -and (-not $Pages -or $Pages -contains [int]$Matches[1])
} | Sort-Object Name
if (-not $made) { throw '栅格化未产出 PNG。' }
Write-Host ("渲染页： " + (($made | ForEach-Object { $_.Name }) -join ', '))
Write-Host "输出目录：$OutDir"
