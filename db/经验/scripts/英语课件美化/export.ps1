param([string]$pptx, [string]$outdir, [int]$width = 1280, [string]$slides = "")
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force $outdir | Out-Null
$app = New-Object -ComObject PowerPoint.Application
if ($app.Path -notlike "*Microsoft Office*") { throw "ABORT: PowerPoint automation is served by $($app.Path) (WPS?). Close WPS and re-run." }
$pres = $app.Presentations.Open($pptx, $true, $false, $false)
$h = [int]($width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
"Size: $($pres.PageSetup.SlideWidth) x $($pres.PageSetup.SlideHeight)"
$list = @()
if ($slides -ne "") { $list = $slides.Split(",") | ForEach-Object { [int]$_ } } else { $list = 1..$pres.Slides.Count }
foreach ($i in $list) {
  $f = Join-Path $outdir ("s{0:D2}.png" -f $i)
  $pres.Slides.Item($i).Export($f, "PNG", $width, $h)
}
$pres.Close()
"done $($list.Count)"
