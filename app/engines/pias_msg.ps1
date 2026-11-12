[Console]::OutputEncoding = [System.Text.Encoding]::UTF8   # the app reads this output as UTF-8
# If PIAS shows a message box (title 'Insurace Solution 2022'), print its text and press OK. Prints nothing otherwise.
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type @"
using System; using System.Runtime.InteropServices;
public class UM { [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l); }
"@
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$ALL = [System.Windows.Automation.Condition]::TrueCondition
$p = Get-Process Pias2022 -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $p -or $p.MainWindowHandle -eq 0) { exit 0 }
foreach ($w in $AE::FromHandle($p.MainWindowHandle).FindAll($TS::Children, $ALL)) {
  if ($w.Current.Name -ne 'Insurace Solution 2022') { continue }
  $els = @($w.FindAll($TS::Descendants, $ALL))
  $text = ($els | Where-Object { $_.Current.Name -and $_.Current.Name -notin 'OK', 'Insurace Solution 2022' } |
           ForEach-Object { $_.Current.Name } | Select-Object -Unique) -join ' '
  $ok = $els | Where-Object { $_.Current.Name -eq 'OK' } | Select-Object -First 1
  if (-not $ok) { continue }
  try {                           # keep a picture of every box answered (app\dump\msgs), the text is painted
    Add-Type -AssemblyName System.Drawing
    Add-Type @"
using System; using System.Runtime.InteropServices;
public class PWm { [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f); }
"@
    $r = $w.Current.BoundingRectangle
    $bmp = New-Object System.Drawing.Bitmap ([int]$r.Width), ([int]$r.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp); $dc = $g.GetHdc()
    [PWm]::PrintWindow([IntPtr]$w.Current.NativeWindowHandle, $dc, 2) | Out-Null; $g.ReleaseHdc($dc); $g.Dispose()
    $dir = Join-Path (Split-Path $PSScriptRoot) "dump\msgs"; New-Item -ItemType Directory -Force $dir | Out-Null
    $bmp.Save((Join-Path $dir ("msg_" + (Get-Date -Format yyyyMMdd_HHmmss) + ".png"))); $bmp.Dispose()
  } catch {}
  $lp = [IntPtr]((11 -shl 16) -bor 37)
  foreach ($m in 0x0200, 0x0201, 0x0202) { [UM]::PostMessage([IntPtr]$ok.Current.NativeWindowHandle, $m, [IntPtr][int]($m -eq 0x0201), $lp) | Out-Null }
  "MSG:" + [int]$w.Current.BoundingRectangle.Width + ": $text"
}