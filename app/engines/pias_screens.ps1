param([switch]$CloseExtra, [switch]$CloseAll)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# How many copies of "Tìm kiếm đơn Bảo hiểm" are open, and close the spare ones.
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type @"
using System; using System.Runtime.InteropServices;
public class SC { [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l); }
"@
$AE = [System.Windows.Automation.AutomationElement]; $TS = [System.Windows.Automation.TreeScope]
$ALL = [System.Windows.Automation.Condition]::TrueCondition
$p = Get-Process Pias2022 -ErrorAction Stop | Select-Object -First 1
$root = $AE::FromHandle($p.MainWindowHandle)
$wins = @($root.FindAll($TS::Children, $ALL) | Where-Object { $_.Current.Name -eq 'Tìm kiếm đơn Bảo hiểm' })
"screens: $($wins.Count)"
if ($CloseExtra -or $CloseAll) {
  $skip = if ($CloseAll) { 0 } else { 1 }
  for ($i = $skip; $i -lt $wins.Count; $i++) {
    $b = @($wins[$i].FindAll($TS::Descendants, $ALL) | Where-Object { $_.Current.Name -eq 'Thoát' })[0]
    if ($b) {
      $h = [IntPtr]$b.Current.NativeWindowHandle
      $lp = [IntPtr]((11 -shl 16) -bor 30)
      foreach ($m in 0x0200, 0x0201, 0x0202) { [SC]::PostMessage($h, $m, [IntPtr][int]($m -eq 0x0201), $lp) | Out-Null }
      Start-Sleep -Milliseconds 700
    }
  }
  $left = @($root.FindAll($TS::Children, $ALL) | Where-Object { $_.Current.Name -eq 'Tìm kiếm đơn Bảo hiểm' })
  "left: $($left.Count)"
}
