param([int]$Year = 0, [int]$Wait = 180, [string]$Exe = "C:\PVI\Pias2022new\Pias2022.exe")
# Make sure PIAS is open and logged in (saved account). Starts it when needed; -Year sets the login year.
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type @"
using System; using System.Runtime.InteropServices;
public class UL { [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
[DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, string l); }
"@
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$ALL = [System.Windows.Automation.Condition]::TrueCondition
$EXE = $Exe
function Title { $p = Get-Process Pias2022 -ErrorAction SilentlyContinue | Select-Object -First 1; if ($p) { $p.MainWindowTitle } }
function WaitFor([scriptblock]$cond, [int]$sec) {
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $sec) { if (& $cond) { return $true }; Start-Sleep -Milliseconds 1000 }
  return $false
}
if ((Title) -like 'Insurances Solution*') { "PIAS ready"; exit 0 }
if (-not (Get-Process Pias2022 -ErrorAction SilentlyContinue)) {
  Start-Process $EXE -WorkingDirectory (Split-Path $EXE)
}
if (-not (WaitFor { (Title) -eq 'Đăng nhập' -or (Title) -like 'Insurances Solution*' } $Wait)) { throw "PIAS login window did not appear" }
if ((Title) -like 'Insurances Solution*') { "PIAS ready"; exit 0 }
$p = Get-Process Pias2022 | Select-Object -First 1
$win = $AE::FromHandle($p.MainWindowHandle)
$els = @($win.FindAll($TS::Descendants, $ALL))
if ($Year) {
  $y = $els | Where-Object { $_.Current.ClassName -like '*EDIT*' -and $_.Current.Name -match '^\d{4}$' } | Select-Object -First 1
  if ($y -and $y.Current.Name -ne "$Year") { [UL]::SendMessage([IntPtr]$y.Current.NativeWindowHandle, 0x000C, [IntPtr]0, "$Year") | Out-Null; "year set to $Year" }
}
$ok = $els | Where-Object { $_.Current.Name -eq 'Chấp nhận' } | Select-Object -First 1
$lp = [IntPtr]((12 -shl 16) -bor 43)
foreach ($m in 0x0200, 0x0201, 0x0202) { [UL]::PostMessage([IntPtr]$ok.Current.NativeWindowHandle, $m, [IntPtr]($m -eq 0x0201), $lp) | Out-Null }
if (-not (WaitFor { (Title) -like 'Insurances Solution*' } $Wait)) { throw "PIAS did not log in" }
"PIAS ready"