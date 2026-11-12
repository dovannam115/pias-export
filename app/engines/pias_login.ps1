param([int]$Year = 0, [int]$Wait = 180, [string]$Exe = "C:\PVI\Pias2022new\Pias2022.exe")
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8   # the app reads this output as UTF-8
# Make sure PIAS is open and logged in (saved account). Starts it when needed; -Year sets the login year.
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type @"
using System; using System.Runtime.InteropServices;
public class UL { [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
[DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, string l);
[DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr p);
[DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
[DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool f);
[DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr h); }
"@
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$ALL = [System.Windows.Automation.Condition]::TrueCondition
$EXE = $Exe
function Title { $p = Get-Process Pias2022 -ErrorAction SilentlyContinue | Select-Object -First 1; if ($p) { $p.MainWindowTitle } }
function DismissUpdate([int]$sec = 0) {
  # PIAS asks "Chương trình có phiên bản mới trên máy chủ..." (title = install folder): always answer No
  $t0 = Get-Date
  do {
    $p = Get-Process Pias2022 -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($p -and $p.MainWindowHandle -ne 0) {
      $boxes = @($AE::FromHandle($p.MainWindowHandle).FindAll($TS::Children, $ALL) | Where-Object { $_.Current.Name -like 'C:\PVI\*' })
      foreach ($b in $boxes) {
        $no = @($b.FindAll($TS::Descendants, $ALL) | Where-Object { $_.Current.Name -eq 'No' })[0]
        if ($no) {
          $lp = [IntPtr]((11 -shl 16) -bor 37)
          foreach ($m in 0x0200, 0x0201, 0x0202) { [UL]::PostMessage([IntPtr]$no.Current.NativeWindowHandle, $m, [IntPtr][int]($m -eq 0x0201), $lp) | Out-Null }
          "update prompt: answered No"; Start-Sleep -Milliseconds 800; return
        }
      }
    }
    if ($sec) { Start-Sleep -Milliseconds 700 }
  } while (((Get-Date) - $t0).TotalSeconds -lt $sec)
}
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
  if ($y -and $y.Current.Name -ne "$Year") {
    # typed, not set: a set text only changes what is shown (same as the date boxes)
    $h = [IntPtr]$y.Current.NativeWindowHandle
    $tid = [UL]::GetWindowThreadProcessId($h, [IntPtr]::Zero); $me = [UL]::GetCurrentThreadId()
    [UL]::AttachThreadInput($me, $tid, $true) | Out-Null
    try {
      [UL]::SetFocus($h) | Out-Null; Start-Sleep -Milliseconds 200
      [UL]::SendMessage($h, 0x00B1, [IntPtr]0, [IntPtr](-1)) | Out-Null            # select all
      foreach ($c in "$Year".ToCharArray()) { [UL]::PostMessage($h, 0x0102, [IntPtr][int]$c, [IntPtr]0) | Out-Null; Start-Sleep -Milliseconds 80 }
      Start-Sleep -Milliseconds 500
    } finally { [UL]::AttachThreadInput($me, $tid, $false) | Out-Null }
    $now = $AE::FromHandle($h).Current.Name
    if ($now -ne "$Year") { throw "could not set the login year (shows $now)" }
    "year set to $Year"
  }
}
$ok = $els | Where-Object { $_.Current.Name -eq 'Chấp nhận' } | Select-Object -First 1
$lp = [IntPtr]((12 -shl 16) -bor 43)
Start-Sleep -Milliseconds 800
$in = $false
for ($try = 1; $try -le 3 -and -not $in; $try++) {     # the first press after typing the year is sometimes ignored
  foreach ($m in 0x0200, 0x0201, 0x0202) { [UL]::PostMessage([IntPtr]$ok.Current.NativeWindowHandle, $m, [IntPtr][int]($m -eq 0x0201), $lp) | Out-Null }
  $in = WaitFor { (Title) -like 'Insurances Solution*' } ([int]($Wait / 3))
}
if (-not $in) { throw "PIAS did not log in" }
DismissUpdate 15
"PIAS ready"