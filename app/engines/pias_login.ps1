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
function DismissUpdateTop() {
  # the same question can come up as PIAS's own main window (its title is the install folder)
  $p = Get-Process Pias2022 -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $p -or $p.MainWindowHandle -eq 0) { return $false }
  if ($p.MainWindowTitle -notlike 'C:\*') { return $false }
  $w = $AE::FromHandle($p.MainWindowHandle)
  $no = @($w.FindAll($TS::Descendants, $ALL) | Where-Object { $_.Current.Name -in 'No', 'Không' })[0]
  if (-not $no) { return $false }
  $lp = [IntPtr]((11 -shl 16) -bor 37)
  foreach ($m in 0x0200, 0x0201, 0x0202) { [UL]::PostMessage([IntPtr]$no.Current.NativeWindowHandle, $m, [IntPtr][int]($m -eq 0x0201), $lp) | Out-Null }
  "update prompt (main window): answered No"
  Start-Sleep -Milliseconds 900
  return $true
}
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
# the year this tool logged PIAS in with (the toolbar year box is painted, so it cannot be read back)
$YearFile = Join-Path (Split-Path $PSScriptRoot) 'dump\pias_year.txt'
function SessionYear() {
  try { if (Test-Path $YearFile) { return (Get-Content $YearFile -Raw).Trim() } } catch {}
  return ""
}
if ((Title) -like 'Insurances Solution*') {
  $cur = SessionYear
  if (-not $Year -or $cur -eq "$Year") { "PIAS ready (year $cur)"; exit 0 }
  "session year '$cur' is not $Year, restarting PIAS"
  Stop-Process -Name Pias2022 -Force
  Start-Sleep -Seconds 4
}
if (-not (Get-Process Pias2022 -ErrorAction SilentlyContinue)) {
  Start-Process $EXE -WorkingDirectory (Split-Path $EXE)
}
# the "a new version is on the server" box (its title is the install folder) can come up first and block
$t0 = Get-Date
$ready = $false
while (((Get-Date) - $t0).TotalSeconds -lt $Wait -and -not $ready) {
  DismissUpdate | Out-Null
  DismissUpdateTop | Out-Null
  $ti = Title
  if ($ti -eq 'Đăng nhập' -or $ti -like 'Insurances Solution*') { $ready = $true; break }
  Start-Sleep -Milliseconds 800
}
if (-not $ready) { throw "PIAS login window did not appear (it shows: " + (Title) + ")" }
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
if ($Year) {
  try {
    New-Item -ItemType Directory -Force (Split-Path $YearFile) | Out-Null
    Set-Content -Path $YearFile -Value "$Year" -Encoding ascii
  } catch {}
}
"PIAS ready"