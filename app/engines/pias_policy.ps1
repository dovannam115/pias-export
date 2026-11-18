param(
  [string]$From, [string]$To,          # dd/MM/yyyy
  [string]$Product = "",               # 2-digit line, e.g. 01 (Con người)
  [string]$Save,                       # full path of the file to write
  [int]$SearchWait = 900,              # seconds to wait for the search to come back
  [int]$SaveWait = 120
)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# Drive the PIAS screen "Tìm kiếm đơn Bảo hiểm": set the period and the line, search, then Xuất Exel and
# answer the Save As box with $Save. Clicks are posted to window handles and the dates are typed into the
# focused box, so the user's mouse and keyboard are not touched.
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class PS2 {
 [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, string l);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr p);
 [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
 [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool f);
 [DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr h);
 [DllImport("user32.dll")] public static extern bool EnumWindows(Cb c, IntPtr l);
 [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, Cb c, IntPtr l);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
 [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
 public delegate bool Cb(IntPtr h, IntPtr l); }
"@
$AE = [System.Windows.Automation.AutomationElement]; $TS = [System.Windows.Automation.TreeScope]
$ALL = [System.Windows.Automation.Condition]::TrueCondition
$PIAS = Get-Process Pias2022 -ErrorAction Stop | Select-Object -First 1

function Screen() {
  $root = $AE::FromHandle($PIAS.MainWindowHandle)
  foreach ($w in $root.FindAll($TS::Children, $ALL)) { if ($w.Current.Name -eq 'Tìm kiếm đơn Bảo hiểm') { return $w } }
  throw "the policy search screen is not open"
}
function Click([IntPtr]$h, [int]$x = 20, [int]$y = 8) {
  $lp = [IntPtr](($y -shl 16) -bor ($x -band 0xFFFF))
  foreach ($m in 0x0200, 0x0201, 0x0202) { [PS2]::PostMessage($h, $m, [IntPtr][int]($m -eq 0x0201), $lp) | Out-Null }
}
function TypeIn([IntPtr]$h, [string]$t, [switch]$Replace) {
  $tid = [PS2]::GetWindowThreadProcessId($h, [IntPtr]::Zero); $me = [PS2]::GetCurrentThreadId()
  [PS2]::AttachThreadInput($me, $tid, $true) | Out-Null
  try {
    [PS2]::SetFocus($h) | Out-Null; Start-Sleep -Milliseconds 200
    if ($Replace) {
      # Home then Delete: EM_SETSEL(-1) makes the DevExpress masked editors overflow and crash PIAS
      [PS2]::PostMessage($h, 0x0100, [IntPtr]0x24, [IntPtr]0) | Out-Null
      [PS2]::PostMessage($h, 0x0101, [IntPtr]0x24, [IntPtr]0) | Out-Null
      for ($i = 0; $i -lt 14; $i++) {
        [PS2]::PostMessage($h, 0x0100, [IntPtr]0x2E, [IntPtr]0) | Out-Null
        [PS2]::PostMessage($h, 0x0101, [IntPtr]0x2E, [IntPtr]0) | Out-Null
        Start-Sleep -Milliseconds 30
      }
      Start-Sleep -Milliseconds 150
    }
    [PS2]::PostMessage($h, 0x0100, [IntPtr]0x24, [IntPtr]0) | Out-Null        # Home
    [PS2]::PostMessage($h, 0x0101, [IntPtr]0x24, [IntPtr]0) | Out-Null
    foreach ($c in ($t -replace '\D', '').ToCharArray()) {
      [PS2]::PostMessage($h, 0x0102, [IntPtr][int]$c, [IntPtr]0) | Out-Null
      Start-Sleep -Milliseconds 70
    }
    Start-Sleep -Milliseconds 300
  } finally { [PS2]::AttachThreadInput($me, $tid, $false) | Out-Null }
}
function Boxes() { @((Screen).FindAll($TS::Descendants, $ALL) | Where-Object { $_.Current.ClassName -like '*EDIT*' }) }
function DateBoxes() {
  # once the grid holds rows it grows date boxes of its own, so keep only the pair on the top row
  $all = @(Boxes | Where-Object { $_.Current.Name -match '^\d\d/\d\d/\d{4}$' })
  if (-not $all) { return @() }
  $top = ($all | ForEach-Object { $_.Current.BoundingRectangle.Y } | Measure-Object -Minimum).Minimum
  @($all | Where-Object { $_.Current.BoundingRectangle.Y -lt $top + 6 } |
    Sort-Object { $_.Current.BoundingRectangle.X })
}
function FieldRight([string]$label) {
  $all = @((Screen).FindAll($TS::Descendants, $ALL))
  $lab = @($all | Where-Object { $_.Current.Name -eq $label -and $_.Current.ClassName -like '*STATIC*' })[0]
  if (-not $lab) { throw "field '$label' not found" }
  $y = $lab.Current.BoundingRectangle.Y; $x = $lab.Current.BoundingRectangle.X
  $b = @($all | Where-Object { $_.Current.ClassName -like '*EDIT*' -and
          [Math]::Abs($_.Current.BoundingRectangle.Y - $y) -lt 8 -and $_.Current.BoundingRectangle.X -gt $x } |
        Sort-Object { $_.Current.BoundingRectangle.X })[0]
  if (-not $b) { throw "no box next to '$label'" }
  return $b
}
function Button([string]$name) {
  $b = @((Screen).FindAll($TS::Descendants, $ALL) | Where-Object { $_.Current.Name -eq $name })[0]
  if (-not $b) { throw "button '$name' not found" }
  return $b
}
function SaveDialog() {
  $found = [IntPtr]::Zero
  $cb = [PS2+Cb]{
    param($h, $l)
    $pid2 = 0; [PS2]::GetWindowThreadProcessId($h, [ref]$pid2) | Out-Null
    if ($pid2 -eq $PIAS.Id -and [PS2]::IsWindowVisible($h)) {
      $c = New-Object System.Text.StringBuilder 64
      [PS2]::GetClassNameW($h, $c, 64) | Out-Null
      if ($c.ToString() -eq '#32770') { $script:found = $h; return $false }
    }
    return $true
  }
  [PS2]::EnumWindows($cb, [IntPtr]::Zero) | Out-Null
  return $script:found
}
function DismissBoxes() {
  # PIAS says "xuất dữ liệu thành công" in a box of its own; left alone it blocks the whole screen
  $said = @()
  for ($i = 1; $i -le 8; $i++) {
    $root = $AE::FromHandle($PIAS.MainWindowHandle)
    $box = @($root.FindAll($TS::Children, $ALL) | Where-Object { $_.Current.Name -eq 'Insurace Solution 2022' })[0]
    if (-not $box) { break }
    $els = @($box.FindAll($TS::Descendants, $ALL))
    $ok = $els | Where-Object { $_.Current.Name -eq 'OK' } | Select-Object -First 1
    if (-not $ok) { break }
    $txt = ($els | Where-Object { $_.Current.Name -and $_.Current.Name -notin 'OK', 'Insurace Solution 2022' } |
            ForEach-Object { $_.Current.Name } | Select-Object -Unique) -join ' '
    $said += $txt
    $lp = [IntPtr]((11 -shl 16) -bor 37)
    foreach ($m in 0x0200, 0x0201, 0x0202) { [PS2]::PostMessage([IntPtr]$ok.Current.NativeWindowHandle, $m, [IntPtr][int]($m -eq 0x0201), $lp) | Out-Null }
    Start-Sleep -Milliseconds 800
  }
  return $said
}

function FillSave([IntPtr]$dlg, [string]$path) {
  $kids = New-Object System.Collections.ArrayList
  $cb = [PS2+Cb]{ param($h, $l) [void]$kids.Add($h); return $true }
  [PS2]::EnumChildWindows($dlg, $cb, [IntPtr]::Zero) | Out-Null
  $edit = [IntPtr]::Zero; $save = [IntPtr]::Zero
  foreach ($h in $kids) {
    $c = New-Object System.Text.StringBuilder 64; [PS2]::GetClassNameW($h, $c, 64) | Out-Null
    $t = New-Object System.Text.StringBuilder 200; [PS2]::GetWindowTextW($h, $t, 200) | Out-Null
    if ($c.ToString() -eq 'Edit' -and $edit -eq [IntPtr]::Zero) { $edit = $h }
    if ($c.ToString() -eq 'Button' -and ($t.ToString() -replace '&', '') -in 'Save', 'Lưu', 'OK') { $save = $h }
  }
  if ($edit -eq [IntPtr]::Zero) { throw "the Save As box has no file name field" }
  [PS2]::SendMessage($edit, 0x000C, [IntPtr]0, $path) | Out-Null       # WM_SETTEXT
  Start-Sleep -Milliseconds 400
  if ($save -ne [IntPtr]::Zero) { [PS2]::PostMessage($save, 0x00F5, [IntPtr]0, [IntPtr]0) | Out-Null }
  else { [PS2]::PostMessage($dlg, 0x0111, [IntPtr]1, [IntPtr]0) | Out-Null }   # IDOK
}

# ---------------------------------------------------------------- 1. period and line
$now = ""
for ($try = 1; $try -le 3; $try++) {
  $d = DateBoxes
  if ($d.Count -lt 2) { throw "the two date boxes were not found" }
  TypeIn ([IntPtr]$d[0].Current.NativeWindowHandle) $From
  TypeIn ([IntPtr]$d[1].Current.NativeWindowHandle) $To
  Start-Sleep -Milliseconds 800
  $now = (DateBoxes | ForEach-Object { $_.Current.Name }) -join ' - '
  if ($now -eq "$From - $To") { break }
}
"period: $now"
if ($now -ne "$From - $To") { throw "period not set (got $now)" }

if ($Product) {
  $pb = FieldRight 'Sản phẩm'
  for ($try = 1; $try -le 3; $try++) {
    TypeIn ([IntPtr](FieldRight 'Sản phẩm').Current.NativeWindowHandle) $Product -Replace
    Start-Sleep -Milliseconds 500
    $got = (FieldRight 'Sản phẩm').Current.Name
    if ($got -eq $Product) { break }
  }
  "line: $got"
  if ($got -ne $Product) { throw "line not set (got '$got')" }
}

# ---------------------------------------------------------------- 2. search
Click ([IntPtr](Button 'Tìm kiếm').Current.NativeWindowHandle)
"searching"
$t0 = Get-Date
do {
  Start-Sleep -Seconds 2
  $busy = $false
  foreach ($w in $AE::RootElement.FindAll($TS::Children, $ALL)) {
    if ($w.Current.ProcessId -eq $PIAS.Id -and $w.Current.Name -eq 'Form1') { $busy = $true }
  }
} while ($busy -and ((Get-Date) - $t0).TotalSeconds -lt $SearchWait)
"search took $([int]((Get-Date) - $t0).TotalSeconds)s"
Start-Sleep -Seconds 2

# ---------------------------------------------------------------- 3. export and answer Save As
Click ([IntPtr](Button 'Xuất Exel').Current.NativeWindowHandle)
"export clicked"
$t0 = Get-Date
do {
  Start-Sleep -Milliseconds 700
  $dlg = SaveDialog
} while ($dlg -eq [IntPtr]::Zero -and ((Get-Date) - $t0).TotalSeconds -lt $SaveWait)
if ($dlg -eq [IntPtr]::Zero) { throw "PIAS did not ask where to save (no rows to export?)" }
FillSave $dlg $Save
"saving to $Save"

