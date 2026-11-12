param(
  [int]$Report,                 # number in the 'Báo cáo hiệu quả kinh doanh' list (2 = Kỹ thuật, 10 = Con người, ...)
  [string]$From, [string]$To,   # dd/MM/yyyy
  [switch]$RevenueOnly,         # untick Thực thu and Bồi thường in 'Chọn kiểu in'
  [int]$WaitDialog = 60,
  [string]$Snap = ""            # save a picture of 'Chọn kiểu in' before OK (diagnostics)
)
# Drive PIAS (Pias2022.exe, already logged in, 'Báo cáo hiệu quả kinh doanh' list open) to run one report.
# Clicks are posted to the window handles, so the user's mouse and keyboard are not touched.
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type @"
using System; using System.Runtime.InteropServices;
public class U { [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
[DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, string l);
[DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
[DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
[DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr h);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr p);
[DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
[DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool f);
[DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr h); }
"@
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$ALL = [System.Windows.Automation.Condition]::TrueCondition

function Click([IntPtr]$h, [int]$x, [int]$y) {
  $lp = [IntPtr](($y -shl 16) -bor ($x -band 0xFFFF))
  [U]::PostMessage($h, 0x0200, [IntPtr]0, $lp) | Out-Null
  [U]::PostMessage($h, 0x0201, [IntPtr]1, $lp) | Out-Null
  [U]::PostMessage($h, 0x0202, [IntPtr]0, $lp) | Out-Null
}
function SetText([IntPtr]$h, [string]$t) {
  [U]::SendMessage($h, 0x000C, [IntPtr]0, $t) | Out-Null
  $id = [U]::GetDlgCtrlID($h)
  [U]::SendMessage([U]::GetParent($h), 0x0111, [IntPtr](($id -band 0xFFFF) -bor (0x0300 -shl 16)), $h) | Out-Null
}
function TypeDate([IntPtr]$h, [string]$t) {
  # Setting the text only changes what is shown; PIAS reads the date box's own value, which follows typing.
  # Focus is given inside PIAS's thread (AttachThreadInput + SetFocus): no foreground change, user's keyboard untouched.
  $tid = [U]::GetWindowThreadProcessId($h, [IntPtr]::Zero)
  $me = [U]::GetCurrentThreadId()
  [U]::AttachThreadInput($me, $tid, $true) | Out-Null
  try {
    [U]::SetFocus($h) | Out-Null
    Start-Sleep -Milliseconds 200
    [U]::PostMessage($h, 0x0100, [IntPtr]0x24, [IntPtr]0) | Out-Null      # Home
    [U]::PostMessage($h, 0x0101, [IntPtr]0x24, [IntPtr]0) | Out-Null
    foreach ($c in ($t -replace '\D', '').ToCharArray()) {                # dd MM yyyy, the mask skips the '/'
      [U]::PostMessage($h, 0x0102, [IntPtr][int]$c, [IntPtr]0) | Out-Null
      Start-Sleep -Milliseconds 80
    }
    Start-Sleep -Milliseconds 400
  } finally { [U]::AttachThreadInput($me, $tid, $false) | Out-Null }
}
function Main() { $p = Get-Process Pias2022 -ErrorAction Stop | Select-Object -First 1; $AE::FromHandle($p.MainWindowHandle) }
function Child([string]$prefix, [int]$timeout) {
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $timeout) {
    foreach ($w in (Main).FindAll($TS::Children, $ALL)) { if ($w.Current.Name -like "$prefix*") { return $w } }
    Start-Sleep -Milliseconds 500
  }
  throw "window '$prefix' not found"
}
function Find($root, [scriptblock]$pred) { @($root.FindAll($TS::Descendants, $ALL) | Where-Object $pred) }

# 1. open the report from the list (NavBar items are painted, so click by row: 22 px per item, first at y=38)
$open = @((Main).FindAll($TS::Children, $ALL) | Where-Object { $_.Current.Name -like "$Report. *" })
if (-not $open) {
  $nav = Find (Main) { $_.Current.Name -eq 'NavBarControl1' -and $_.Current.BoundingRectangle.Width -gt 1000 } | Select-Object -First 1
  Click ([IntPtr]$nav.Current.NativeWindowHandle) 192 (38 + 22 * ($Report - 1))
}
$dlg = Child "$Report. " $WaitDialog
"opened: $($dlg.Current.Name)"

# 2. period: the two date edits, left = Từ ngày, right = Đến ngày
#    the dialog fills its own default dates while loading, so wait, set, then read back (and retry)
Start-Sleep -Seconds 3
$want = "$From - $To"
for ($try = 1; $try -le 3; $try++) {
  $edits = Find $dlg { $_.Current.ClassName -like '*EDIT*' } | Sort-Object { $_.Current.BoundingRectangle.X }
  foreach ($pair in @(@($edits[0], $From), @($edits[1], $To))) {
    $e = $pair[0]; $h = [IntPtr]$e.Current.NativeWindowHandle
    TypeDate $h $pair[1]
    Start-Sleep -Milliseconds 400
  }
  Start-Sleep -Seconds 2
  $got = (Find $dlg { $_.Current.ClassName -like '*EDIT*' } | Sort-Object { $_.Current.BoundingRectangle.X } |
          ForEach-Object { $_.Current.Name }) -join ' - '
  if ($got -eq $want) { break }
}
"period: $got"
if ($got -ne $want) { throw "period not set (got $got)" }
$ok = Find $dlg { $_.Current.Name -eq 'Chấp nhận' } | Select-Object -First 1
Click ([IntPtr]$ok.Current.NativeWindowHandle) 49 11

# 3. 'Chọn kiểu in': row 2 = Doanh thu, Thực thu, Bồi thường
$kind = Child 'Chọn kiểu in' 600
if ($RevenueOnly) {
  $boxes = Find $kind { $_.Current.Name -eq '' -and $_.Current.ClassName -like '*Window.b.*' } |
           Sort-Object { $_.Current.BoundingRectangle.Y }, { $_.Current.BoundingRectangle.X }
  $row2 = $boxes | Where-Object { $_.Current.BoundingRectangle.Y -gt $boxes[0].Current.BoundingRectangle.Y + 5 }
  foreach ($b in ($row2 | Select-Object -Skip 1)) { Click ([IntPtr]$b.Current.NativeWindowHandle) 8 9 }
  Start-Sleep -Milliseconds 600
}
if ($Snap) {                     # diagnostics: picture of 'Chọn kiểu in' as it is sent
  Add-Type -AssemblyName System.Drawing
  Add-Type @"
using System; using System.Runtime.InteropServices;
public class PW { [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f); }
"@
  $r = $kind.Current.BoundingRectangle
  $bmp = New-Object System.Drawing.Bitmap ([int]$r.Width), ([int]$r.Height)
  $g = [System.Drawing.Graphics]::FromImage($bmp); $dc = $g.GetHdc()
  [PW]::PrintWindow([IntPtr]$kind.Current.NativeWindowHandle, $dc, 2) | Out-Null
  $g.ReleaseHdc($dc); $g.Dispose(); $bmp.Save($Snap); "snap: $Snap"
}
$okk = Find $kind { $_.Current.Name -eq 'OK' } | Select-Object -First 1
Click ([IntPtr]$okk.Current.NativeWindowHandle) 41 11
"started at $(Get-Date -Format HH:mm:ss)"
