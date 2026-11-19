# Explorer-style folder picker that comes up in front of the browser window.
# Prints the chosen folder (UTF-8) or nothing when cancelled.
param([string]$Start = "")
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8

Add-Type -ReferencedAssemblies System.Windows.Forms -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class PickFolder {
    [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")] class FileOpenDialogRCW {}
    [ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IFileDialog {
        [PreserveSig] int Show(IntPtr parent);
        void SetFileTypes(uint c, IntPtr f); void SetFileTypeIndex(uint i); void GetFileTypeIndex(out uint i);
        void Advise(IntPtr e, out uint c); void Unadvise(uint c);
        void SetOptions(uint o); void GetOptions(out uint o);
        void SetDefaultFolder(IShellItem i); void SetFolder(IShellItem i); void GetFolder(out IShellItem i);
        void GetCurrentSelection(out IShellItem i); void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string n);
        void GetFileName(out IntPtr n); void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string t);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string t); void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string t);
        void GetResult(out IShellItem i);
    }
    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellItem {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv); void GetParent(out IShellItem i);
        void GetDisplayName(uint sigdn, [MarshalAs(UnmanagedType.LPWStr)] out string name);
    }
    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    static extern void SHCreateItemFromParsingName(string path, IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem item);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);

    public static void Front(IntPtr h) {
        // a background process may not take the foreground; one synthetic Alt tap lifts that lock
        keybd_event(0x12, 0, 0, UIntPtr.Zero); keybd_event(0x12, 0, 2, UIntPtr.Zero);
        SetForegroundWindow(h);
    }
    public static string Pick(IntPtr owner, string start) {
        IFileDialog d = (IFileDialog)new FileOpenDialogRCW();
        d.SetOptions(0x20 | 0x40 | 0x800);           // pick folders, file system only, path must exist
        d.SetTitle("Chọn thư mục lưu");
        if (!string.IsNullOrEmpty(start) && System.IO.Directory.Exists(start)) {
            try { IShellItem it; SHCreateItemFromParsingName(start, IntPtr.Zero, typeof(IShellItem).GUID, out it); d.SetFolder(it); } catch {}
        }
        if (d.Show(owner) != 0) return "";
        IShellItem r; d.GetResult(out r); string p; r.GetDisplayName(0x80058000, out p); return p;
    }
}
"@

$w = New-Object System.Windows.Forms.Form
$w.TopMost = $true; $w.ShowInTaskbar = $false; $w.Opacity = 0
$w.FormBorderStyle = 'None'; $w.StartPosition = 'CenterScreen'; $w.Size = New-Object System.Drawing.Size(1, 1)
$w.Show(); [PickFolder]::Front($w.Handle); $w.Activate(); [System.Windows.Forms.Application]::DoEvents()
try { $p = [PickFolder]::Pick($w.Handle, $Start) }
catch {
    $d = New-Object System.Windows.Forms.FolderBrowserDialog
    $d.SelectedPath = $Start; $d.ShowNewFolderButton = $true
    $p = if ($d.ShowDialog($w) -eq 'OK') { $d.SelectedPath } else { "" }
}
$w.Close()
if ($p) { $p }
