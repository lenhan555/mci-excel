# Thu vien dung chung cho cac lenh chup B1-NX (dot-source: . "$CapDir\lib.ps1")
# Khong dat bien $log, $f, $code, $res, $stop (bien cua watcher).
if (-not ('W32' -as [type])) {
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class W32 {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int hh, bool r);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
  public static List<string> List() {
    var o = new List<string>();
    EnumWindows((h, l) => { if (IsWindowVisible(h)) { var sb = new StringBuilder(512); GetWindowText(h, sb, 512);
      if (sb.Length > 0) { RECT r; GetWindowRect(h, out r); uint p; GetWindowThreadProcessId(h, out p);
        o.Add(h.ToInt64() + "|" + p + "|" + r.Left + "," + r.Top + "," + r.Right + "," + r.Bottom + "|" + sb.ToString()); } } return true; }, IntPtr.Zero);
    return o;
  }
  public static IntPtr Find(string part) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, l) => { if (IsWindowVisible(h)) { var sb = new StringBuilder(512); GetWindowText(h, sb, 512);
      if (sb.ToString().Contains(part)) { found = h; return false; } } return true; }, IntPtr.Zero);
    return found;
  }
}
'@
}
$CapDir = 'D:\Onedrive\MCI Excel\1_Sessions\B1_Nen_tang_Excel_Du_lieu\Homework\Cham_bai\_capture'
$PracDir = 'D:\Onedrive\MCI Excel\1_Sessions\B1_Nen_tang_Excel_Du_lieu\Practice'
function Retry([scriptblock]$sb) {
  for ($i = 0; $i -lt 60; $i++) { try { return & $sb } catch { if ($_.Exception.Message -match '0x8001010A|0x80010001|busy|RPC_E') { Start-Sleep -Milliseconds 500 } else { throw } } }
  throw 'Excel ban qua lau'
}
function Get-Xl { try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { return $null } }
function Place([string]$titlePart, [int]$x, [int]$y, [int]$w, [int]$h) {
  $hw = [W32]::Find($titlePart)
  if ($hw -eq [IntPtr]::Zero) { return "khong thay cua so '$titlePart'" }
  [void][W32]::ShowWindow($hw, 9); [void][W32]::MoveWindow($hw, $x, $y, $w, $h, $true)
  $r = New-Object W32+RECT; [void][W32]::GetWindowRect($hw, [ref]$r)
  "placed '$titlePart' -> $($r.Left),$($r.Top),$($r.Right),$($r.Bottom)"
}
function Wins() { [W32]::List() }
function SnapN([string]$name, [int]$n = 3) { for ($i = 1; $i -le $n; $i++) { Snap ("{0}_{1}.png" -f $name, $i); Start-Sleep -Milliseconds 350 } }
function Get-XlFor([string]$path) {
  $x = Get-Xl; if ($x) { return $x }
  try { $wbk = [Runtime.InteropServices.Marshal]::BindToMoniker($path); return $wbk.Application } catch { return $null }
}
