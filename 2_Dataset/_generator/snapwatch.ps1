# Chup RIENG cua so Excel bang PrintWindow (khong dinh panel khac). Ten file chua "full_" thi chup ca man hinh. Ghi ten png vao _snap.trigger; ghi STOP de thoat.
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public struct RECT { public int L, T, R, B; }
public class W {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
}
"@
[void][W]::SetProcessDPIAware()
$root = "D:\Onedrive\MCI Excel"
$trg = "$root\2_Dataset\_generator\vba\_snap.trigger"
$out = "$root\Materials\B6\Images\_raw"
New-Item -ItemType Directory -Force -Path $out | Out-Null
while ($true) {
  if (Test-Path $trg) {
    $name = (Get-Content $trg -Raw).Trim()
    Remove-Item $trg -Force
    if ($name -eq "STOP") { break }
    Start-Sleep -Milliseconds 400
    if ($name -like "*full_*") {
      $b = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize
      $bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height
      $g = [System.Drawing.Graphics]::FromImage($bmp)
      $g.CopyFromScreen(0, 0, 0, 0, $bmp.Size)
      $bmp.Save("$out\$name", [System.Drawing.Imaging.ImageFormat]::Png)
      $g.Dispose(); $bmp.Dispose()
      Set-Content "$root\2_Dataset\_generator\vba\_snap.done" "$name $($b.Width) 0 0 $($b.Width) $($b.Height)"
      continue
    }
    $p = Get-Process excel | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
    $h = $p.MainWindowHandle
    $r = New-Object RECT
    [void][W]::GetWindowRect($h, [ref]$r)
    $w = $r.R - $r.L; $ht = $r.B - $r.T
    $bmp = New-Object System.Drawing.Bitmap $w, $ht
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    [void][W]::PrintWindow($h, $hdc, 2)
    $g.ReleaseHdc($hdc)
    $bmp.Save("$out\$name", [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    $sw = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width
    Set-Content "$root\2_Dataset\_generator\vba\_snap.done" "$name $sw $($r.L) $($r.T) $w $ht"
  }
  Start-Sleep -Milliseconds 300
}
