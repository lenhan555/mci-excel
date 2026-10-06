# MCI Excel B5 - capture helper (tam thoi). Chay cac lenh dat trong .\cmd\*.ps1, ghi log vao .\out\.
# Tu dung sau 4 gio hoac khi gap lenh co '#STOP'.
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public static class U32 {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
}
'@
[void][U32]::SetProcessDPIAware()
$Base = Split-Path -Parent $MyInvocation.MyCommand.Path
$CmdDir = Join-Path $Base 'cmd'; $OutDir = Join-Path $Base 'out'
New-Item -ItemType Directory -Force -Path $CmdDir, $OutDir | Out-Null
function Snap([string]$name) {
  $b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
  $bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($b.X, $b.Y, 0, 0, $b.Size)
  $p = Join-Path $OutDir $name
  $bmp.Save($p, [System.Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bmp.Dispose()
  "snap $p $($b.Width)x$($b.Height)"
}
$deadline = (Get-Date).AddHours(4)
"started $(Get-Date -Format o) pid=$PID" | Out-File (Join-Path $OutDir 'watcher.log') -Append -Encoding UTF8
$stop = $false
while (-not $stop -and (Get-Date) -lt $deadline) {
  $f = Get-ChildItem -Path $CmdDir -Filter '*.ps1' -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object -First 1
  if ($f) {
    Start-Sleep -Milliseconds 300
    $code = [IO.File]::ReadAllText($f.FullName, [Text.Encoding]::UTF8)
    Remove-Item $f.FullName -Force
    $log = Join-Path $OutDir ($f.BaseName + '.log')
    try { $res = . ([ScriptBlock]::Create($code)) 2>&1 | Out-String -Width 400 }
    catch { $res = "ERROR: " + ($_ | Out-String) }
    [IO.File]::WriteAllText($log, "$res`nDONE", [Text.Encoding]::UTF8)
    if ($code -match '#STOP') { $stop = $true }
  }
  Start-Sleep -Milliseconds 250
}
"stopped $(Get-Date -Format o)" | Out-File (Join-Path $OutDir 'watcher.log') -Append -Encoding UTF8
