. 'D:\Onedrive\MCI Excel\1_Sessions\B1_Nen_tang_Excel_Du_lieu\Homework\Cham_bai\_capture\lib.ps1'
$WorkDir = "$CapDir\work"
function Ensure-Xl {
  $x = $null
  if ($global:XL) { try { $null = $global:XL.Version; $x = $global:XL } catch { $x = $null } }
  if (-not $x) { $x = Get-Xl }
  if (-not $x) { throw 'Chua co Excel dang chay (mo Excel bang tay / computer_open_application)' }
  $global:XL = $x
  Retry { $x.Visible = $true; $x.DisplayAlerts = $false }
  return $x
}
function Open-Work([string]$name) {
  $x = Ensure-Xl
  foreach ($w in $x.Workbooks) { if ($w.Name -eq $name) { return $w } }
  $wb = Retry { $x.Workbooks.Open("$WorkDir\$name", 0, $false) }
  return $wb
}
function Close-All { $x = Get-Xl; if ($x) { foreach ($w in @($x.Workbooks)) { Retry { $w.Close($false) } } } }
# Dat cua so Excel (x,y,w,h pixel), an Ribbon, zoom 100
function Frame($wb, [string]$sheet, [string]$goto, [int]$x=30, [int]$y=40, [int]$w=1000, [int]$h=420, [bool]$ribbon=$false) {
  $xl = $wb.Application
  $wb.Activate()
  if ($sheet) { $wb.Worksheets.Item($sheet).Activate() }
  $xl.WindowState = -4143
  $xl.Left = $x*0.75; $xl.Top = $y*0.75; $xl.Width = $w*0.75; $xl.Height = $h*0.75
  $ActiveWindow = $xl.ActiveWindow
  $ActiveWindow.Zoom = 100
  try { $xl.ExecuteExcel4Macro("SHOW.TOOLBAR(`"Ribbon`",$($ribbon.ToString().ToUpper()))") | Out-Null } catch {}
  if ($goto) { $xl.Goto($wb.ActiveSheet.Range($goto), $true) | Out-Null }
}

if (-not ('DWM' -as [type])) {
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public static class DWM {
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h, int a, out RECT r, int sz);
  public static int[] Frame(IntPtr h) { RECT r; DwmGetWindowAttribute(h, 9, out r, 16); return new int[]{r.L,r.T,r.R,r.B}; }
}
'@
}
$Root = 'D:\Onedrive\MCI Excel\1_Sessions\B1_Nen_tang_Excel_Du_lieu\Homework\Cham_bai'
function Crop-Rect([string]$src,[string]$dst,[int]$l,[int]$t,[int]$r,[int]$b) {
  $bmp = [System.Drawing.Bitmap]::FromFile($src)
  $rect = New-Object System.Drawing.Rectangle $l,$t,($r-$l),($b-$t)
  $c = $bmp.Clone($rect, $bmp.PixelFormat); $bmp.Dispose()
  New-Item -ItemType Directory -Force -Path (Split-Path $dst) | Out-Null
  $c.Save($dst, [System.Drawing.Imaging.ImageFormat]::Png); $c.Dispose()
}
# Chup cua so (theo phan tieu de) -> png da cat
function SnapWindow([string]$title,[string]$dst) {
  $hw = [W32]::Find($title); if ($hw -eq [IntPtr]::Zero) { return "khong thay '$title'" }
  Snap 'tmp_full.png' | Out-Null
  $f = [DWM]::Frame($hw)
  Crop-Rect "$CapDir\out\tmp_full.png" $dst $f[0] $f[1] $f[2] $f[3]
  "saved $dst ($($f[2]-$f[0])x$($f[3]-$f[1]))"
}
function Shot([string]$student,[string]$name,[string]$wbName,[string]$sheet,[string]$goto,[string]$sel,[string[]]$red=@(),[string[]]$hide=@(),[int]$w=1000,[int]$h=420,[scriptblock]$pre=$null) {
  $wb = Open-Work $wbName; $xl = $wb.Application
  Frame $wb $sheet $null 30 40 $w $h
  $win = $xl.ActiveWindow; try { $win.View = 1 } catch {}
  try { $xl.FormulaBarHeight = 1 } catch {}
  $ws = $wb.ActiveSheet
  foreach ($c in $hide) { $ws.Range("$c`:$c").EntireColumn.Hidden = $true }
  $win.ScrollRow = 1; $win.ScrollColumn = 1
  if ($goto) { $xl.Goto($ws.Range($goto), $true) | Out-Null }
  if ($sel) { $ws.Range($sel).Select() | Out-Null }
  if ($pre) { & $pre }
  $shapes = @()
  foreach ($r in $red) { $rg = $ws.Range($r); $s = $ws.Shapes.AddShape(1,$rg.Left,$rg.Top,$rg.Width,$rg.Height); $s.Fill.Visible = 0; $s.Line.ForeColor.RGB = 255; $s.Line.Weight = 2.25; $shapes += $s }
  Start-Sleep -Milliseconds 700
  Snap 'tmp_full.png' | Out-Null
  $f = [DWM]::Frame([IntPtr]$xl.Hwnd)
  $dst = "$Root\$student\Images\$name.png"
  Crop-Rect "$CapDir\out\tmp_full.png" $dst $f[0] $f[1] $f[2] $f[3]
  foreach ($s in $shapes) { $s.Delete() | Out-Null }
  foreach ($c in $hide) { $ws.Range("$c`:$c").EntireColumn.Hidden = $false }
  "saved $dst  frame=$($f -join ',')"
}
function Wrap-Insight([string]$wbName,[string]$sheet) {
  $wb = Open-Work $wbName; $ws = $wb.Worksheets.Item($sheet); $ws.Columns.Item(1).ColumnWidth = 95; $ws.Range('A1:A5').WrapText = $true; $ws.Rows('1:5').AutoFit() | Out-Null
}
$ANS = 'B1_BTVN_MCI_Mart_T4_2025_Answer.xlsx'
$Q1 = 'B1_MCI_Mart_Q1_2025_Answer.xlsx'

function Kill-Dialog([string]$title) {
  for ($i=0;$i -lt 10;$i++) { $hw=[W32]::Find($title); if ($hw -eq [IntPtr]::Zero) { return }; Focus-Win $title; [System.Windows.Forms.SendKeys]::SendWait('{ESC}'); Start-Sleep -Milliseconds 500 }
}
function Mso-Async([string]$id) {
  $cmd = "`$x=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); `$x.CommandBars.ExecuteMso('$id')"
  Start-Process powershell -ArgumentList '-NoProfile','-WindowStyle','Hidden','-Command',$cmd
}
# Mo hop thoai qua ExecuteMso, (tuy chon) chuyen tab bang Ctrl+PgDn, chup hop thoai
function Dlg([string]$student,[string]$name,[string]$wbName,[string]$sheet,[string]$goto,[string]$sel,[string]$mso,[string]$title,[int]$tabs=0,[int]$downs=0) {
  $wb = Open-Work $wbName; $xl = $wb.Application
  Frame $wb $sheet $null 30 40 1000 560
  try { $xl.ActiveWindow.View = 1 } catch {}
  $ws = $wb.ActiveSheet; $xl.ActiveWindow.ScrollRow=1; $xl.ActiveWindow.ScrollColumn=1
  if ($goto) { $xl.Goto($ws.Range($goto), $true) | Out-Null }
  if ($sel) { $ws.Range($sel).Select() | Out-Null }
  Mso-Async $mso
  $hw=[IntPtr]::Zero
  for ($i=0;$i -lt 30;$i++) { Start-Sleep -Milliseconds 400; $hw=[W32]::Find($title); if ($hw -ne [IntPtr]::Zero) {break} }
  if ($hw -eq [IntPtr]::Zero) { return "KHONG mo duoc '$title'" }
  Start-Sleep -Milliseconds 800
  Focus-Win $title
  for ($k=0;$k -lt $tabs;$k++) { [System.Windows.Forms.SendKeys]::SendWait('^{PGDN}'); Start-Sleep -Milliseconds 400 }
  for ($k=0;$k -lt $downs;$k++) { [System.Windows.Forms.SendKeys]::SendWait('{DOWN}'); Start-Sleep -Milliseconds 250 }
  Start-Sleep -Milliseconds 700
  $r = SnapWindow $title "$Root\$student\Images\$name.png"
  Kill-Dialog $title
  $r
}

function Kill-Dialog([string]$title) {
  for ($i=0;$i -lt 10;$i++) { $hw=[W32]::Find($title); if ($hw -eq [IntPtr]::Zero) { return }; Focus-Win $title; [System.Windows.Forms.SendKeys]::SendWait('{ESC}'); Start-Sleep -Milliseconds 500 }
}
function Mso-Async([string]$id) {
  $cmd = "`$x=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); `$x.CommandBars.ExecuteMso('$id')"
  Start-Process powershell -ArgumentList '-NoProfile','-WindowStyle','Hidden','-Command',$cmd
}
function Dlg([string]$student,[string]$name,[string]$wbName,[string]$sheet,[string]$goto,[string]$sel,[string]$mso,[string]$title,[int]$tabs=0,[int]$downs=0) {
  $wb = Open-Work $wbName; $xl = $wb.Application
  Frame $wb $sheet $null 30 40 1000 560
  try { $xl.ActiveWindow.View = 1 } catch {}
  $ws = $wb.ActiveSheet; $xl.ActiveWindow.ScrollRow=1; $xl.ActiveWindow.ScrollColumn=1
  if ($goto) { $xl.Goto($ws.Range($goto), $true) | Out-Null }
  if ($sel) { $ws.Range($sel).Select() | Out-Null }
  Mso-Async $mso
  $hw=[IntPtr]::Zero
  for ($i=0;$i -lt 30;$i++) { Start-Sleep -Milliseconds 400; $hw=[W32]::Find($title); if ($hw -ne [IntPtr]::Zero) {break} }
  if ($hw -eq [IntPtr]::Zero) { return "KHONG mo duoc '$title'" }
  Start-Sleep -Milliseconds 800
  Focus-Win $title
  for ($k=0;$k -lt $tabs;$k++) { [System.Windows.Forms.SendKeys]::SendWait('^{PGDN}'); Start-Sleep -Milliseconds 400 }
  for ($k=0;$k -lt $downs;$k++) { [System.Windows.Forms.SendKeys]::SendWait('{DOWN}'); Start-Sleep -Milliseconds 250 }
  Start-Sleep -Milliseconds 700
  $r = SnapWindow $title "$Root\$student\Images\$name.png"
  Kill-Dialog $title
  $r
}

function Focus-Win([string]$title) { $sh = New-Object -ComObject WScript.Shell; [void]$sh.AppActivate($title); Start-Sleep -Milliseconds 400 }

# Mo hop thoai (khong gui phim): chi dat cua so + mo qua ExecuteMso; chup bang SnapDlg sau khi da bam tab bang computer use
function OpenDlg([string]$wbName,[string]$sheet,[string]$goto,[string]$sel,[string]$mso,[string]$title) {
  $wb = Open-Work $wbName; $xl = $wb.Application
  Frame $wb $sheet $null 30 40 1000 560
  try { $xl.ActiveWindow.View = 1 } catch {}
  $ws = $wb.ActiveSheet; $xl.ActiveWindow.ScrollRow=1; $xl.ActiveWindow.ScrollColumn=1
  if ($goto) { $xl.Goto($ws.Range($goto), $true) | Out-Null }
  if ($sel) { $ws.Range($sel).Select() | Out-Null }
  Mso-Async $mso
  $hw=[IntPtr]::Zero
  for ($i=0;$i -lt 30;$i++) { Start-Sleep -Milliseconds 400; $hw=[W32]::Find($title); if ($hw -ne [IntPtr]::Zero) {break} }
  if ($hw -eq [IntPtr]::Zero) { return "KHONG mo duoc '$title'" }
  "mo duoc '$title' cho $wbName $sel"
}
function SnapDlg([string]$student,[string]$name,[string]$title) { SnapWindow $title "$Root\$student\Images\$name.png" }

if (-not ('PM' -as [type])) {
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public static class PM { [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l); }
'@
}
function Close-Dlg([string]$title) { for ($i=0;$i -lt 10;$i++) { $hw=[W32]::Find($title); if ($hw -eq [IntPtr]::Zero) { return "closed" }; [void][PM]::PostMessage($hw,0x10,[IntPtr]::Zero,[IntPtr]::Zero); Start-Sleep -Milliseconds 500 }; "NOT closed" }
# Dong hop thoai cu (neu con), mo hop thoai moi, chup (tab giu nguyen nhu lan truoc)
function DlgShot([string]$student,[string]$name,[string]$wbName,[string]$sheet,[string]$goto,[string]$sel,[string]$title='Data Validation',[string]$mso='DataValidation') {
  Close-Dlg $title | Out-Null
  $r = OpenDlg $wbName $sheet $goto $sel $mso $title
  Start-Sleep -Milliseconds 1200
  $r; SnapDlg $student $name $title
}

function SnapXl([string]$student,[string]$name,[string]$wbName) {
  $wb = Open-Work $wbName; $xl = $wb.Application
  Snap 'tmp_full.png' | Out-Null
  $f = [DWM]::Frame([IntPtr]$xl.Hwnd)
  Crop-Rect "$CapDir\out\tmp_full.png" "$Root\$student\Images\$name.png" $f[0] $f[1] $f[2] $f[3]
  "saved $name frame=$($f -join ',')"
}
function SetupRibbon($wbName,$sheet,$goto,$sel,[int]$w=1000,[int]$h=600) {
  $wb = Open-Work $wbName; $xl = $wb.Application
  Frame $wb $sheet $null 30 40 $w $h $true
  try { $xl.ActiveWindow.View = 1 } catch {}
  $ws = $wb.ActiveSheet; $xl.ActiveWindow.ScrollRow=1; $xl.ActiveWindow.ScrollColumn=1
  if ($goto) { $xl.Goto($ws.Range($goto), $true) | Out-Null }
  if ($sel) { $ws.Range($sel).Select() | Out-Null }
  return $ws
}

function Close-Books { foreach ($w in @($global:XL.Workbooks)) { if ($w.Name -like 'Book*') { Retry { $w.Close($false) } } } }
function Front($wbName) { $wb = Open-Work $wbName; $wb.Activate(); $wb.Windows.Item(1).Activate(); $wb.Windows.Item(1).WindowState = -4143 }
