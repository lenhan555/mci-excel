# Dot-source sau lib.ps1. SnapClean 'S05a' 3 'Power Pivot for Excel' : thu nho cua so Claude, chup N anh toan man hinh, khoi phuc, dua cua so can thao tac len truoc.
if (-not ('KB' -as [type])) { Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class KB { [DllImport("user32.dll")] public static extern void keybd_event(byte b, byte s, uint f, UIntPtr e); }' }
function Get-ClaudeHwnd { foreach ($l in (Wins)) { $p = $l -split '\|', 4; if ($p[3] -eq 'Claude') { return [IntPtr][int64]$p[0] } } return [IntPtr]::Zero }
function Front([string]$part) {
  $h = [W32]::Find($part)
  if ($h -eq [IntPtr]::Zero) { return "khong thay '$part'" }
  [KB]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero)          # ALT xuong: cho phep SetForegroundWindow
  [void][W32]::ShowWindow($h, 9); [void][W32]::SetForegroundWindow($h)
  [KB]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)          # ALT len
  Start-Sleep -Milliseconds 300
  if (-not $NoEsc) { [KB]::keybd_event(0x1B, 0, 0, [UIntPtr]::Zero); [KB]::keybd_event(0x1B, 0, 2, [UIntPtr]::Zero) }   # Esc: tat KeyTips
  "front '$part'"
}
function SnapClean([string]$name, [int]$n = 3, [string]$front = '') {
  $h = Get-ClaudeHwnd
  if ($h -ne [IntPtr]::Zero) { [void][W32]::ShowWindow($h, 6) }   # minimize
  Start-Sleep -Milliseconds 900
  SnapN $name $n
  if ($h -ne [IntPtr]::Zero -and -not $KeepClaudeMin) { [void][W32]::ShowWindow($h, 4) }   # restore, khong gianh focus (tru khi $KeepClaudeMin)
  if ($front) { Start-Sleep -Milliseconds 400; Front $front }
}

function ClaudeMin { $h = Get-ClaudeHwnd; if ($h -ne [IntPtr]::Zero) { [void][W32]::ShowWindow($h, 6); 'claude minimized' } }
function ClaudeRestore { $h = Get-ClaudeHwnd; if ($h -ne [IntPtr]::Zero) { [void][W32]::ShowWindow($h, 9); 'claude restored' } }
