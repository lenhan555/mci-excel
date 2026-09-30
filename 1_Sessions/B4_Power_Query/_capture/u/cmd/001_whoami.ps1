$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
"u watcher admin=$isAdmin pid=$PID"
whoami /groups | Select-String 'Mandatory Label'
"excel: " + ((Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) -join ',')
