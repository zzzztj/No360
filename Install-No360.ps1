param(
  [string]$ServiceName = "no360",
  [string]$DisplayName = "no360",
  [string]$Description = "Blocks installation of disallowed software (bundle-aware; EXE/MSI).",
  [string]$TargetDir = "C:\Program Files\no360"
)

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
Copy-Item -Path "$here\*" -Destination $TargetDir -Recurse -Force

$exe = Join-Path $TargetDir "no360.exe"

if (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue) {
  sc.exe stop $ServiceName | Out-Null
  sc.exe delete $ServiceName | Out-Null
  Start-Sleep -Seconds 1
}
sc.exe create $ServiceName binPath= "`"$exe`"" start= auto DisplayName= "`"$DisplayName`"" | Out-Null
sc.exe description $ServiceName "$Description" | Out-Null
Start-Service $ServiceName
$notificationScript = Join-Path $TargetDir 'Notifications\Notify-Blocks.ps1'
if (Test-Path -LiteralPath $notificationScript) {
  $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
  $command = '"' + $powershell + '" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $notificationScript + '"'
  New-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'No360Notifications' -Value $command -PropertyType String -Force | Out-Null
}
Write-Host "Service '$ServiceName' installed and started."
