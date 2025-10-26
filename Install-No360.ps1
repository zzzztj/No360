param(
  [string]$ServiceName = "InstallReferee",
  [string]$DisplayName = "Install Referee",
  [string]$Description = "Blocks installation of disallowed software (bundle-aware; EXE/MSI).",
  [string]$TargetDir = "C:\Program Files\InstallReferee"
)

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
Copy-Item -Path "$here\*" -Destination $TargetDir -Recurse -Force

$exe = Join-Path $TargetDir "InstallReferee.exe"

if (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue) {
  sc.exe stop $ServiceName | Out-Null
  sc.exe delete $ServiceName | Out-Null
  Start-Sleep -Seconds 1
}
sc.exe create $ServiceName binPath= "`"$exe`"" start= auto DisplayName= "`"$DisplayName`"" | Out-Null
sc.exe description $ServiceName "$Description" | Out-Null
Start-Service $ServiceName
Write-Host "Service '$ServiceName' installed and started."
