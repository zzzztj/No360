[CmdletBinding(SupportsShouldProcess)]
param(
  [string]$ServiceName = "no360",
  [string]$DisplayName = "no360",
  [string]$Description = "Blocks installation of disallowed software (bundle-aware; EXE/MSI).",
  [string]$TargetDir = "C:\Program Files\no360"
)

$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$TargetDir = [IO.Path]::GetFullPath($TargetDir).TrimEnd('\')
if ($TargetDir -eq $here.TrimEnd('\')) { throw 'Extract the release to a separate folder before installing.' }
if ($TargetDir.StartsWith($here.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'The installation folder must be outside the extracted release folder.' }
if (-not (Test-Path -LiteralPath (Join-Path $here 'no360.exe'))) { throw 'Run this script from the extracted binary release, alongside no360.exe.' }
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($service) {
  $serviceInfo = Get-CimInstance Win32_Service | Where-Object Name -eq $ServiceName
  if ($serviceInfo.PathName.Trim('"') -ne (Join-Path $TargetDir 'no360.exe')) {
    throw 'The existing service uses a different executable path. Specify its installation folder with -TargetDir.'
  }
}
if (-not $PSCmdlet.ShouldProcess($TargetDir, "Install or upgrade $ServiceName, preserving existing settings")) { return }
$principal = [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run this installer from an administrator PowerShell window.' }
$configPath = Join-Path $TargetDir 'appsettings.json'
$backup = Join-Path $env:ProgramData ('no360\Backups\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N'))
$wasRunning = $service -and $service.Status -eq 'Running'
$copied = $false
$hadInstallation = Test-Path -LiteralPath $TargetDir
if ($hadInstallation) {
  New-Item -ItemType Directory -Path $backup -Force | Out-Null
  Get-ChildItem -LiteralPath $TargetDir -Force | Copy-Item -Destination $backup -Recurse -Force
  Write-Host "Backup: $backup"
}
try {
  if ($service) {
    Stop-Service -Name $ServiceName
    (Get-Service $ServiceName).WaitForStatus('Stopped', [TimeSpan]::FromSeconds(30))
  }
  New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
  $preserveConfig = Test-Path -LiteralPath $configPath
  $copied = $true
  Get-ChildItem -LiteralPath $here | Where-Object { $_.Name -ne 'appsettings.json' } |
    Copy-Item -Destination $TargetDir -Recurse -Force
  if (-not $preserveConfig) {
    $settings = Get-Content -LiteralPath (Join-Path $here 'appsettings.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    # Resolve folders here, before LocalSystem expands them to its own profile.
    $settings.WatchFolders = @($settings.WatchFolders | ForEach-Object {
      [Environment]::ExpandEnvironmentVariables($_)
    } | Select-Object -Unique)
    $settings | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $configPath -Encoding UTF8
  }
  if (-not $service) {
    New-Service -Name $ServiceName -DisplayName $DisplayName -Description $Description `
      -BinaryPathName ('"' + (Join-Path $TargetDir 'no360.exe') + '"') -StartupType Automatic | Out-Null
  }
  Start-Service -Name $ServiceName
  (Get-Service $ServiceName).WaitForStatus('Running', [TimeSpan]::FromSeconds(30))
  Start-Sleep -Seconds 5
  if ((Get-Service $ServiceName).Status -ne 'Running') { throw 'Service did not stay running.' }
} catch {
  if ($copied -and $hadInstallation) {
    Stop-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (Get-Service $ServiceName -ErrorAction SilentlyContinue) {
      (Get-Service $ServiceName).WaitForStatus('Stopped', [TimeSpan]::FromSeconds(30))
    }
    Get-ChildItem -LiteralPath $backup -Force | Copy-Item -Destination $TargetDir -Recurse -Force
  }
  if ($wasRunning) { Start-Service -Name $ServiceName }
  throw
}
$notificationScript = Join-Path $TargetDir 'Notifications\Notify-Blocks.ps1'
if (Test-Path -LiteralPath $notificationScript) {
  $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
  $command = '"' + $powershell + '" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $notificationScript + '"'
  New-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'No360Notifications' -Value $command -PropertyType String -Force | Out-Null
}
Write-Host "Service '$ServiceName' installed and started."
Write-Host 'Notifications start at next sign-in. For this session, start the helper without administrator privileges as described in INSTALL.txt.'
