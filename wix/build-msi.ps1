param(
  [string]$ProjectRoot,
  [string]$PublishDir,
  [string]$OutputDir,
  [string]$Version = "1.0.0"
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

function Find-WixToolPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ToolName
  )

  $command = Get-Command $ToolName -ErrorAction SilentlyContinue
  if ($command) {
    return $command.Source
  }

  $candidateRoots = @()
  if ($env:WIX) {
    $candidateRoots += $env:WIX
    $candidateRoots += (Join-Path $env:WIX 'bin')
  }

  $programFilesRoots = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ }
  foreach ($root in $programFilesRoots) {
    $candidateRoots += Get-ChildItem -Path $root -Filter 'WiX Toolset*' -Directory -ErrorAction SilentlyContinue |
      ForEach-Object { $_.FullName; Join-Path $_.FullName 'bin' }
  }

  foreach ($candidateRoot in $candidateRoots | Where-Object { $_ }) {
    $toolPath = Join-Path $candidateRoot $ToolName
    if (Test-Path $toolPath) {
      return (Resolve-Path $toolPath).Path
    }
  }

  return $null
}

$script:wixInstallAttempted = $false
function Ensure-WixTool {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ToolName
  )

  $path = Find-WixToolPath -ToolName $ToolName
  if ($path) {
    return $path
  }

  if (-not $script:wixInstallAttempted) {
    $script:wixInstallAttempted = $true
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) {
      throw "WiX tool '$ToolName' was not found and winget.exe is unavailable. Install the WiX Toolset manually."
    }

    Write-Host "WiX tool '$ToolName' not found. Installing WiX Toolset via winget..."
    $arguments = @('install', '--id', 'WiXToolset.WiXToolset', '--exact', '--accept-package-agreements', '--accept-source-agreements')
    $process = Start-Process -FilePath $winget.Source -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
    if ($process.ExitCode -ne 0) {
      throw "winget failed to install the WiX Toolset (exit code $($process.ExitCode))."
    }

    $path = Find-WixToolPath -ToolName $ToolName
    if ($path) {
      return $path
    }
  }

  throw "Unable to locate '$ToolName'. Install the WiX Toolset and ensure it is on PATH."
}

Push-Location $scriptRoot
try {
  if (-not $PSBoundParameters.ContainsKey('ProjectRoot')) {
    $ProjectRoot = (Resolve-Path (Join-Path $scriptRoot '..')).Path
  } else {
    $ProjectRoot = (Resolve-Path $ProjectRoot).Path
  }

  if (-not $PSBoundParameters.ContainsKey('PublishDir')) {
    $PublishDir = Join-Path $ProjectRoot 'publish'
  }

  if (-not (Test-Path $PublishDir)) {
    throw "Publish directory '$PublishDir' not found. Run 'dotnet publish -o ""$PublishDir""' first or pass -PublishDir."
  }

  $PublishDir = (Resolve-Path $PublishDir).Path

  if (-not $PSBoundParameters.ContainsKey('OutputDir')) {
    $OutputDir = Join-Path $scriptRoot 'bin'
  }

  New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
  $OutputDir = (Resolve-Path $OutputDir).Path

  $objDir = Join-Path $scriptRoot 'obj'
  New-Item -ItemType Directory -Force -Path $objDir | Out-Null

  # Where are WiX tools?
  $candle = Ensure-WixTool -ToolName 'candle.exe'
  $light  = Ensure-WixTool -ToolName 'light.exe'

  Write-Host "Using candle: $candle"
  Write-Host "Using light : $light"

  $wixObj = Join-Path $objDir 'Product.wixobj'
  $wixSource = Join-Path $scriptRoot 'Product.wxs'

  # Compile
  & $candle -nologo -arch x64 -dPublishDir="$PublishDir" -dProductVersion="$Version" -dWixUILicenseRtf="" `
    -out $wixObj $wixSource

  # Link (UI extension gives us the simple InstallDir dialog)
  $msiPath = Join-Path $OutputDir ("InstallReferee-$Version-x64.msi")
  & $light -nologo -ext WixUIExtension -cultures:en-us `
    -out $msiPath $wixObj

  Write-Host "Built MSI ->" $msiPath
}
finally {
  Pop-Location
}
