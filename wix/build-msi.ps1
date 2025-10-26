param(
  [string]$ProjectRoot,
  [string]$PublishDir,
  [string]$OutputDir,
  [string]$Version = "1.0.0"
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
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
    throw "Publish directory '$PublishDir' not found. Run 'dotnet publish -o \"$PublishDir\"' first or pass -PublishDir."
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
  $candle = (Get-Command candle.exe -ErrorAction Stop).Source
  $light  = (Get-Command light.exe  -ErrorAction Stop).Source

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
