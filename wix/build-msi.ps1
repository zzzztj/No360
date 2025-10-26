param(
  [string]$ProjectRoot = (Resolve-Path "..\..").Path,   # adjust if needed
  [string]$PublishDir  = (Resolve-Path "..\..\publish").Path,
  [string]$OutputDir   = (Resolve-Path ".\bin").Path,
  [string]$Version     = "1.0.0"
)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
New-Item -ItemType Directory -Force -Path ".\obj" | Out-Null

# Where are WiX tools?
$candle = (Get-Command candle.exe -ErrorAction Stop).Source
$light  = (Get-Command light.exe  -ErrorAction Stop).Source

Write-Host "Using candle: $candle"
Write-Host "Using light : $light"

# Compile
& $candle -nologo -arch x64 -dPublishDir="$PublishDir" -dProductVersion="$Version" -dWixUILicenseRtf="" `
  -out ".\obj\Product.wixobj" ".\Product.wxs"

# Link (UI extension gives us the simple InstallDir dialog)
& $light -nologo -ext WixUIExtension -cultures:en-us `
  -out (Join-Path $OutputDir ("InstallReferee-$Version-x64.msi")) ".\obj\Product.wixobj"

Write-Host "Built MSI ->" (Join-Path $OutputDir ("InstallReferee-$Version-x64.msi"))
