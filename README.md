# Build Instruction

dotnet publish .\No360\No360.csproj `
  -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:PublishTrimmed=false `
  -o .\publish

# Wix Installer


# No360
As a manager of number of user owned PCs one of the things that I spend most of my time doing is removing Qihoo 360+. 

This is a little helper service that tries to prevent 360 from being installed as part of other bundles.

If you like it and it's useful to 1.4 billion people (lol) please drop me a tip

BTC bitcoin:BC1Q505QL4WP2SZ0JGA8Y689NZ960NX9JD37CYRX99
