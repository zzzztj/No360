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

## Windows notifications

The installer registers a hidden per-user notification helper for the next
Windows sign-in. It reads successful no360 block events from the Application
event log and shows short, silent Windows notifications. Blocks within five
seconds are grouped, and the latest alert replaces the previous no360 entry.
Old events are not replayed at sign-in. Windows notification settings and Do Not
Disturb still apply; blocking continues if notifications are disabled.

For the current session, start the helper without administrator privileges:

```powershell
Start-Process powershell.exe -WindowStyle Hidden -ArgumentList '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Program Files\no360\Notifications\Notify-Blocks.ps1"'
```

The service runs as LocalSystem; the helper runs as the signed-in user and has
no blocking privileges. It uses bounded event reads and disposes event-log
handles after every read. It does not collect an in-memory event history.
Disable its automatic launch in Windows Settings > Apps > Startup under
No360Notifications. This does not disable the service.

Check notification formatting and event reading on Windows with:

```powershell
powershell.exe -NoProfile -File tests/Test-Notifications.ps1
```

If you like it and it's useful to 1.4 billion people (lol) please drop me a tip

BTC bitcoin:BC1Q505QL4WP2SZ0JGA8Y689NZ960NX9JD37CYRX99
