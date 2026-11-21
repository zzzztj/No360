param(
    [switch]$LibraryOnly,
    [string]$TestMessage
)

$ErrorActionPreference = 'Stop'
$appId = 'No360.BlockNotifications'

function New-BlockToastXml([string]$Message, [int]$Count = 1) {
    # Treat event text strictly as text, including filenames containing XML characters.
    $text = ($Message -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ' ').Trim()
    if ($text.Length -gt 180) { $text = $text.Substring(0, 177) + '...' }
    if ($Count -gt 1) { $text = "$Count items blocked. Latest: $text" }
    $escaped = [Security.SecurityElement]::Escape($text)
    return "<toast duration=`"short`"><visual><binding template=`"ToastGeneric`"><text>no360 blocked software</text><text>$escaped</text></binding></visual><audio silent=`"true`"/></toast>"
}

function Read-BlockEvents([long]$AfterRecordId) {
    $xpath = "*[System[Provider[@Name='no360'] and EventID=3600 and EventRecordID &gt; $AfterRecordId]]".Replace('&gt;', '>')
    $query = [Diagnostics.Eventing.Reader.EventLogQuery]::new('Application', [Diagnostics.Eventing.Reader.PathType]::LogName, $xpath)
    $query.ReverseDirection = $true
    $reader = [Diagnostics.Eventing.Reader.EventLogReader]::new($query)
    try {
        # Bound each read even during an event storm; only the newest batch is shown.
        for ($index = 0; $index -lt 64; $index++) {
            $record = $reader.ReadEvent()
            if ($null -eq $record) { break }
            try {
                [pscustomobject]@{
                    Id = [long]$record.RecordId
                    Text = ([string]$record.Properties[0].Value -split '\r?\n', 2)[0]
                }
            } finally { $record.Dispose() }
        }
    } finally { $reader.Dispose() }
}

function Show-BlockToast([string]$Message, [int]$Count = 1) {
    $xml = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType=WindowsRuntime]::new()
    $xml.LoadXml((New-BlockToastXml $Message $Count))
    $toast = [Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType=WindowsRuntime]::new($xml)
    # Replace the previous entry instead of filling Notification Center indefinitely.
    $toast.Tag = 'blocked'
    $toast.Group = 'no360'
    $toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes(5)
    $notifier = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime]::CreateToastNotifier($appId)
    $notifier.Show($toast)
}

if ($LibraryOnly) { return }

$mutex = [Threading.Mutex]::new($false, 'Local\No360.BlockNotifications')
$ownsMutex = $false
try {
    try { $ownsMutex = $mutex.WaitOne(0) }
    catch [Threading.AbandonedMutexException] { $ownsMutex = $true }
    if (-not $ownsMutex) { return }

    # Register a distinct notification identity for this signed-in user.
    $registration = "HKCU:\Software\Classes\AppUserModelId\$appId"
    New-Item -Path $registration -Force | Out-Null
    New-ItemProperty -Path $registration -Name DisplayName -Value 'no360' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $registration -Name ShowInSettings -Value 1 -PropertyType DWord -Force | Out-Null
    if ($TestMessage) { Show-BlockToast $TestMessage; return }

    # Start at the current end: do not replay old blocks at sign-in.
    $initial = @(Read-BlockEvents 0)
    $lastId = 0L
    if ($initial.Count) { $lastId = $initial[0].Id }
    while ($true) {
        Start-Sleep -Seconds 5
        try {
            $events = @(Read-BlockEvents $lastId)
            if ($events.Count) {
                Show-BlockToast $events[0].Text $events.Count
                $lastId = $events[0].Id
            }
            # Recover if the Application log has been cleared.
            $latest = @(Read-BlockEvents 0)
            if (-not $latest.Count -or $latest[0].Id -lt $lastId) { $lastId = 0L }
        } catch {
            # Notifications are best effort. Keep protecting even if Windows
            # temporarily denies access to its notification or event-log services.
            Start-Sleep -Seconds 10
        }
    }
} finally {
    if ($ownsMutex) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
