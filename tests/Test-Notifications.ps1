$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\Notifications\Notify-Blocks.ps1') -LibraryOnly
$xml = [xml](New-BlockToastXml 'Moved to quarantine: a&b<test>.exe')
if ($xml.toast.visual.binding.text[1] -ne 'Moved to quarantine: a&b<test>.exe') { throw 'Text escaping failed.' }
if ($xml.toast.audio.silent -ne 'true' -or $xml.toast.duration -ne 'short') { throw 'Expected a short silent toast.' }
$xml = [xml](New-BlockToastXml ('x' * 500) 3)
if ($xml.toast.visual.binding.text[1] -notlike '3 items blocked. Latest: *') { throw 'Batch summary missing.' }
if ($xml.toast.visual.binding.text[1].Length -gt 220) { throw 'Notification text is unbounded.' }
$events = @(Read-BlockEvents 0)
if ($events.Count -gt 64) { throw 'Event batch is unbounded.' }
if ($events.Count) {
    $newer = @(Read-BlockEvents $events[0].Id)
    if ($newer | Where-Object { $_.Id -le $events[0].Id }) { throw 'Record cursor filter failed.' }
}
Write-Output 'PASS: XML escaping, concise silent content, batch summary, bounded event reads, and record cursor.'
