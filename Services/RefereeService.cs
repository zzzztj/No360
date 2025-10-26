using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using No360.Configuration;
using No360.Models;

namespace No360.Services;

public sealed class RefereeService : BackgroundService
{
    private const string EventSourceName = "no360";
    private readonly ILogger<RefereeService> _logger;
    private readonly BlockRules _rules;
    private readonly RefereeSettings _settings;
    private ManagementEventWatcher? _processWatcher;
    private readonly List<FileSystemWatcher> _fileSystemWatchers = new();

    public RefereeService(ILogger<RefereeService> logger, BlockRules rules, RefereeSettings settings)
    {
        _logger = logger;
        _rules = rules;
        _settings = settings;
        EnsureEventLogSource();
    }

    protected override Task ExecuteAsync(CancellationToken stoppingToken)
    {
        StartProcessWatcher();
        StartFolderWatchers();
        _logger.LogInformation(
            "no360 started. Watching processes and {Count} folders.",
            _settings.WatchFolders.Length);
        return Task.CompletedTask;
    }

    public override Task StopAsync(CancellationToken cancellationToken)
    {
        _processWatcher?.Stop();
        _processWatcher?.Dispose();

        foreach (var watcher in _fileSystemWatchers)
        {
            watcher.EnableRaisingEvents = false;
            watcher.Dispose();
        }

        return base.StopAsync(cancellationToken);
    }

    private void StartProcessWatcher()
    {
        var query = new WqlEventQuery("SELECT * FROM Win32_ProcessStartTrace");
        _processWatcher = new ManagementEventWatcher(query);
        _processWatcher.EventArrived += OnProcessStarted;
        _processWatcher.Start();
    }

    private void OnProcessStarted(object? sender, EventArrivedEventArgs e)
    {
        try
        {
            var pid = Convert.ToInt32(e.NewEvent["ProcessID"]);
            var processName = Convert.ToString(e.NewEvent["ProcessName"]) ?? string.Empty;
            var processPath = ProcessPath(pid);

            if (string.IsNullOrEmpty(processPath))
            {
                return;
            }

            if (processName.Equals("msiexec.exe", StringComparison.OrdinalIgnoreCase))
            {
                var commandLine = GetCommandLine(pid);
                var msiPath = TryExtractMsiPath(commandLine);

                if (!string.IsNullOrEmpty(msiPath) && File.Exists(msiPath))
                {
                    var metadata = MsiMeta.TryLoad(msiPath);
                    if (metadata != null && _rules.MatchesMsi(metadata))
                    {
                        TryKill(pid);
                        Quarantine(msiPath);
                        LogBlock(
                            $"Blocked MSI install '{metadata.ProductName ?? "?"}' by '{metadata.Manufacturer ?? "?"}' from {msiPath}");
                        return;
                    }
                }
            }

            var fingerprint = FileFingerprint.From(processPath, processName);
            if (_rules.Matches(fingerprint))
            {
                TryKill(pid);
                Quarantine(processPath);
                LogBlock(
                    $"Blocked process '{processName}' (PID {pid}) from '{processPath}' [{fingerprint.Summary()}]");
            }
        }
        catch (Exception exception)
        {
            _logger.LogError(exception, "Error handling process start event.");
        }
    }

    private void StartFolderWatchers()
    {
        foreach (var folder in _settings.WatchFolders.Where(Directory.Exists))
        {
            var watcher = new FileSystemWatcher(folder)
            {
                Filter = "*.*",
                IncludeSubdirectories = true,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime | NotifyFilters.Size
            };

            watcher.Created += (_, args) => OnNewFile(args.FullPath);
            watcher.Renamed += (_, args) => OnNewFile(args.FullPath);
            watcher.EnableRaisingEvents = true;
            _fileSystemWatchers.Add(watcher);
        }
    }

    private void OnNewFile(string path)
    {
        try
        {
            var extension = Path.GetExtension(path).ToLowerInvariant();
            if (extension is not (".exe" or ".msi"))
            {
                return;
            }

            for (var attempt = 0; attempt < 10; attempt++)
            {
                if (CanOpen(path))
                {
                    break;
                }

                Thread.Sleep(100);
            }

            if (extension == ".msi")
            {
                var metadata = MsiMeta.TryLoad(path);
                if (metadata != null && _rules.MatchesMsi(metadata))
                {
                    Quarantine(path);
                    LogBlock(
                        $"Quarantined MSI '{metadata.ProductName ?? "?"}' by '{metadata.Manufacturer ?? "?"}' from {path}");
                    return;
                }
            }

            var fingerprint = FileFingerprint.From(path, Path.GetFileName(path));
            if (_rules.Matches(fingerprint))
            {
                Quarantine(path);
                LogBlock($"Quarantined installer '{path}' [{fingerprint.Summary()}]");
            }
        }
        catch (Exception exception)
        {
            _logger.LogError(exception, "Error handling new file {Path}", path);
        }
    }

    private void Quarantine(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return;
            }

            var destination = Path.Combine(
                _settings.QuarantineDirectory,
                $"{DateTime.UtcNow:yyyyMMdd-HHmmssfff}-{Path.GetFileName(path)}");

            Directory.CreateDirectory(_settings.QuarantineDirectory);
            File.Move(path, destination, overwrite: true);
        }
        catch (Exception exception)
        {
            _logger.LogWarning(exception, "Failed to quarantine {Path}", path);
        }
    }

    private static void TryKill(int pid)
    {
        try
        {
            using var process = Process.GetProcessById(pid);
            process.Kill(entireProcessTree: true);
        }
        catch
        {
            // Ignore failures when terminating processes.
        }
    }

    private static bool CanOpen(string path)
    {
        try
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void LogBlock(string message)
    {
        _logger.LogWarning(message);

        try
        {
            EventLog.WriteEntry(EventSourceName, message, EventLogEntryType.Warning, 3600);
        }
        catch
        {
            // Event log entry is best-effort only.
        }
    }

    private static string? GetCommandLine(int pid)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT CommandLine FROM Win32_Process WHERE ProcessId=" + pid);
            foreach (ManagementObject instance in searcher.Get())
            {
                return instance["CommandLine"]?.ToString();
            }
        }
        catch
        {
            // Swallow and return null.
        }

        return null;
    }

    private static string? TryExtractMsiPath(string? commandLine)
    {
        if (string.IsNullOrWhiteSpace(commandLine))
        {
            return null;
        }

        var parts = SplitArgs(commandLine);
        for (var index = 0; index < parts.Length; index++)
        {
            var argument = parts[index].Trim();
            if (argument.Equals("/i", StringComparison.OrdinalIgnoreCase) ||
                argument.Equals("/package", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 1 < parts.Length)
                {
                    return Unquote(parts[index + 1]);
                }
            }

            if (argument.EndsWith(".msi", StringComparison.OrdinalIgnoreCase))
            {
                return Unquote(argument);
            }
        }

        return null;

        static string Unquote(string value) => value.Trim().Trim('"');
    }

    private static string[] SplitArgs(string commandLine)
    {
        var list = new List<string>();
        var inQuotes = false;
        var current = new StringBuilder();

        foreach (var character in commandLine)
        {
            if (character == '"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (!inQuotes && char.IsWhiteSpace(character))
            {
                if (current.Length > 0)
                {
                    list.Add(current.ToString());
                    current.Clear();
                }
            }
            else
            {
                current.Append(character);
            }
        }

        if (current.Length > 0)
        {
            list.Add(current.ToString());
        }

        return list.ToArray();
    }

    private static string? ProcessPath(int pid)
    {
        const uint ProcessQueryLimitedInformation = 0x1000;
        var handle = OpenProcess(ProcessQueryLimitedInformation, false, (uint)pid);
        if (handle == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            var builder = new StringBuilder(32768);
            var size = builder.Capacity;
            return QueryFullProcessImageName(handle, 0, builder, ref size)
                ? builder.ToString(0, size)
                : null;
        }
        finally
        {
            CloseHandle(handle);
        }
    }

    private void EnsureEventLogSource()
    {
        try
        {
            if (!EventLog.SourceExists(EventSourceName))
            {
                EventLog.CreateEventSource(EventSourceName, "Application");
            }
        }
        catch (Exception exception)
        {
            _logger.LogDebug(exception, "Unable to verify Windows Event Log source {Source}", EventSourceName);
        }
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenProcess(uint access, bool inheritHandle, uint processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool QueryFullProcessImageName(IntPtr hProcess, int flags, System.Text.StringBuilder text, ref int size);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);
}
