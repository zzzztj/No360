using System.ComponentModel;
using System.Diagnostics;
using System.Management; // WMI
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

var host = Host.CreateDefaultBuilder(args)
    .UseWindowsService()
    .ConfigureAppConfiguration(cfg =>
    {
        cfg.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
    })
    .ConfigureServices((ctx, services) =>
    {
        services.AddSingleton(ctx.Configuration.GetSection("BlockRules").Get<BlockRules>() ?? new BlockRules());
        services.AddSingleton(new RefereeSettings(ctx.Configuration));
        services.AddHostedService<RefereeService>();
    })
    .Build();

await host.RunAsync();

sealed class RefereeSettings
{
    public string QuarantineDir { get; }
    public string[] WatchFolders { get; }
    public RefereeSettings(IConfiguration cfg)
    {
        QuarantineDir = Expand(cfg["QuarantineDir"] ?? @"C:\ProgramData\InstallReferee\Quarantine");
        Directory.CreateDirectory(QuarantineDir);
        WatchFolders = (cfg.GetSection("WatchFolders").Get<string[]>() ?? Array.Empty<string>())
            .Select(Expand).Distinct().ToArray();
        foreach (var f in WatchFolders) Directory.CreateDirectory(f);
    }
    static string Expand(string s) => Environment.ExpandEnvironmentVariables(s);
}

sealed class BlockRules
{
    public string[] Publishers { get; set; } = Array.Empty<string>();
    public string[] CompanyNames { get; set; } = Array.Empty<string>();
    public string[] ProductNames { get; set; } = Array.Empty<string>();
    public string[] FileNamePatterns { get; set; } = Array.Empty<string>();
    public string[] MsiManufacturers { get; set; } = Array.Empty<string>();
    public string[] MsiProductNames { get; set; } = Array.Empty<string>();

    public bool Matches(FileFingerprint f)
        => HasAny(Publishers, f.SignerSubject, f.SignerIssuer)
        || HasAny(CompanyNames, f.CompanyName)
        || HasAny(ProductNames, f.ProductName)
        || FileNamePatterns.Any(p => f.FileName.Contains(p, StringComparison.OrdinalIgnoreCase));

    public bool MatchesMsi(MsiMeta m)
        => HasAny(MsiManufacturers, m.Manufacturer) || HasAny(MsiProductNames, m.ProductName);

    static bool HasAny(string[] needles, params string?[] hays)
        => needles.Any(n => hays.Any(h => !string.IsNullOrEmpty(h) &&
             h!.IndexOf(n, StringComparison.OrdinalIgnoreCase) >= 0));
}

sealed class RefereeService : BackgroundService
{
    private readonly ILogger<RefereeService> _log;
    private readonly BlockRules _rules;
    private readonly RefereeSettings _cfg;
    private ManagementEventWatcher? _procWatcher;
    private readonly List<FileSystemWatcher> _fsWatchers = new();

    public RefereeService(ILogger<RefereeService> log, BlockRules rules, RefereeSettings cfg)
    { _log = log; _rules = rules; _cfg = cfg; EnsureEventLogSource(); }

    protected override Task ExecuteAsync(CancellationToken stoppingToken)
    {
        StartProcessWatcher();
        StartFolderWatchers();
        _log.LogInformation("Install Referee started. Watching processes and {Count} folders.", _cfg.WatchFolders.Length);
        return Task.CompletedTask;
    }

    public override Task StopAsync(CancellationToken cancellationToken)
    {
        _procWatcher?.Stop(); _procWatcher?.Dispose();
        foreach (var w in _fsWatchers) { w.EnableRaisingEvents = false; w.Dispose(); }
        return base.StopAsync(cancellationToken);
    }

    void StartProcessWatcher()
    {
        // Win32_ProcessStartTrace fires on every new process; we’ll query command line & parent
        var q = new WqlEventQuery("SELECT * FROM Win32_ProcessStartTrace");
        _procWatcher = new ManagementEventWatcher(q);
        _procWatcher.EventArrived += (_, e) =>
        {
            try
            {
                int pid = Convert.ToInt32(e.NewEvent["ProcessID"]);
                string name = Convert.ToString(e.NewEvent["ProcessName"]) ?? "";
                string? path = ProcessPath(pid);
                if (string.IsNullOrEmpty(path)) return;

                // Special handling: MSI installs (bundlers love /qn)
                if (name.Equals("msiexec.exe", StringComparison.OrdinalIgnoreCase))
                {
                    string? cmd = GetCommandLine(pid);
                    string? msi = TryExtractMsiPath(cmd);
                    if (msi != null && File.Exists(msi))
                    {
                        var meta = MsiMeta.TryLoad(msi);
                        if (meta != null && _rules.MatchesMsi(meta))
                        {
                            TryKill(pid);
                            Quarantine(msi);
                            LogBlock($"Blocked MSI install '{meta.ProductName ?? "?"}' by '{meta.Manufacturer ?? "?"}' from {msi}");
                            return;
                        }
                    }
                }

                // Generic EXE payloads (child of a bundler)
                var fp = FileFingerprint.From(path, name);
                if (_rules.Matches(fp))
                {
                    TryKill(pid);
                    Quarantine(path);
                    LogBlock($"Blocked process '{name}' (PID {pid}) from '{path}' [{fp.Summary()}]");
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Error handling process start event.");
            }
        };
        _procWatcher.Start();
    }

    void StartFolderWatchers()
    {
        foreach (var folder in _cfg.WatchFolders.Where(Directory.Exists))
        {
            var w = new FileSystemWatcher(folder)
            {
                Filter = "*.*",
                IncludeSubdirectories = true,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime | NotifyFilters.Size
            };
            w.Created += (_, e) => OnNewFile(e.FullPath);
            w.Renamed += (_, e) => OnNewFile(e.FullPath);
            w.EnableRaisingEvents = true;
            _fsWatchers.Add(w);
        }
    }

    void OnNewFile(string path)
    {
        try
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext is not (".exe" or ".msi")) return;

            // Wait briefly for the file to settle
            for (int i = 0; i < 10; i++) { if (CanOpen(path)) break; Thread.Sleep(100); }

            if (ext == ".msi")
            {
                var meta = MsiMeta.TryLoad(path);
                if (meta != null && _rules.MatchesMsi(meta))
                {
                    Quarantine(path);
                    LogBlock($"Quarantined MSI '{meta.ProductName ?? "?"}' by '{meta.Manufacturer ?? "?"}' from {path}");
                    return;
                }
            }

            var fp = FileFingerprint.From(path, Path.GetFileName(path));
            if (_rules.Matches(fp))
            {
                Quarantine(path);
                LogBlock($"Quarantined installer '{path}' [{fp.Summary()}]");
            }
        }
        catch (Exception ex)
        { _log.LogError(ex, "Error handling new file {Path}", path); }
    }

    void Quarantine(string path)
    {
        try
        {
            if (!File.Exists(path)) return;
            var dest = Path.Combine(_cfg.QuarantineDir, $"{DateTime.UtcNow:yyyyMMdd-HHmmssfff}-{Path.GetFileName(path)}");
            Directory.CreateDirectory(_cfg.QuarantineDir);
            File.Move(path, dest, overwrite: true);
        }
        catch (Exception ex)
        { _log.LogWarning(ex, "Failed to quarantine {Path}", path); }
    }

    static void TryKill(int pid)
    {
        try { using var p = Process.GetProcessById(pid); p.Kill(entireProcessTree: true); } catch { }
    }

    static bool CanOpen(string path)
    {
        try { using var s = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read); return true; }
        catch { return false; }
    }

    void LogBlock(string message)
    {
        _log.LogWarning(message);
        try { System.Diagnostics.EventLog.WriteEntry("InstallReferee", message, EventLogEntryType.Warning, 3600); } catch { }
    }

    static string? GetCommandLine(int pid)
    {
        try
        {
            using var s = new ManagementObjectSearcher(
                "SELECT CommandLine FROM Win32_Process WHERE ProcessId=" + pid);
            foreach (ManagementObject o in s.Get())
                return o["CommandLine"]?.ToString();
        }
        catch { }
        return null;
    }

    static string? TryExtractMsiPath(string? cmd)
    {
        if (string.IsNullOrWhiteSpace(cmd)) return null;
        // Common patterns: /i "C:\path\foo.msi"  |  /package C:\path\foo.msi  |  /qn /i foo.msi
        var parts = SplitArgs(cmd);
        for (int i = 0; i < parts.Length; i++)
        {
            var a = parts[i].Trim();
            if (a.Equals("/i", StringComparison.OrdinalIgnoreCase) ||
                a.Equals("/package", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 < parts.Length) return Unquote(parts[i + 1]);
            }
            if (a.EndsWith(".msi", StringComparison.OrdinalIgnoreCase)) return Unquote(a);
        }
        return null;

        static string Unquote(string s) => s.Trim().Trim('"');
    }

    static string[] SplitArgs(string cmd)
    {
        var list = new List<string>();
        bool inQ = false; var cur = new System.Text.StringBuilder();
        foreach (var ch in cmd)
        {
            if (ch == '"') { inQ = !inQ; continue; }
            if (!inQ && char.IsWhiteSpace(ch)) { if (cur.Length > 0) { list.Add(cur.ToString()); cur.Clear(); } }
            else cur.Append(ch);
        }
        if (cur.Length > 0) list.Add(cur.ToString());
        return list.ToArray();
    }

    static string? ProcessPath(int pid)
    {
        const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;
        var h = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, (uint)pid);
        if (h == IntPtr.Zero) return null;
        try
        {
            var sb = new System.Text.StringBuilder(32768);
            int size = sb.Capacity;
            return QueryFullProcessImageName(h, 0, sb, ref size) ? sb.ToString(0, size) : null;
        }
        finally { CloseHandle(h); }
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr OpenProcess(uint access, bool inheritHandle, uint processId);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool QueryFullProcessImageName(IntPtr hProcess, int flags, System.Text.StringBuilder text, ref int size);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool CloseHandle(IntPtr hObject);
}

sealed class FileFingerprint
{
    public string FileName { get; private set; } = "";
    public string Path { get; private set; } = "";
    public string? CompanyName { get; private set; }
    public string? ProductName { get; private set; }
    public string? SignerSubject { get; private set; }
    public string? SignerIssuer { get; private set; }

    public static FileFingerprint From(string path, string name)
    {
        var fp = new FileFingerprint { Path = path, FileName = name };
        try { var vi = FileVersionInfo.GetVersionInfo(path); fp.CompanyName = N(vi.CompanyName); fp.ProductName = N(vi.ProductName); } catch { }
        try { var cert = X509Certificate.CreateFromSignedFile(path); var c2 = new X509Certificate2(cert); fp.SignerSubject = c2.Subject; fp.SignerIssuer = c2.Issuer; } catch { }
        return fp;
        static string? N(string? s) => string.IsNullOrWhiteSpace(s) ? null : s;
    }

    public string Summary()
        => $"Company='{CompanyName ?? "?"}', Product='{ProductName ?? "?"}', Subject='{SignerSubject ?? "?"}'";
}

sealed class MsiMeta
{
    public string? Manufacturer { get; init; }
    public string? ProductName  { get; init; }

    public static MsiMeta? TryLoad(string msiPath)
    {
        try
        {
            // Use Windows Installer COM automation (present on Windows)
            Type? t = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            if (t == null) return null;
            dynamic inst = Activator.CreateInstance(t)!;
            dynamic db = inst.OpenDatabase(msiPath, 0);
            string q = "SELECT `Property`,`Value` FROM `Property` WHERE `Property` IN ('Manufacturer','ProductName')";
            dynamic view = db.OpenView(q); view.Execute(null);
            string? man = null, prod = null;
            while (true)
            {
                dynamic rec = view.Fetch(); if (rec == null) break;
                string prop = rec.StringData(1); string val = rec.StringData(2);
                if (prop == "Manufacturer") man = val;
                else if (prop == "ProductName") prod = val;
            }
            return new MsiMeta { Manufacturer = man, ProductName = prod };
        }
        catch { return null; }
    }
}
