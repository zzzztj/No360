using System.Diagnostics;
using System.Runtime.InteropServices;
using No360.Models;

// No test-framework dependencies. Run on Windows:
// dotnet run --project tests/ResourceStress -c Release -- 5000
var iterations = args.Length == 0 ? 5000 : int.Parse(args[0]);
var directory = Path.Combine(Path.GetTempPath(), "no360-stress-" + Guid.NewGuid());
Directory.CreateDirectory(directory);
var msi = Path.Combine(directory, "metadata.msi");
var invalid = Path.Combine(directory, "invalid.msi");
var signedExe = args.Length > 1 ? args[1] : Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet", "dotnet.exe");
try
{
    CreateFixture(msi);
    File.WriteAllText(invalid, "Not an MSI database");
    for (var i = 0; i < 100; i++) Inspect();
    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();
    using var process = Process.GetCurrentProcess();
    process.Refresh();
    var startMemory = process.PrivateMemorySize64;
    var startHandles = process.HandleCount;
    for (var i = 0; i < iterations; i++) Inspect();
    process.Refresh();
    Console.WriteLine($"PASS: {iterations} scans; private memory delta: {(process.PrivateMemorySize64 - startMemory) / 1048576.0:F1} MiB; handle delta: {process.HandleCount - startHandles}");
}
catch (Exception exception)
{
    Console.Error.WriteLine(exception);
    Environment.ExitCode = 1;
}
finally
{
    Directory.Delete(directory, recursive: true);
}

void Inspect()
{
    var metadata = MsiMeta.TryLoad(msi);
    if (metadata?.Manufacturer != "No360 Test" || metadata.ProductName != "Resource Test")
        throw new Exception("MSI metadata changed or could not be read.");
    // An inspection must release the database immediately, without a GC cycle.
    using (File.Open(msi, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
    if (MsiMeta.TryLoad(invalid) != null || MsiMeta.TryLoad(msi + ".missing") != null)
        throw new Exception("Invalid/missing MSI should return null.");
    using (File.Open(invalid, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
    var fingerprint = FileFingerprint.From(signedExe, Path.GetFileName(signedExe));
    if (string.IsNullOrEmpty(fingerprint.SignerSubject))
        throw new Exception("Pass an executable with an embedded signature as the second argument.");
    FileFingerprint.From(invalid, "invalid.msi");
}

static void CreateFixture(string path)
{
    dynamic installer = Activator.CreateInstance(Type.GetTypeFromProgID("WindowsInstaller.Installer")!)!;
    object? database = null;
    try
    {
        database = installer.OpenDatabase(path, 3);
        Execute("CREATE TABLE `Property` (`Property` CHAR(72) NOT NULL, `Value` CHAR(0) LOCALIZABLE PRIMARY KEY `Property`)");
        Execute("INSERT INTO `Property` (`Property`, `Value`) VALUES ('Manufacturer', 'No360 Test')");
        Execute("INSERT INTO `Property` (`Property`, `Value`) VALUES ('ProductName', 'Resource Test')");
        ((dynamic)database).Commit();
    }
    finally
    {
        if (database != null) Marshal.ReleaseComObject(database);
        Marshal.ReleaseComObject(installer);
    }

    void Execute(string sql)
    {
        dynamic view = ((dynamic)database!).OpenView(sql);
        try { view.Execute(null); }
        finally
        {
            try { view.Close(); }
            finally { Marshal.ReleaseComObject(view); }
        }
    }
}
