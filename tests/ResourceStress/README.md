# Native resource regression check

Run on Windows with .NET 8 or later:

```powershell
dotnet run --project tests/ResourceStress -c Release -- 5000
```

The optional second argument selects an executable with an embedded Authenticode
signature. The default is the installed `dotnet.exe`.

This creates a temporary MSI database without installing anything, checks both
metadata fields, and verifies that each inspection releases its file lock before
garbage collection. It also exercises malformed/missing MSIs, unsigned files,
and signed executable inspection. Private memory and handle deltas are reported
after warmup; they are diagnostic values, not machine-dependent pass thresholds.

The original MSI code uses an unsupported `IN` query. With only that query
corrected so the success path is reachable, the old code fails the exclusive-file
open assertion on its first inspection. The resource-lifetime fix passes it.

This check does not launch the blocking service, kill processes, or quarantine
files. A production-duration service soak is still needed to establish whether
these resource leaks account for all reported memory growth.
