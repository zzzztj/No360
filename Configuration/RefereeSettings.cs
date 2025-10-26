using System;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;

namespace No360.Configuration;

public sealed class RefereeSettings
{
    public RefereeSettings(IConfiguration configuration)
    {
        QuarantineDirectory = Expand(configuration["QuarantineDir"] ??
                                     @"C:\\ProgramData\\no360\\Quarantine");
        Directory.CreateDirectory(QuarantineDirectory);

        WatchFolders = (configuration.GetSection("WatchFolders").Get<string[]>() ?? Array.Empty<string>())
            .Select(Expand)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        foreach (var folder in WatchFolders)
        {
            Directory.CreateDirectory(folder);
        }
    }

    public string QuarantineDirectory { get; }

    public string[] WatchFolders { get; }

    private static string Expand(string value)
    {
        return Environment.ExpandEnvironmentVariables(value);
    }
}
