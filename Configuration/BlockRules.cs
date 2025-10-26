using System;
using System.Linq;
using No360.Models;

namespace No360.Configuration;

public sealed class BlockRules
{
    public string[] Publishers { get; set; } = Array.Empty<string>();
    public string[] CompanyNames { get; set; } = Array.Empty<string>();
    public string[] ProductNames { get; set; } = Array.Empty<string>();
    public string[] FileNamePatterns { get; set; } = Array.Empty<string>();
    public string[] MsiManufacturers { get; set; } = Array.Empty<string>();
    public string[] MsiProductNames { get; set; } = Array.Empty<string>();

    public bool Matches(FileFingerprint fingerprint)
    {
        return HasAny(Publishers, fingerprint.SignerSubject, fingerprint.SignerIssuer)
               || HasAny(CompanyNames, fingerprint.CompanyName)
               || HasAny(ProductNames, fingerprint.ProductName)
               || FileNamePatterns.Any(pattern =>
                   fingerprint.FileName.Contains(pattern, StringComparison.OrdinalIgnoreCase));
    }

    public bool MatchesMsi(MsiMeta meta)
    {
        return HasAny(MsiManufacturers, meta.Manufacturer) || HasAny(MsiProductNames, meta.ProductName);
    }

    private static bool HasAny(string[] needles, params string?[] haystack)
    {
        return needles.Any(needle => haystack.Any(entry =>
            !string.IsNullOrEmpty(entry) &&
            entry!.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0));
    }
}
