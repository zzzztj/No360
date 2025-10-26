using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;

namespace No360.Models;

public sealed class FileFingerprint
{
    private FileFingerprint()
    {
    }

    public string FileName { get; private set; } = string.Empty;

    public string Path { get; private set; } = string.Empty;

    public string? CompanyName { get; private set; }

    public string? ProductName { get; private set; }

    public string? SignerSubject { get; private set; }

    public string? SignerIssuer { get; private set; }

    public static FileFingerprint From(string path, string name)
    {
        var fingerprint = new FileFingerprint
        {
            Path = path,
            FileName = name
        };

        try
        {
            var versionInfo = FileVersionInfo.GetVersionInfo(path);
            fingerprint.CompanyName = Normalize(versionInfo.CompanyName);
            fingerprint.ProductName = Normalize(versionInfo.ProductName);
        }
        catch
        {
            // Ignore metadata extraction failures and continue.
        }

        try
        {
            var certificate = X509Certificate.CreateFromSignedFile(path);
            var certificate2 = new X509Certificate2(certificate);
            fingerprint.SignerSubject = certificate2.Subject;
            fingerprint.SignerIssuer = certificate2.Issuer;
        }
        catch
        {
            // Ignore signature inspection failures and continue.
        }

        return fingerprint;

        static string? Normalize(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;
    }

    public string Summary()
    {
        return $"Company='{CompanyName ?? "?"}', Product='{ProductName ?? "?"}', Subject='{SignerSubject ?? "?"}'";
    }
}
