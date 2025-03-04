using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

public class SBOMGenerator
{
    private readonly string _installPath;
    private readonly string _nistApiKey;
    private readonly List<string> _thirdPartyVendors = new List<string>();
    private readonly List<SBOMEntry> _sbomEntries = new List<SBOMEntry>();
    private readonly Dictionary<string, List<CVEEntry>> _vendorCVEs = new Dictionary<string, List<CVEEntry>>();
    private static readonly HttpClient _httpClient = new HttpClient();
    private readonly List<SBOMComponent> _sbomComponents = new List<SBOMComponent>();

    public SBOMGenerator(string installPath, IConfiguration configuration)
    {

        _installPath = installPath.Trim().Trim('"', '“', '”');

        if (!Directory.Exists(_installPath))
        {
            throw new DirectoryNotFoundException($"Invalid install path: {_installPath}");
        }
        if (Path.IsPathRooted(_installPath))
        {
            _installPath = Path.GetFullPath(new Uri(_installPath).LocalPath);
        }
        _nistApiKey = configuration["AzureDevOps:NISTApiKey"] ?? throw new ArgumentNullException("NIST API Key not found in configuration");
    }

    public async Task<string> GenerateSBOMAsync()
    {
        ProcessFiles();
        await FetchCVEDataAsync();
        return GenerateCycloneDxSBOM();
    }

    private void ProcessFiles()
    {
        DirectoryInfo directory = new DirectoryInfo(_installPath);
        try
        {
            foreach (var file in directory.GetFiles("*", SearchOption.AllDirectories))
            {
                try
                {
                    var fileInfo = new FileInfo(file.FullName);
                    var fileVersion = GetFileVersion(file.FullName);
                    var digitalSignature = GetDigitalSignature(file.FullName);
                    var hash = ComputeFileHash(file.FullName);

                    var supplier = ExtractCN(digitalSignature?.Signer) ?? "Unknown";
                    var entry = new SBOMEntry
                    {
                        FileName = fileInfo.Name,
                        FullPath = fileInfo.FullName,
                        FileVersion = fileVersion,
                        CreationDate = fileInfo.CreationTime,
                        Supplier = supplier,
                        DigitalSignature = digitalSignature,
                        Hash = hash,
                        IsThirdParty = IsThirdPartyVendor(supplier)
                    };

                    if (entry.IsThirdParty && !_thirdPartyVendors.Contains(supplier))
                    {
                        _thirdPartyVendors.Add(supplier);
                    }

                    _sbomEntries.Add(entry);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file {file.FullName}: {ex.Message}");
                }
            }
        }
        catch (IOException ioex)
        {
            Console.WriteLine("Error getting files from : " + directory.FullName + "   Exception: " + ioex.Message);
        }
    }

    private static string GetFileVersion(string filePath)
    {
        var versionInfo = FileVersionInfo.GetVersionInfo(filePath);
        return versionInfo.FileVersion ?? "Unknown";
    }

    private static DigitalSignatureInfo GetDigitalSignature(string filePath)
    {
        try
        {
            var cert = new X509Certificate2(filePath);
            return new DigitalSignatureInfo
            {
                Signer = cert.Subject,
                Algorithm = cert.SignatureAlgorithm.FriendlyName
            };
        }
        catch
        {
            return null;
        }
    }

    private static string ComputeFileHash(string filePath)
    {
        using var sha256 = SHA256.Create();
        using var stream = File.OpenRead(filePath);
        return BitConverter.ToString(sha256.ComputeHash(stream)).Replace("-", "").ToLower();
    }

    private static string ExtractCN(string subject)
    {
        if (string.IsNullOrEmpty(subject))
            return null;

        var match = Regex.Match(subject, "CN=([^,]+)");
        return match.Success ? match.Groups[1].Value.Trim() : subject;
    }

    private bool IsThirdPartyVendor(string supplier)
    {
        return !string.IsNullOrEmpty(supplier) && !supplier.Contains("OurCompanyName");
    }

    private async Task FetchCVEDataAsync()
    {
        foreach (var vendor in _thirdPartyVendors)
        {
            try
            {
                if (vendor.ToString() != "Unknown")
                {
                    var requestUrl = $"https://services.nvd.nist.gov/rest/json/cves/2.0?keywordSearch={vendor}";
                    var request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Add("apiKey", _nistApiKey);

                    var response = await _httpClient.SendAsync(request);
                    response.EnsureSuccessStatusCode();

                    var cveResponse = await response.Content.ReadAsStringAsync();
                    var cveData = JsonConvert.DeserializeObject<CVEResponse>(cveResponse);

                    if (cveData?.Vulnerabilities != null)
                    {
                        _vendorCVEs[vendor] = cveData.Vulnerabilities.Select(v => new CVEEntry
                        {
                            Id = v.CVE.Id,
                            Description = v.CVE.Descriptions?.FirstOrDefault()?.Value ?? "No description available"
                        }).ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching CVE data for {vendor}: {ex.Message}");
            }
        }
    }

    private string GenerateCycloneDxSBOM()
    {
        var sbom = new CycloneDxSBOM
        {
            BomFormat = "CycloneDX",
            SpecVersion = "1.4",
            SerialNumber = "urn:uuid:" + Guid.NewGuid(),
            Version = 1,
            Metadata = new Metadata
            {
                Timestamp = DateTime.UtcNow.ToString("o"),
                Tools = new List<Tool> { new Tool { Vendor = "Custom SBOM Generator", Name = "SBOMGen", Version = "1.0.0" } }
            },
            Components = _sbomComponents,
            Vulnerabilities = _vendorCVEs.Select(v => new SBOMVulnerability
            {
                Id = v.Value.FirstOrDefault()?.Id,
                Source = new Source { Name = "NVD" },
                References = v.Value.FirstOrDefault()?.References,
                Affects = new List<AffectedComponent>
                {
                    new AffectedComponent { Ref = v.Value.FirstOrDefault()?.Id }
                },
                Description = v.Value.FirstOrDefault()?.Description
            }).ToList()
        };

        return JsonConvert.SerializeObject(sbom, Formatting.Indented);
    }
}
public class HashEntry
{
    public string Alg { get; set; }
    public string Content { get; set; }
}

public class Supplier
{
    public string Name { get; set; }
}

public class ReferenceEntry
{
    public string Url { get; set; }
}

public class Metadata
{
    public string Timestamp { get; set; }
    public List<Tool> Tools { get; set; }
}

public class Tool
{
    public string Vendor { get; set; }
    public string Name { get; set; }
    public string Version { get; set; }
}

public class Source
{
    public string Name { get; set; }
}

public class AffectedComponent
{
    public string Ref { get; set; }
}

public class SBOMEntry
{
    public string FileName { get; set; }
    public string FullPath { get; set; }
    public string FileVersion { get; set; }
    public DateTime CreationDate { get; set; }
    public string Supplier { get; set; }
    public DigitalSignatureInfo DigitalSignature { get; set; }
    public string Hash { get; set; }
    public bool IsThirdParty { get; set; }
}

public class DigitalSignatureInfo
{
    public string Signer { get; set; }
    public string Algorithm { get; set; }
}

public class CVEEntry
{
    public string Id { get; set; }
    public string Description { get; set; }
    public List<ReferenceEntry> References { get; set; }
}

public class CVEResponse
{
    public int TotalResults { get; set; }
    public List<Vulnerability> Vulnerabilities { get; set; }
}

public class Vulnerability
{
    public CVEDetails CVE { get; set; }
}

public class CVEDetails
{
    public string Id { get; set; }
    public List<CVEDescription> Descriptions { get; set; }
}

public class CVEDescription
{
    public string Lang { get; set; }
    public string Value { get; set; }
}

public class CycloneDxSBOM
{
    public string BomFormat { get; set; }
    public string SpecVersion { get; set; }
    public string SerialNumber { get; set; }
    public int Version { get; set; }
    public Metadata Metadata { get; set; }
    public List<SBOMComponent> Components { get; set; }
    public List<SBOMVulnerability> Vulnerabilities { get; set; }
}


public class SBOMComponent
{
    public string Name { get; set; }
    public string Version { get; set; }
    public string Supplier { get; set; }
    public Dictionary<string, string> Hashes { get; set; }
    public string Type { get; set; }
}

public class SBOMVulnerability
{
    public string Id { get; set; }
    public Source Source { get; set; }
    public List<ReferenceEntry> References { get; set; }
    public List<AffectedComponent> Affects { get; set; }
    public string Description { get; set; }
}