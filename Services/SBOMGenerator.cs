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
    private readonly List<SBOMComponent> _sbomComponents = new List<SBOMComponent>();
    private readonly Dictionary<string, List<CVEEntry>> _vendorCVEs = new Dictionary<string, List<CVEEntry>>();
    private static readonly HttpClient _httpClient = new HttpClient();
    private readonly List<SBOMVulnerability> _sbomVulnerabilities = new List<SBOMVulnerability>();

    public SBOMGenerator(string installPath, IConfiguration configuration)
    {
        _installPath = installPath.Trim().Trim('"', '“', '”');
        if (!Path.IsPathRooted(_installPath))
        {
            throw new ArgumentException("Invalid install path. Path must be absolute.");
        }
        _nistApiKey = configuration["AzureDevOps:NISTApiKey"] ?? throw new ArgumentNullException("NIST API Key not found in configuration");
    }

    public async Task<string> GenerateSBOMAsync()
    {
        ProcessFiles();
        await FetchCVEDataAsync();
        // Generate Sample SBOM with 3 Entries First
        string sampleSbom = GenerateCycloneDxSBOM(true);
        File.WriteAllText(@"C:\temp\sample_sbom.json", sampleSbom);
        Console.WriteLine("Sample SBOM saved: C:\\temp\\sample_sbom.json");
        return sampleSbom;
    }

    private void ProcessFiles()
    {
        DirectoryInfo directory = new DirectoryInfo(_installPath);
        if (!directory.Exists)
        {
            Console.WriteLine($"Directory does not exist: {directory.FullName}");
            return;
        }

        foreach (var file in directory.EnumerateFiles("*", SearchOption.AllDirectories))
        {
            try
            {
                var fileInfo = new FileInfo(file.FullName);
                var fileVersion = GetFileVersion(file.FullName);
                var digitalSignature = GetDigitalSignature(file.FullName);
                var hash = ComputeFileHash(file.FullName);
                var supplier = ExtractCN(digitalSignature?.Signer) ?? "Unknown";

                var component = new SBOMComponent
                {
                    bomRef = GeneratePurl(fileInfo.Name, fileVersion, supplier),
                    type = "file",
                    name = fileInfo.Name,
                    version = fileVersion,
                    supplier = new Supplier { name = supplier },
                    hashes = new List<HashEntry> { new HashEntry { alg = "SHA-256", content = hash } },
                    purl = GeneratePurl(fileInfo.Name, fileVersion, supplier)
                };

                if (!string.IsNullOrEmpty(supplier) && !_thirdPartyVendors.Contains(supplier))
                {
                    _thirdPartyVendors.Add(supplier);
                }

                _sbomComponents.Add(component);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file {file.FullName}: {ex.Message}");
            }
        }
    }

    private async Task FetchCVEDataAsync()
    {
        foreach (var component in _sbomComponents)
        {
            try
            {
                var requestUrl = $"https://services.nvd.nist.gov/rest/json/cves/2.0?keywordSearch={component.supplier.name}";
                var request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                request.Headers.Add("apiKey", _nistApiKey);

                var response = await _httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var cveResponse = await response.Content.ReadAsStringAsync();
                var cveData = JsonConvert.DeserializeObject<CVEResponse>(cveResponse);

                if (cveData?.vulnerabilities != null)
                {
                    var limitedCves = cveData.vulnerabilities.Take(3); // Limit total vulnerabilities to 3

                    foreach (var v in limitedCves)
                    {
                        _sbomVulnerabilities.Add(new SBOMVulnerability
                        {
                            id = v.cve.id,
                            source = new Source { name = "NVD" },
                            references = new List<ReferenceEntry>
                            {
                                new ReferenceEntry { type = "vulnerability", url = $"https://nvd.nist.gov/vuln/detail/{v.cve.id}" }
                            },
                            affects = new List<AffectedComponent> { new AffectedComponent { @ref = component.purl } },
                            description = v.cve.descriptions?.FirstOrDefault()?.value ?? "No description available"
                        });

                        if (_sbomVulnerabilities.Count >= 3) break; // Stop at exactly 3 vulnerabilities
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching CVE data for {component.supplier.name}: {ex.Message}");
            }
        }
    }

    private string GenerateCycloneDxSBOM(bool sampleMode)
    {
        var selectedComponents = sampleMode ? _sbomComponents.Take(1).ToList() : _sbomComponents;
        var selectedVulnerabilities = sampleMode ? _sbomVulnerabilities.Take(3).ToList() : _sbomVulnerabilities;

        var sbom = new CycloneDxSBOM
        {
            bomFormat = "CycloneDX",
            specVersion = "1.4",
            serialNumber = "urn:uuid:" + Guid.NewGuid(),
            version = 1,
            metadata = new Metadata
            {
                timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"), // Enforce timestamp format
                tools = new List<Tool> { new Tool { vendor = "Custom SBOM Generator", name = "SBOMGen", version = "1.0.0" } }
            },
            components = selectedComponents,
            vulnerabilities = selectedVulnerabilities
        };

        return JsonConvert.SerializeObject(sbom, Formatting.Indented);
    }


    private static string GetFileVersion(string filePath) => FileVersionInfo.GetVersionInfo(filePath).FileVersion ?? "Unknown";

    private static DigitalSignatureInfo GetDigitalSignature(string filePath)
    {
        try
        {
            var cert = new X509Certificate2(filePath);
            return new DigitalSignatureInfo { Signer = cert.Subject, Algorithm = cert.SignatureAlgorithm.FriendlyName };
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
        var match = Regex.Match(subject ?? "", "CN=([^,]+)");
        return match.Success ? match.Groups[1].Value.Trim() : subject;
    }

    private static string GeneratePurl(string name, string version, string supplier)
    {
        return $"pkg:{supplier.ToLower()}/{name}@{version}".Replace(" ", "-");
    }
}

public class DigitalSignatureInfo { public string Signer { get; set; } public string Algorithm { get; set; } }
public class HashEntry { public string alg { get; set; } public string content { get; set; } }

public class Supplier { public string name { get; set; } }
public class ReferenceEntry { public string type { get; set; } public string url { get; set; } }
public class Metadata { public string timestamp { get; set; } public List<Tool> tools { get; set; } }
public class Tool { public string vendor { get; set; } public string name { get; set; } public string version { get; set; } }
public class Source { public string name { get; set; } }
public class AffectedComponent { public string @ref { get; set; } }
public class SBOMVulnerability { public string id { get; set; } public Source source { get; set; } public List<ReferenceEntry> references { get; set; } public List<AffectedComponent> affects { get; set; } public string description { get; set; } }
public class CycloneDxSBOM { public string bomFormat { get; set; } public string specVersion { get; set; } public string serialNumber { get; set; } public int version { get; set; } public Metadata metadata { get; set; } public List<SBOMComponent> components { get; set; } public List<SBOMVulnerability> vulnerabilities { get; set; } }
public class SBOMComponent { public string bomRef { get; set; } public string type { get; set; } public string name { get; set; } public string version { get; set; } public Supplier supplier { get; set; } public List<HashEntry> hashes { get; set; } public string purl { get; set; } }
public class CVEEntry { public string Id { get; set; } public string Description { get; set; } public List<ReferenceEntry> References { get; set; } }
public class CVEResponse { public List<Vulnerability> vulnerabilities { get; set; } }
public class Vulnerability { public CVEDetails cve { get; set; } }
public class CVEDetails { public string id { get; set; } public List<CVEDescription> descriptions { get; set; } }
public class CVEDescription { public string Lang { get; set; } public string value { get; set; } }
