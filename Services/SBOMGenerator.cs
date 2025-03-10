﻿using Microsoft.Extensions.Configuration;
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
    private Dictionary<string, string> componentCPEQueries = new Dictionary<string, string>();

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
        string sampleSbom = GenerateCycloneDxSBOM();
        Console.WriteLine("🔍 Checking SBOM output:");
        foreach (var component in _sbomComponents.Take(5)) // Only print the first 5 to keep it readable
        {
            Console.WriteLine($"Component: {component.name}, bom-ref: {component.bomRef}");
        }
        return sampleSbom;
    }

    private string GetProductName(string filePath)
    {
        try
        {
            var fileVersionInfo = FileVersionInfo.GetVersionInfo(filePath);
            return !string.IsNullOrWhiteSpace(fileVersionInfo.ProductName) ? fileVersionInfo.ProductName : Path.GetFileNameWithoutExtension(filePath);
        }
        catch
        {
            return Path.GetFileNameWithoutExtension(filePath); // Fallback to filename
        }
    }

    private string GetMajorVersion(string productVersion)
    {
        if (string.IsNullOrEmpty(productVersion))
            return "Unknown";

        // Remove metadata after "+"
        productVersion = productVersion.Split('+')[0];

        // Extract first two segments (major.minor)
        var versionParts = productVersion.Split('.');
        return versionParts.Length >= 2 ? $"{versionParts[0]}.{versionParts[1]}" : productVersion;
    }

    private string NormalizeProductName(string productName, string majorVersion)
    {
        if (string.IsNullOrWhiteSpace(productName))
            return "Unknown";

        // 🔥 Remove trademark symbols but keep everything else intact
        string cleanedProduct = Regex.Replace(productName, "[\u00AE\u2122]", "").Trim();

        // Ensure the major version is part of the product name
        if (!cleanedProduct.Contains(majorVersion))
        {
            cleanedProduct = $"{cleanedProduct} {majorVersion}";
        }

        Console.WriteLine($"🔍 Normalized Product Name: '{productName}' → '{cleanedProduct}' (Major Version: {majorVersion})");
        return cleanedProduct;
    }




    private string NormalizeVendorName(string vendor)
    {
        if (string.IsNullOrEmpty(vendor)) return "Unknown";

        // 🔥 Remove trademark symbols (® ™) but keep everything else intact
        string cleanedVendor = Regex.Replace(vendor, "[\u00AE\u2122]", "").Trim();

        // 🔥 Keep all existing vendor name corrections
        var vendorMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "Microsoft Corporation", "Microsoft" },
        { "Google LLC", "Google" },
        { "Apple Inc.", "Apple" },
        { "Oracle America, Inc.", "Oracle" },
        { "IBM Corporation", "IBM" },
        { "Red Hat, Inc.", "Red Hat" }
    };

        return vendorMap.ContainsKey(cleanedVendor) ? vendorMap[cleanedVendor] : cleanedVendor;
    }



    private void ProcessFiles()
    {
        DirectoryInfo directory = new DirectoryInfo(_installPath);
        if (!directory.Exists)
        {
            Console.WriteLine($"Directory does not exist: {directory.FullName}");
            return;
        }

        bool dotNetDetected = false;

        foreach (var file in directory.EnumerateFiles("*", SearchOption.AllDirectories))
        {
            try
            {
                var fileInfo = new FileInfo(file.FullName);
                var productVersion = GetProductVersion(file.FullName);
                var majorVersion = GetMajorVersion(productVersion);
                var digitalSignature = GetDigitalSignature(file.FullName);
                var hash = ComputeFileHash(file.FullName);

                string rawVendor = ExtractCN(digitalSignature?.signer) ?? "Unknown";
                string vendor = NormalizeVendorName(rawVendor);

                string rawProduct = GetProductName(file.FullName);
                string product = NormalizeProductName(rawProduct, majorVersion);
                //if (string.Compare(vendor, "unknown", true) == 0 && product.ToLower().Contains("unknown"))
                //{
                //    Console.WriteLine("Unknown Product and Vendor, not adding to SBOM.");
                //    continue;
                //}
                var component = new SBOMComponent
                {
                    type = "file",
                    name = product,
                    version = majorVersion,
                    supplier = new Supplier { name = vendor },
                    hashes = new List<HashEntry> { new HashEntry { alg = "SHA-256", content = hash } },
                    purl = GeneratePurl(product, majorVersion, vendor)
                };

                _sbomComponents.Add(component);

                // 🔥 Detect the ACTUAL runtime config file
                if (!dotNetDetected && fileInfo.Name.EndsWith(".runtimeconfig.json", StringComparison.OrdinalIgnoreCase))
                {
                    string frameworkVersion = GetDotNetVersionFromRuntimeConfig(file.FullName);
                    if (!string.IsNullOrEmpty(frameworkVersion))
                    {
                        Console.WriteLine($"✅ Detected .NET Framework: {frameworkVersion}");
                        _sbomComponents.Add(new SBOMComponent
                        {
                            type = "framework",
                            name = "Microsoft .NET",
                            version = frameworkVersion,
                            supplier = new Supplier { name = "Microsoft" },
                            hashes = new List<HashEntry>(),
                            purl = GeneratePurl("Microsoft .NET", frameworkVersion, "Microsoft")
                        });
                        dotNetDetected = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file {file.FullName}: {ex.Message}");
            }
        }
    }






    private string GetProductVersion(string filePath)
    {
        try
        {
            var fileVersionInfo = FileVersionInfo.GetVersionInfo(filePath);
            return !string.IsNullOrWhiteSpace(fileVersionInfo.ProductVersion) ? fileVersionInfo.ProductVersion : "Unknown";
        }
        catch
        {
            return "Unknown";
        }
    }

    private async Task<string> GetCorrectCPEForDotNet(string dotNetVersion, string rawProductName)
    {
        try
        {
            // ✅ Uses NormalizeProductName with both parameters
            string productName = NormalizeProductName(rawProductName, dotNetVersion);
            string keywordQuery = $"{productName} {dotNetVersion}";

            string cpeSearchUrl = $"https://services.nvd.nist.gov/rest/json/cpes/2.0?keywordSearch={Uri.EscapeDataString(keywordQuery)}";

            Console.WriteLine($"🔍 Searching for CPE: {keywordQuery}");

            var cpeRequest = new HttpRequestMessage(HttpMethod.Get, cpeSearchUrl);
            cpeRequest.Headers.Add("apiKey", _nistApiKey);
            cpeRequest.Headers.Add("User-Agent", "SBOM-Generator/1.0");

            var cpeResponse = await _httpClient.SendAsync(cpeRequest);
            cpeResponse.EnsureSuccessStatusCode();

            var cpeResponseString = await cpeResponse.Content.ReadAsStringAsync();
            var cpeData = JsonConvert.DeserializeObject<CPEApiResponse>(cpeResponseString);

            if (cpeData?.products != null && cpeData.products.Any())
            {
                foreach (var product in cpeData.products)
                {
                    string foundCpe = product.cpe.cpeName;

                    if (foundCpe.Contains(".net", StringComparison.OrdinalIgnoreCase) &&
                        foundCpe.Contains(dotNetVersion) &&
                        foundCpe.StartsWith("cpe:2.3:a:microsoft"))
                    {
                        Console.WriteLine($"✅ Found valid CPE: {foundCpe}");
                        return foundCpe; // ✅ Use exact CPE Name
                    }
                }
            }

            Console.WriteLine($"❌ No exact CPE match found for .NET {dotNetVersion}");
            return "Unknown";
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error retrieving CPE: {ex.Message}");
            return "Unknown";
        }
    }


    private async Task FetchCVEDataAsync()
    {
        List<string> failedCPEComponents = new List<string>(); // 🔥 Store components with no CPE match
        Dictionary<string, string> resolvedCPEs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // 🔥 Track searched CPEs for components
        HashSet<string> processedCPEs = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // 🔥 Prevent duplicate CVE queries
        List<SBOMVulnerability> collectedVulnerabilities = new List<SBOMVulnerability>(); // 🔥 Store vulnerabilities before adding to SBOM

        foreach (var component in _sbomComponents)
        {
            try
            {
                // ✅ Restore correct product name normalization
                string cleanedProductName = NormalizeProductName(component.name, component.version);

                string bestCpe = "Unknown";

                // 🔥 Check if we have already found a CPE for this component name
                if (resolvedCPEs.ContainsKey(cleanedProductName))
                {
                    bestCpe = resolvedCPEs[cleanedProductName];
                    Console.WriteLine($"⏩ Using cached CPE for: {cleanedProductName} → {bestCpe}");
                }
                else
                {
                    // 🔥 Special handling for .NET components (keep this working!)
                    if (cleanedProductName.ToLower().Contains("microsoft .net"))
                    {
                        bestCpe = await GetCorrectCPEForDotNet(component.version, cleanedProductName);
                    }
                    else
                    {
                        // 🔥 For all other components, attempt a generic CPE lookup
                        bestCpe = await GetCorrectCPEForProduct(component.name, component.version);
                    }

                    // 🔥 Store the resolved CPE for future identical components
                    if (bestCpe != "Unknown")
                    {
                        resolvedCPEs[cleanedProductName] = bestCpe;
                    }
                }

                // 🚨 Skip if no valid CPE was found
                if (bestCpe == "Unknown")
                {
                    failedCPEComponents.Add(component.name); // 🔥 Store the component instead of printing immediately
                    continue;
                }

                Console.WriteLine($"✅ Found CPE: {bestCpe} for {component.name}");

                // 🔥 Prevent duplicate CVE lookups for the same CPE
                if (processedCPEs.Contains(bestCpe))
                {
                    Console.WriteLine($"⏩ Skipping duplicate CVE query for CPE: {bestCpe}");
                    continue;
                }
                processedCPEs.Add(bestCpe);

                // 🔍 **Correctly format the CVE query**
                string encodedCpe = Uri.EscapeDataString(bestCpe);
                var cveUrl = $"https://services.nvd.nist.gov/rest/json/cves/2.0?cpeName={encodedCpe}";

                Console.WriteLine($"🔍 Querying CVEs for CPE: {bestCpe}");
                Console.WriteLine($"🛠️ FINAL URL: {cveUrl}");

                var cveRequest = new HttpRequestMessage(HttpMethod.Get, cveUrl);
                cveRequest.Headers.Add("apiKey", _nistApiKey);
                cveRequest.Headers.Add("User-Agent", "SBOM-Generator/1.0");

                var cveResponse = await _httpClient.SendAsync(cveRequest);

                if (cveResponse.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    Console.WriteLine($"❌ No CVEs found for CPE: {bestCpe} (404 Not Found). Retrying in 10 seconds...");
                    await Task.Delay(10000);

                    // Retry the request
                    cveResponse = await _httpClient.SendAsync(cveRequest);

                    if (cveResponse.StatusCode == System.Net.HttpStatusCode.NotFound)
                    {
                        Console.WriteLine($"❌ No CVEs found for CPE: {bestCpe} (404 Not Found) after retry.");
                        continue;
                    }
                }

                cveResponse.EnsureSuccessStatusCode();

                var cveResponseString = await cveResponse.Content.ReadAsStringAsync();
                var cveData = JsonConvert.DeserializeObject<CVEResponse>(cveResponseString);


                if (cveData?.vulnerabilities != null && cveData.vulnerabilities.Any())
                {
                    Console.WriteLine($"✅ Found {cveData.vulnerabilities.Count} vulnerabilities for {component.name}");

                    // 🔥 Ensure vulnerabilities are stored in _sbomVulnerabilities
                    foreach (var vuln in cveData.vulnerabilities)
                    {
                        var ratings = new List<VulnerabilityRating>();

                        // 🔥 Extract CVSS ratings from the API response
                        if (vuln.cve.metrics != null)
                        {
                            if (vuln.cve.metrics.cvssV3 != null && vuln.cve.metrics.cvssV3.Any())
                            {
                                var cvss3 = vuln.cve.metrics.cvssV3.First();
                                ratings.Add(new VulnerabilityRating
                                {
                                    method = "CVSSv3.1",
                                    severity = cvss3.cvssData.baseSeverity ?? "None",
                                    score = cvss3.cvssData.baseScore
                                });
                            }
                            else if (vuln.cve.metrics.cvssV2 != null && vuln.cve.metrics.cvssV2.Any())
                            {
                                var cvss2 = vuln.cve.metrics.cvssV2.First();
                                ratings.Add(new VulnerabilityRating
                                {
                                    method = "CVSSv2",
                                    severity = cvss2.cvssData.baseSeverity ?? "None",
                                    score = cvss2.cvssData.baseScore
                                });
                            }
                        }

                        var sbomVuln = new SBOMVulnerability
                        {
                            id = vuln.cve.id,
                            source = new Source { name = "NVD" },
                            references = vuln.cve.references?.Select(r => new ReferenceEntry { url = r.url, type = "cve" }).ToList() ?? new List<ReferenceEntry>(),
                            affects = new List<AffectedComponent> { new AffectedComponent { @ref = component.bomRef } },
                            description = vuln.cve.descriptions?.FirstOrDefault()?.value ?? "No description available",
                            ratings = ratings // 🔥 Now correctly stores CVSS ratings
                        };

                        collectedVulnerabilities.Add(sbomVuln);
                    }
                }
                else
                {
                    Console.WriteLine($"✅ No vulnerabilities found for CPE: {bestCpe}");
                }
            }
            catch (HttpRequestException httpEx)
            {
                Console.WriteLine($"⚠️ HTTP Error fetching CVEs: {httpEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Unexpected Error fetching CVEs: {ex.Message}");
            }
        }

        // 🔥 Add all collected vulnerabilities to _sbomVulnerabilities at the end
        _sbomVulnerabilities.AddRange(collectedVulnerabilities);

        // 🔥 Print all failed CPE lookups at the end for better readability
        if (failedCPEComponents.Count > 0)
        {
            Console.WriteLine("\n❌ The following components could not be matched to a CPE:");
            foreach (var component in failedCPEComponents)
            {
                Console.WriteLine($"   - {component}");
            }
        }
    }




    private async Task<string> GetCorrectCPEForProduct(string productName, string version)
    {
        try
        {
            string keywordQuery = $"{productName} {version}";

            string cpeSearchUrl = $"https://services.nvd.nist.gov/rest/json/cpes/2.0?keywordSearch={Uri.EscapeDataString(keywordQuery)}";

            Console.WriteLine($"🔍 Searching for CPE: {keywordQuery}");

            var cpeRequest = new HttpRequestMessage(HttpMethod.Get, cpeSearchUrl);
            cpeRequest.Headers.Add("apiKey", _nistApiKey);
            cpeRequest.Headers.Add("User-Agent", "SBOM-Generator/1.0");

            var cpeResponse = await _httpClient.SendAsync(cpeRequest);
            cpeResponse.EnsureSuccessStatusCode();

            var cpeResponseString = await cpeResponse.Content.ReadAsStringAsync();
            var cpeData = JsonConvert.DeserializeObject<CPEApiResponse>(cpeResponseString);

            if (cpeData?.products != null && cpeData.products.Any())
            {
                foreach (var product in cpeData.products)
                {
                    string foundCpe = product.cpe.cpeName;
                    if (foundCpe.Contains(productName, StringComparison.OrdinalIgnoreCase) &&
                        foundCpe.Contains(version))
                    {
                        Console.WriteLine($"✅ Found valid CPE: {foundCpe}");
                        return foundCpe;
                    }
                }
            }

            Console.WriteLine($"❌ No exact CPE match found for {productName} {version}");
            return "Unknown";
        }
        catch (Exception ex)
        {
            Console.WriteLine($"⚠️ Error retrieving CPE for " + productName + ": {ex.Message}");
            return "Unknown";
        }
    }

    private string GetDotNetVersionFromRuntimeConfig(string filePath)
    {
        try
        {
            // 🔥 Ensure we are working with a valid directory
            string? installDirectory = Path.GetDirectoryName(filePath);
            if (string.IsNullOrWhiteSpace(installDirectory) || !Directory.Exists(installDirectory))
            {
                Console.WriteLine($"❌ Invalid install directory: {installDirectory}");
                return string.Empty;
            }

            // 🔥 Find the correct .runtimeconfig.json file in the same directory as the executable
            string[] runtimeConfigFiles = Directory.GetFiles(installDirectory, "*.runtimeconfig.json", SearchOption.TopDirectoryOnly);

            if (runtimeConfigFiles.Length == 0)
            {
                Console.WriteLine("❌ No .runtimeconfig.json files found.");
                return string.Empty;
            }

            foreach (var configFile in runtimeConfigFiles)
            {
                string jsonContent = File.ReadAllText(configFile);
                dynamic json = JsonConvert.DeserializeObject(jsonContent);

                if (json?.runtimeOptions?.tfm != null)
                {
                    string version = json.runtimeOptions.tfm.ToString().Replace("net", "").Trim();
                    Console.WriteLine($"✅ Detected .NET Version: {version} from {configFile}");
                    return version;
                }
            }

            Console.WriteLine("❌ No valid .NET version found in runtime config files.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"⚠️ Error reading .NET version: {ex.Message}");
        }

        return string.Empty;
    }





    private string GetDotNetVersionFromCommand()
    {
        try
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "dotnet",
                    Arguments = "--list-runtimes",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            // Extract latest .NET version from output
            var lines = output.Split('\n');
            foreach (var line in lines.Reverse())
            {
                if (line.Contains("Microsoft.NETCore.App"))
                {
                    return line.Split(' ')[1]; // Extract version
                }
            }
        }
        catch
        {
            return "Unknown";
        }

        return "Unknown";
    }


    private string DetectDotNetVersion(string exePath)
    {
        try
        {
            var fileVersionInfo = FileVersionInfo.GetVersionInfo(exePath);

            // Check if it is a .NET Core/.NET 5+ app
            if (!string.IsNullOrWhiteSpace(fileVersionInfo.ProductName) && fileVersionInfo.ProductName.Contains(".NET"))
            {
                return fileVersionInfo.ProductVersion;
            }

            // Check CLR runtime version (may help for some .NET applications)
            if (!string.IsNullOrWhiteSpace(fileVersionInfo.FileDescription) && fileVersionInfo.FileDescription.Contains("CLR"))
            {
                return fileVersionInfo.FileVersion;
            }
        }
        catch
        {
            return "Unknown";
        }

        return "Unknown";
    }

    private string GetDotNetVersionFromPEHeaders(string exePath)
    {
        try
        {
            var fileVersionInfo = FileVersionInfo.GetVersionInfo(exePath);

            // .NET Core / .NET 5+ apps store runtime version in ProductName
            if (!string.IsNullOrWhiteSpace(fileVersionInfo.ProductName) && fileVersionInfo.ProductName.Contains(".NET"))
            {
                return fileVersionInfo.ProductVersion;
            }

            // Check CLR runtime version for older .NET Framework
            if (!string.IsNullOrWhiteSpace(fileVersionInfo.FileDescription) && fileVersionInfo.FileDescription.Contains("CLR"))
            {
                return fileVersionInfo.FileVersion;
            }
        }
        catch
        {
            return "Unknown";
        }

        return "Unknown";
    }

    private bool IsDotNetAssembly(string exePath)
    {
        try
        {
            using (var stream = new FileStream(exePath, FileMode.Open, FileAccess.Read))
            using (var reader = new BinaryReader(stream))
            {
                stream.Seek(0x3C, SeekOrigin.Begin);
                int peHeaderOffset = reader.ReadInt32();
                stream.Seek(peHeaderOffset + 4, SeekOrigin.Begin);

                ushort machineType = reader.ReadUInt16();
                ushort sections = reader.ReadUInt16();
                stream.Seek(92, SeekOrigin.Current);

                uint clrHeaderRVA = reader.ReadUInt32();

                return clrHeaderRVA != 0;
            }
        }
        catch
        {
            return false;
        }
    }


    private string GetDotNetVersionFromRuntimeDll()
    {
        try
        {
            var coreclrPaths = new[]
            {
            @"C:\Program Files\dotnet\shared\Microsoft.NETCore.App\8.0.0\coreclr.dll",
            @"C:\Windows\Microsoft.NET\Framework\v4.0.30319\clr.dll",
            @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\clr.dll"
        };

            foreach (var path in coreclrPaths)
            {
                if (File.Exists(path))
                {
                    var fileVersion = FileVersionInfo.GetVersionInfo(path);
                    return fileVersion.ProductVersion.Split('+')[0]; // Remove metadata
                }
            }
        }
        catch { }

        return "Unknown";
    }


    private string GenerateCycloneDxSBOM()
    {
        var sbom = new CycloneDxSBOM
        {
            bomFormat = "CycloneDX",
            specVersion = "1.4",
            serialNumber = "urn:uuid:" + Guid.NewGuid(),
            version = 1,
            metadata = new Metadata
            {
                timestamp = DateTime.UtcNow.ToString("o"),
                tools = new List<Tool> { new Tool { vendor = "Custom SBOM Generator", name = "SBOMGen", version = "1.0.0" } }
            },
            components = _sbomComponents.Select(c => new SBOMComponent
            {
                bomRef = c.bomRef,  // 🔥 Ensure this value is explicitly copied into the JSON
                type = c.type,
                name = c.name,
                version = c.version,
                supplier = c.supplier,
                hashes = c.hashes,
                purl = c.purl
            }).ToList(),
            vulnerabilities = _sbomVulnerabilities
        };

        return JsonConvert.SerializeObject(sbom, Formatting.Indented);
    }


    private static string GetFileVersion(string filePath) => FileVersionInfo.GetVersionInfo(filePath).FileVersion ?? "Unknown";

    private static DigitalSignatureInfo GetDigitalSignature(string filePath)
    {
        try
        {
            var cert = new X509Certificate2(filePath);
            return new DigitalSignatureInfo { signer = cert.Subject, algorithm = cert.SignatureAlgorithm.FriendlyName };
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

public class CPEApiResponse
{
    public List<CPEProduct> products { get; set; }
}

public class CPEProduct
{
    public CPE cpe { get; set; }
}

public class CPE
{
    public string cpeName { get; set; }
}

public class DigitalSignatureInfo { public string signer { get; set; } public string algorithm { get; set; } }
public class HashEntry { public string alg { get; set; } public string content { get; set; } }

public class Supplier { public string name { get; set; } }
public class ReferenceEntry { public string type { get; set; } public string url { get; set; } }
public class Metadata { public string timestamp { get; set; } public List<Tool> tools { get; set; } }
public class Tool { public string vendor { get; set; } public string name { get; set; } public string version { get; set; } }
public class Source { public string name { get; set; } }
public class AffectedComponent { public string @ref { get; set; } }
//public class SBOMVulnerability { public string id { get; set; } public Source source { get; set; } public List<ReferenceEntry> references { get; set; } public List<AffectedComponent> affects { get; set; } public string description { get; set; } }
public class SBOMVulnerability
{
    public string id { get; set; }
    public Source source { get; set; }
    public List<ReferenceEntry> references { get; set; }
    public List<AffectedComponent> affects { get; set; }
    public string description { get; set; }
    public List<VulnerabilityRating> ratings { get; set; }  // 🔥 NEW FIELD
}
public class VulnerabilityRating
{
    public string method { get; set; }   // Example: "CVSSv3.1"
    public string severity { get; set; } // Example: "High"
    public double score { get; set; }    // Example: 7.8
}


public class CycloneDxSBOM { public string bomFormat { get; set; } public string specVersion { get; set; } public string serialNumber { get; set; } public int version { get; set; } public Metadata metadata { get; set; } public List<SBOMComponent> components { get; set; } public List<SBOMVulnerability> vulnerabilities { get; set; } }
//public class SBOMComponent { public string bomRef { get; set; } public string type { get; set; } public string name { get; set; } public string version { get; set; } public Supplier supplier { get; set; } public List<HashEntry> hashes { get; set; } public string purl { get; set; } }
public class SBOMComponent
{
    [JsonProperty("bom-ref")]
    public string bomRef { get; set; }

    public string type { get; set; }
    public string name { get; set; }
    public string version { get; set; }
    public Supplier supplier { get; set; }
    public List<HashEntry> hashes { get; set; }
    public string purl { get; set; }

    public SBOMComponent()
    {
        // 🔥 Generate a guaranteed unique bom-ref
        bomRef = $"pkg:{(supplier?.name ?? "unknown").ToLower()}/{name}-{version}@{version}-{Guid.NewGuid()}".Replace(" ", "-");
    }
}



public class CVEEntry { public string Id { get; set; } public string Description { get; set; } public List<ReferenceEntry> References { get; set; } }
public class CVEResponse { public List<Vulnerability> vulnerabilities { get; set; } }
public class Vulnerability { public CVEDetails cve { get; set; } }
public class CVEDetails
{
    public string id { get; set; }
    public List<CVEDescription> descriptions { get; set; }
    public List<CVEReference> references { get; set; }
    public CVEMetrics metrics { get; set; }  // 🔥 NEW FIELD for CVSS scores
}
public class CVEMetrics
{
    [JsonProperty("cvssMetricV31")]
    public List<CVSSv3> cvssV3 { get; set; }

    [JsonProperty("cvssMetricV2")]
    public List<CVSSv2> cvssV2 { get; set; }
}
public class CVSSv3
{
    public CVSSData cvssData { get; set; }
}

public class CVSSv2
{
    public CVSSData cvssData { get; set; }
}

public class CVSSData
{
    public double baseScore { get; set; }
    public string baseSeverity { get; set; }
}

public class CVEReference
{
    public string type { get; set; } = "cve";
    public string url { get; set; }
}

public class CVEDescription { public string Lang { get; set; } public string value { get; set; } }
