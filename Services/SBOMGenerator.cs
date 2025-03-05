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
        string sampleSbom = GenerateCycloneDxSBOM(false);
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

        string cleanedProduct = productName.Trim();

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

        // Common vendor name corrections based on NVD database names
        var vendorMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "Microsoft Corporation", "Microsoft" },
        { "Google LLC", "Google" },
        { "Apple Inc.", "Apple" },
        { "Oracle America, Inc.", "Oracle" },
        { "IBM Corporation", "IBM" },
        { "Red Hat, Inc.", "Red Hat" }
    };

        return vendorMap.ContainsKey(vendor) ? vendorMap[vendor] : vendor;
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

                var component = new SBOMComponent
                {
                    bomRef = GeneratePurl(product, majorVersion, vendor),
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
                            bomRef = GeneratePurl("Microsoft .NET", frameworkVersion, "Microsoft"),
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
            Console.WriteLine($"⚠️ Error retrieving CPE: {ex.Message}");
            return "Unknown";
        }
    }






    private async Task FetchCVEDataAsync()
    {
        foreach (var component in _sbomComponents)
        {
            try
            {
                // ✅ Restore correct product name normalization
                string cleanedProductName = NormalizeProductName(component.name, component.version);

                // 🔍 If this is a .NET framework, find the correct CPE dynamically
                string bestCpe = "Unknown";
                if (cleanedProductName.ToLower().Contains("microsoft .net"))
                {
                    bestCpe = await GetCorrectCPEForDotNet(component.version, cleanedProductName);
                }

                // 🚨 Skip if no valid CPE was found
                if (bestCpe == "Unknown")
                {
                    Console.WriteLine($"⚠️ No valid CPE found for {component.name}, skipping CVE query.");
                    continue;
                }

                Console.WriteLine($"✅ Found CPE: {bestCpe} for {component.name}");

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
                    Console.WriteLine($"❌ No CVEs found for CPE: {bestCpe} (404 Not Found).");
                    continue;
                }

                cveResponse.EnsureSuccessStatusCode();

                var cveResponseString = await cveResponse.Content.ReadAsStringAsync();
                var cveData = JsonConvert.DeserializeObject<CVEResponse>(cveResponseString);

                if (cveData?.vulnerabilities != null && cveData.vulnerabilities.Any())
                {
                    Console.WriteLine($"✅ Found {cveData.vulnerabilities.Count} vulnerabilities for {component.name}");
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
public class SBOMVulnerability { public string id { get; set; } public Source source { get; set; } public List<ReferenceEntry> references { get; set; } public List<AffectedComponent> affects { get; set; } public string description { get; set; } }
public class CycloneDxSBOM { public string bomFormat { get; set; } public string specVersion { get; set; } public string serialNumber { get; set; } public int version { get; set; } public Metadata metadata { get; set; } public List<SBOMComponent> components { get; set; } public List<SBOMVulnerability> vulnerabilities { get; set; } }
public class SBOMComponent { public string bomRef { get; set; } public string type { get; set; } public string name { get; set; } public string version { get; set; } public Supplier supplier { get; set; } public List<HashEntry> hashes { get; set; } public string purl { get; set; } }
public class CVEEntry { public string Id { get; set; } public string Description { get; set; } public List<ReferenceEntry> References { get; set; } }
public class CVEResponse { public List<Vulnerability> vulnerabilities { get; set; } }
public class Vulnerability { public CVEDetails cve { get; set; } }
public class CVEDetails { public string id { get; set; } public List<CVEDescription> descriptions { get; set; } }
public class CVEDescription { public string Lang { get; set; } public string value { get; set; } }
