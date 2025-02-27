using System;
using System.IO;
using System.Text.Json;
using DocumentProcessor.Models.Configuration;

namespace DocumentProcessor.Services
{
    public static class ConfigurationService
    {
        private const string DefaultConfigFileName = "appsettings.json";

        public static AzureDevOpsConfig LoadAzureDevOpsConfig(string? configFileName = null)
        {
            try
            {
                configFileName ??= DefaultConfigFileName;

                if (!File.Exists(configFileName))
                {
                    return new AzureDevOpsConfig();
                }

                var jsonString = File.ReadAllText(configFileName);
                var config = JsonSerializer.Deserialize<JsonElement>(jsonString);

                if (config.TryGetProperty("AzureDevOps", out var adoConfig))
                {
                    return JsonSerializer.Deserialize<AzureDevOpsConfig>(adoConfig.GetRawText()) 
                        ?? new AzureDevOpsConfig();
                }

                return new AzureDevOpsConfig();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading configuration from {configFileName}: {ex.Message}");
                return new AzureDevOpsConfig();
            }
        }
    }
}