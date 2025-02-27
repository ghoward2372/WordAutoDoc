using System;
using System.IO;
using System.Text.Json;
using DocumentProcessor.Models.Configuration;

namespace DocumentProcessor.Services
{
    public static class ConfigurationService
    {
        private const string ConfigFileName = "appsettings.json";
        
        public static AzureDevOpsConfig LoadAzureDevOpsConfig()
        {
            try
            {
                if (!File.Exists(ConfigFileName))
                {
                    return new AzureDevOpsConfig();
                }

                var jsonString = File.ReadAllText(ConfigFileName);
                var config = JsonSerializer.Deserialize<JsonElement>(jsonString);
                
                if (config.TryGetProperty("AzureDevOps", out var adoConfig))
                {
                    return JsonSerializer.Deserialize<AzureDevOpsConfig>(adoConfig.ToString()) 
                        ?? new AzureDevOpsConfig();
                }
                
                return new AzureDevOpsConfig();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading configuration: {ex.Message}");
                return new AzureDevOpsConfig();
            }
        }
    }
}
