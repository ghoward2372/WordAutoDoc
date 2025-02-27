using System;
using System.IO;
using Microsoft.Extensions.Configuration;
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

                var configuration = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile(configFileName, optional: true, reloadOnChange: true)
                    .Build();

                var adoConfig = configuration.GetSection("AzureDevOps").Get<AzureDevOpsConfig>();
                return adoConfig ?? new AzureDevOpsConfig();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading configuration from {configFileName}: {ex.Message}");
                return new AzureDevOpsConfig();
            }
        }
    }
}