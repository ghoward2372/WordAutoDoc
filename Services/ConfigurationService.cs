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
                    .AddEnvironmentVariables(prefix: "ADO_") // Add environment variables with ADO_ prefix
                    .Build();

                var adoConfig = new AzureDevOpsConfig();
                configuration.GetSection("AzureDevOps").Bind(adoConfig);

                // Override with environment variables if they exist
                if (!string.IsNullOrEmpty(configuration["ORGANIZATION"]))
                    adoConfig.Organization = configuration["ORGANIZATION"];
                if (!string.IsNullOrEmpty(configuration["PAT"]))
                    adoConfig.PersonalAccessToken = configuration["PAT"];
                if (!string.IsNullOrEmpty(configuration["BASEURL"]))
                    adoConfig.BaseUrl = configuration["BASEURL"];

                return adoConfig;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading configuration from {configFileName}: {ex.Message}");
                return new AzureDevOpsConfig();
            }
        }
    }
}