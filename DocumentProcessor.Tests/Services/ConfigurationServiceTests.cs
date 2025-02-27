using System;
using System.IO;
using System.Collections.Generic;
using Xunit;
using DocumentProcessor.Models.Configuration;
using DocumentProcessor.Services;

namespace DocumentProcessor.Tests.Services
{
    public class ConfigurationServiceTests : IDisposable
    {
        private const string TestConfigFile = "appsettings.test.json";
        private readonly Dictionary<string, string> _originalEnvVars;

        public ConfigurationServiceTests()
        {
            // Store original environment variables
            _originalEnvVars = new Dictionary<string, string>
            {
                { "ADO_ORGANIZATION", Environment.GetEnvironmentVariable("ADO_ORGANIZATION") ?? string.Empty },
                { "ADO_PAT", Environment.GetEnvironmentVariable("ADO_PAT") ?? string.Empty },
                { "ADO_BASEURL", Environment.GetEnvironmentVariable("ADO_BASEURL") ?? string.Empty }
            };

            // Clear environment variables for testing
            Environment.SetEnvironmentVariable("ADO_ORGANIZATION", null);
            Environment.SetEnvironmentVariable("ADO_PAT", null);
            Environment.SetEnvironmentVariable("ADO_BASEURL", null);

            // Ensure test file doesn't exist before each test
            if (File.Exists(TestConfigFile))
                File.Delete(TestConfigFile);
        }

        [Fact]
        public void LoadAzureDevOpsConfig_WithValidConfig_ReturnsConfig()
        {
            // Arrange
            var json = @"{
                ""AzureDevOps"": {
                    ""Organization"": ""testorg"",
                    ""PersonalAccessToken"": ""testpat"",
                    ""BaseUrl"": ""https://test.azure.com""
                }
            }";
            File.WriteAllText(TestConfigFile, json);

            // Act
            var result = ConfigurationService.LoadAzureDevOpsConfig(TestConfigFile);

            // Assert
            Assert.Equal("testorg", result.Organization);
            Assert.Equal("testpat", result.PersonalAccessToken);
            Assert.Equal("https://test.azure.com", result.BaseUrl);
        }

        [Fact]
        public void LoadAzureDevOpsConfig_WithMissingFile_ReturnsDefaultConfig()
        {
            // Act
            var result = ConfigurationService.LoadAzureDevOpsConfig("nonexistent.json");

            // Assert
            Assert.NotNull(result);
            Assert.Empty(result.Organization);
            Assert.Empty(result.PersonalAccessToken);
            Assert.Equal("https://dev.azure.com", result.BaseUrl);
        }

        [Fact]
        public void LoadAzureDevOpsConfig_WithEnvironmentVariables_OverridesConfig()
        {
            // Arrange
            var json = @"{
                ""AzureDevOps"": {
                    ""Organization"": ""testorg"",
                    ""PersonalAccessToken"": ""testpat"",
                    ""BaseUrl"": ""https://test.azure.com""
                }
            }";
            File.WriteAllText(TestConfigFile, json);

            Environment.SetEnvironmentVariable("ADO_ORGANIZATION", "envorg");
            Environment.SetEnvironmentVariable("ADO_PAT", "envpat");
            Environment.SetEnvironmentVariable("ADO_BASEURL", "https://env.azure.com");

            // Act
            var result = ConfigurationService.LoadAzureDevOpsConfig(TestConfigFile);

            // Assert
            Assert.Equal("envorg", result.Organization);
            Assert.Equal("envpat", result.PersonalAccessToken);
            Assert.Equal("https://env.azure.com", result.BaseUrl);
        }

        public void Dispose()
        {
            // Restore original environment variables
            foreach (var envVar in _originalEnvVars)
            {
                Environment.SetEnvironmentVariable(envVar.Key, envVar.Value);
            }

            // Cleanup test file
            if (File.Exists(TestConfigFile))
                File.Delete(TestConfigFile);
        }
    }
}