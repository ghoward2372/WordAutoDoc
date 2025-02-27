using System;
using System.IO;
using Xunit;
using DocumentProcessor.Models.Configuration;
using DocumentProcessor.Services;

namespace DocumentProcessor.Tests.Services
{
    public class ConfigurationServiceTests : IDisposable
    {
        private const string TestConfigFile = "appsettings.test.json";

        public ConfigurationServiceTests()
        {
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
        public void GetConnectionUrl_WithValidConfig_ReturnsCorrectUrl()
        {
            // Arrange
            var config = new AzureDevOpsConfig
            {
                Organization = "testorg",
                PersonalAccessToken = "testpat",
                BaseUrl = "https://dev.azure.com"
            };

            // Act & Assert
            Assert.Equal("https://dev.azure.com/testorg", config.GetConnectionUrl());
        }

        [Fact]
        public void GetConnectionUrl_WithMissingOrganization_ThrowsException()
        {
            // Arrange
            var config = new AzureDevOpsConfig
            {
                PersonalAccessToken = "testpat",
                BaseUrl = "https://dev.azure.com"
            };

            // Act & Assert
            var ex = Assert.Throws<InvalidOperationException>(() => config.GetConnectionUrl());
            Assert.Contains("organization not found", ex.Message);
        }

        [Fact]
        public void GetConnectionUrl_WithMissingPat_ThrowsException()
        {
            // Arrange
            var config = new AzureDevOpsConfig
            {
                Organization = "testorg",
                BaseUrl = "https://dev.azure.com"
            };

            // Act & Assert
            var ex = Assert.Throws<InvalidOperationException>(() => config.GetConnectionUrl());
            Assert.Contains("Personal Access Token", ex.Message);
        }

        public void Dispose()
        {
            if (File.Exists(TestConfigFile))
                File.Delete(TestConfigFile);
        }
    }
}