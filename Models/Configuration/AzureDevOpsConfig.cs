using System;
using System.Text.Json.Serialization;

namespace DocumentProcessor.Models.Configuration
{
    public class AzureDevOpsConfig
    {
        [JsonPropertyName("organization")]
        public string Organization { get; set; } = string.Empty;

        [JsonPropertyName("personalAccessToken")]
        public string PersonalAccessToken { get; set; } = string.Empty;

        [JsonPropertyName("baseUrl")]
        public string BaseUrl { get; set; } = "https://dev.azure.com";

        public string GetConnectionUrl()
        {
            if (string.IsNullOrEmpty(Organization))
            {
                throw new InvalidOperationException(
                    "Azure DevOps organization not found in configuration.");
            }

            if (string.IsNullOrEmpty(PersonalAccessToken))
            {
                throw new InvalidOperationException(
                    "Azure DevOps Personal Access Token (PAT) not found in configuration.");
            }

            return $"{BaseUrl.TrimEnd('/')}";
        }
    }
}