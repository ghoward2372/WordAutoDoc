using System;
using System.Text.Json.Serialization;

namespace DocumentProcessor.Models.Configuration
{
    public class AzureDevOpsConfig
    {
        [JsonPropertyName("organization")]
        public string Organization { get; set; } = string.Empty;

        [JsonPropertyName("baseUrl")]
        public string BaseUrl { get; set; } = "https://dev.azure.com";

        public string PersonalAccessToken 
        { 
            get 
            {
                var pat = Environment.GetEnvironmentVariable("ADO_PAT");
                if (string.IsNullOrEmpty(pat))
                {
                    throw new InvalidOperationException(
                        "Azure DevOps Personal Access Token (ADO_PAT) not found in environment variables.");
                }
                return pat;
            }
        }

        public string GetConnectionUrl()
        {
            if (string.IsNullOrEmpty(Organization))
            {
                var envOrg = Environment.GetEnvironmentVariable("ADO_ORGANIZATION");
                if (string.IsNullOrEmpty(envOrg))
                {
                    throw new InvalidOperationException(
                        "Azure DevOps organization not found in config or environment variables (ADO_ORGANIZATION).");
                }
                Organization = envOrg;
            }

            return $"{BaseUrl.TrimEnd('/')}/{Organization.TrimEnd('/')}";
        }
    }
}
