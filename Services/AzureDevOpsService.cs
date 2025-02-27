using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using DocumentProcessor.Models.Configuration;

namespace DocumentProcessor.Services
{
    public interface IAzureDevOpsService
    {
        Task<string> GetWorkItemDocumentTextAsync(int workItemId);
        Task<WorkItemQueryResult> ExecuteQueryAsync(string queryId);
        Task<IEnumerable<WorkItem>> GetWorkItemsAsync(IEnumerable<int> workItemIds);
    }

    public class AzureDevOpsService : IAzureDevOpsService
    {
        private readonly WorkItemTrackingHttpClient _witClient;

        public AzureDevOpsService(WorkItemTrackingHttpClient witClient)
        {
            _witClient = witClient ?? throw new ArgumentNullException(nameof(witClient));
        }

        public static AzureDevOpsService Initialize()
        {
            var config = ConfigurationService.LoadAzureDevOpsConfig();
            var credentials = new VssBasicCredential(string.Empty, config.PersonalAccessToken);
            var connection = new VssConnection(new Uri(config.BaseUrl), credentials);

            try
            {
                var witClient = connection.GetClient<WorkItemTrackingHttpClient>();
                return new AzureDevOpsService(witClient);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to initialize Azure DevOps connection. Please verify your organization name and PAT. Details: {ex.Message}");
            }
        }

        public async Task<string> GetWorkItemDocumentTextAsync(int workItemId)
        {
            try
            {
                var workItem = await _witClient.GetWorkItemAsync(workItemId, expand: WorkItemExpand.All);
                if (workItem?.Fields == null)
                {
                    throw new InvalidOperationException($"Work item {workItemId} or its fields are null");
                }

                return workItem.Fields.TryGetValue("CAFRS.CAFRSSystem.DocumentPart.DocumentText", out object? value)
                    ? value?.ToString() ?? string.Empty
                    : string.Empty;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving work item {workItemId}: {ex.Message}", ex);
            }
        }

        public async Task<WorkItemQueryResult> ExecuteQueryAsync(string queryId)
        {
            try
            {
                if (!Guid.TryParse(queryId, out _))
                    throw new ArgumentException("Invalid query ID format. Expected a GUID.");

                return await _witClient.QueryByIdAsync(new Guid(queryId));
            }
            catch (Exception ex)
            {
                throw new Exception($"Error executing query {queryId}: {ex.Message}", ex);
            }
        }

        public async Task<IEnumerable<WorkItem>> GetWorkItemsAsync(IEnumerable<int> workItemIds)
        {
            try
            {
                if (!workItemIds.Any())
                    return new List<WorkItem>();

                return await _witClient.GetWorkItemsAsync(workItemIds, expand: WorkItemExpand.All);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error retrieving work items: {ex.Message}", ex);
            }
        }
    }
}