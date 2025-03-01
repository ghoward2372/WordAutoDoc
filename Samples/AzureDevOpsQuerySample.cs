using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;

namespace DocumentProcessor.Samples
{
    public class AzureDevOpsQuery
    {
        private readonly WorkItemTrackingHttpClient _workItemClient;
        private readonly string _project;

        public AzureDevOpsQuery(string organizationUrl, string personalAccessToken, string project)
        {
            var credentials = new VssBasicCredential(string.Empty, personalAccessToken);
            var connection = new VssConnection(new Uri(organizationUrl), credentials);
            _workItemClient = connection.GetClient<WorkItemTrackingHttpClient>();
            _project = project;
        }

        /// <summary>
        /// Calls a stored query in Azure DevOps and retrieves work item details.
        /// </summary>
        public async Task<List<WorkItem>> RunStoredQueryAsync(string queryPath)
        {
            try
            {
                // Step 1: Get the stored query ID
                QueryHierarchyItem query = await _workItemClient.GetQueryAsync(_project, queryPath);

                if (query == null || query.Id == Guid.Empty)
                {
                    Console.WriteLine($"Query '{queryPath}' not found.");
                    return new List<WorkItem>();
                }

                Console.WriteLine($"Query found! Executing query: {query.Name}");

                // Step 2: Execute the stored query to get work item references
                WorkItemQueryResult queryResult = await _workItemClient.QueryByIdAsync(query.Id);

                if (queryResult.WorkItems.Count() == 0)
                {
                    Console.WriteLine("No work items found for this query.");
                    return new List<WorkItem>();
                }

                // Step 3: Retrieve details for each work item
                List<int> workItemIds = queryResult.WorkItems.Select(wi => wi.Id).ToList();
                return await GetWorkItemsByIdsAsync(workItemIds);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing stored query: {ex.Message}");
                return new List<WorkItem>();
            }
        }

        /// <summary>
        /// Fetches details for a list of work item IDs.
        /// </summary>
        private async Task<List<WorkItem>> GetWorkItemsByIdsAsync(List<int> workItemIds)
        {
            const int batchSize = 200; // API Limit

            List<WorkItem> allWorkItems = new List<WorkItem>();

            for (int i = 0; i < workItemIds.Count; i += batchSize)
            {
                var batch = workItemIds.Skip(i).Take(batchSize).ToList();
                var workItems = await _workItemClient.GetWorkItemsAsync(batch, 
                    fields: new[] { "System.Title", "System.State", "System.AssignedTo" });
                allWorkItems.AddRange(workItems);
            }

            return allWorkItems;
        }
    }
}