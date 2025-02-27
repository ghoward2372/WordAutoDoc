using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;

namespace DocumentProcessor.Samples
{
    public class QueryExample
    {
        public static async Task RunQueryExampleAsync()
        {
            try
            {
                // Initialize the client
                var organization = Environment.GetEnvironmentVariable("ADO_ORGANIZATION");
                var pat = Environment.GetEnvironmentVariable("ADO_PAT");
                var baseUrl = $"https://dev.azure.com/{organization}";
                
                var credentials = new VssBasicCredential(string.Empty, pat);
                var connection = new VssConnection(new Uri(baseUrl), credentials);
                var witClient = connection.GetClient<WorkItemTrackingHttpClient>();

                // Example query ID (replace with your actual query ID)
                var queryId = "00000000-0000-0000-0000-000000000000";

                // Get query definition to see what fields are included
                var query = await witClient.GetQueryAsync(string.Empty, queryId);
                Console.WriteLine($"Query Name: {query.Name}");
                Console.WriteLine($"Query Columns: {string.Join(", ", query.Columns.Select(c => c.Name))}");

                // Execute the query
                var queryResult = await witClient.QueryByIdAsync(new Guid(queryId));
                Console.WriteLine($"Found {queryResult.WorkItems.Count()} work items");

                // Get full work item details
                var workItemIds = queryResult.WorkItems.Select(wi => wi.Id);
                var workItems = await witClient.GetWorkItemsAsync(
                    ids: workItemIds,
                    fields: query.Columns.Select(c => c.ReferenceName),
                    expand: WorkItemExpand.None
                );

                // Display results
                foreach (var workItem in workItems)
                {
                    Console.WriteLine($"\nWork Item {workItem.Id}:");
                    foreach (var field in workItem.Fields)
                    {
                        Console.WriteLine($"  {field.Key}: {field.Value}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing query: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                }
            }
        }
    }
}
