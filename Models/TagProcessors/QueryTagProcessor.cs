using DocumentProcessor.Services;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DocumentProcessor.Models.TagProcessors
{
    public class QueryTagProcessor : ITagProcessor
    {
        private readonly IAzureDevOpsService _azureDevOpsService;
        private readonly IHtmlToWordConverter _htmlConverter;
        private const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public QueryTagProcessor(IAzureDevOpsService azureDevOpsService, IHtmlToWordConverter htmlConverter)
        {
            _azureDevOpsService = azureDevOpsService ?? throw new ArgumentNullException(nameof(azureDevOpsService));
            _htmlConverter = htmlConverter ?? throw new ArgumentNullException(nameof(htmlConverter));
        }

        public Task<ProcessingResult> ProcessTagAsync(string tagContent)
        {
            return ProcessTagAsync(tagContent, null);
        }

        public async Task<ProcessingResult> ProcessTagAsync(string tagContent, DocumentProcessingOptions? options)
        {
            try
            {
                Console.WriteLine($"Processing query tag: {tagContent}");

                if (!Guid.TryParse(tagContent, out var queryId))
                {
                    return ProcessingResult.FromText("Invalid query ID format. Expected a GUID.");
                }

                // First get the query definition to determine columns
                var query = await _azureDevOpsService.GetQueryAsync(tagContent);
                if (query?.Columns == null || !query.Columns.Any())
                    return ProcessingResult.FromText("No columns defined in query.");

                // Execute the query to get work item references
                var queryResult = await _azureDevOpsService.ExecuteQueryAsync(tagContent);
                if (queryResult?.WorkItems == null || !queryResult.WorkItems.Any())
                    return ProcessingResult.FromText("No results found for query.");

                // Get work items with only the fields specified in the query
                var workItems = await _azureDevOpsService.GetWorkItemsAsync(
                    queryResult.WorkItems.Select(wi => wi.Id),
                    query.Columns.Select(c => c.ReferenceName)
                );

                if (!workItems.Any())
                    return ProcessingResult.FromText("No work items found.");

                Console.WriteLine($"Query returned {workItems.Count()} work items");

                // Create table data - header row first
                var tableData = new List<string[]>
                {
                    // Header row using column names from query
                    query.Columns.Select(c => c.Name).ToArray()
                };

                // Add one row per work item
                foreach (var workItem in workItems)
                {
                    var row = query.Columns
                        .Select(col => GetFieldValue(workItem.Fields, col.ReferenceName))
                        .ToArray();
                    tableData.Add(row);
                }

                Console.WriteLine("Creating table from work item data...");
                foreach (var row in tableData)
                {
                    Console.WriteLine($"Row: {string.Join(" | ", row)}");
                }

                var table = _htmlConverter.CreateTable(tableData.ToArray());
                return ProcessingResult.FromTable(table);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing query {tagContent}: {ex.Message}");
                return ProcessingResult.FromText($"Error processing query: {ex.Message}");
            }
        }

        private static string GetFieldValue(IDictionary<string, object> fields, string fieldName)
        {
            return fields.TryGetValue(fieldName, out var value) ? value?.ToString() ?? string.Empty : string.Empty;
        }
    }
}