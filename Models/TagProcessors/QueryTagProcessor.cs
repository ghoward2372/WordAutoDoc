using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using DocumentProcessor.Services;

namespace DocumentProcessor.Models.TagProcessors
{
    public class QueryTagProcessor : ITagProcessor
    {
        private readonly IAzureDevOpsService _azureDevOpsService;
        private readonly IHtmlToWordConverter _htmlConverter;

        public QueryTagProcessor(IAzureDevOpsService azureDevOpsService, IHtmlToWordConverter htmlConverter)
        {
            _azureDevOpsService = azureDevOpsService ?? throw new ArgumentNullException(nameof(azureDevOpsService));
            _htmlConverter = htmlConverter ?? throw new ArgumentNullException(nameof(htmlConverter));
        }

        public async Task<string> ProcessTagAsync(string tagContent)
        {
            try
            {
                var query = await _azureDevOpsService.GetQueryAsync(tagContent);
                if (query?.Columns == null || !query.Columns.Any())
                    return "No columns defined in query.";

                var queryResult = await _azureDevOpsService.ExecuteQueryAsync(tagContent);
                if (queryResult?.WorkItems == null || !queryResult.WorkItems.Any())
                    return "No results found for query.";

                // Get work items with only the fields specified in the query
                var workItems = await _azureDevOpsService.GetWorkItemsAsync(
                    queryResult.WorkItems.Select(wi => wi.Id),
                    query.Columns.Select(c => c.ReferenceName)
                );

                if (!workItems.Any())
                    return "No work items found.";

                // Create table header row from query columns
                var tableData = new[]
                {
                    query.Columns.Select(c => c.Name).ToArray()
                }.Concat(
                    workItems.Select(wi => query.Columns
                        .Select(col => GetFieldValue(wi.Fields, col.ReferenceName))
                        .ToArray()
                    )
                ).ToArray();

                var table = _htmlConverter.CreateTable(tableData);
                return table.OuterXml;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing query {tagContent}: {ex.Message}");
                return $"Error processing query: {ex.Message}";
            }
        }

        private static string GetFieldValue(IDictionary<string, object> fields, string fieldName)
        {
            return fields.TryGetValue(fieldName, out var value) ? value?.ToString() ?? string.Empty : string.Empty;
        }
    }
}