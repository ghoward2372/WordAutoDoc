using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
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
            var queryResult = await _azureDevOpsService.ExecuteQueryAsync(tagContent);

            if (queryResult?.WorkItems == null || !queryResult.WorkItems.Any())
                return "No results found for query.";

            // Get full work item details
            var workItemIds = queryResult.WorkItems.Select(wi => wi.Id).ToList();
            var workItems = await _azureDevOpsService.GetWorkItemsAsync(workItemIds);

            if (!workItems.Any())
                return "No work items found.";

            // Convert work items to table format with explicit null handling
            var tableData = workItems
                .Select(wi => new[]
                {
                    wi.Id.ToString(),
                    GetFieldValue(wi.Fields, "System.Title"),
                    GetFieldValue(wi.Fields, "System.State")
                })
                .ToArray();

            // Add header row
            var headerRow = new[] { "ID", "Title", "State" };
            var fullTableData = new[] { headerRow }
                .Concat(tableData)
                .Select(row => row.Select(cell => cell ?? string.Empty).ToArray())
                .ToArray();

            return ConvertToMarkdownTable(fullTableData);
        }

        private static string GetFieldValue(IDictionary<string, object> fields, string fieldName)
        {
            return fields.TryGetValue(fieldName, out var value) ? value?.ToString() ?? string.Empty : string.Empty;
        }

        private string ConvertToMarkdownTable(string[][] tableData)
        {
            if (tableData == null || tableData.Length == 0)
                return string.Empty;

            var table = new System.Text.StringBuilder();

            // Header
            table.AppendLine(string.Join(" | ", tableData[0]));
            table.AppendLine(string.Join(" | ", tableData[0].Select(_ => "---")));

            // Data rows
            for (int i = 1; i < tableData.Length; i++)
            {
                table.AppendLine(string.Join(" | ", tableData[i]));
            }

            return table.ToString();
        }
    }
}