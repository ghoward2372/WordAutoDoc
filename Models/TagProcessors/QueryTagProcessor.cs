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

            // Create table header row
            var tableData = new[]
            {
                new[] { "ID", "Title", "State" }
            }.Concat(
                workItems.Select(wi => new[]
                {
                    wi.Id.ToString(),
                    GetFieldValue(wi.Fields, "System.Title"),
                    GetFieldValue(wi.Fields, "System.State")
                })
            ).ToArray();

            var table = _htmlConverter.CreateTable(tableData);
            return table.OuterXml;
        }

        private static string GetFieldValue(IDictionary<string, object> fields, string fieldName)
        {
            return fields.TryGetValue(fieldName, out var value) ? value?.ToString() ?? string.Empty : string.Empty;
        }
    }
}