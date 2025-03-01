using DocumentProcessor.Services;
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

        public QueryTagProcessor(IAzureDevOpsService azureDevOpsService, IHtmlToWordConverter htmlConverter)
        {
            _azureDevOpsService = azureDevOpsService ?? throw new ArgumentNullException(nameof(azureDevOpsService));
            _htmlConverter = htmlConverter ?? throw new ArgumentNullException(nameof(htmlConverter));
        }

        public Task<string> ProcessTagAsync(string tagContent)
        {
            var options = new DocumentProcessingOptions
            {
                SourcePath = string.Empty,
                OutputPath = string.Empty,
                AzureDevOpsService = null,
                AcronymProcessor = new AcronymProcessor(new Models.Configuration.AcronymConfiguration 
                { 
                    KnownAcronyms = new Dictionary<string, string>(),
                    IgnoredAcronyms = new HashSet<string>() 
                }),
                HtmlConverter = new HtmlToWordConverter(),
                FQDocumentField = string.Empty
            };

            return ProcessTagAsync(tagContent, options);
        }

        public async Task<string> ProcessTagAsync(string tagContent, DocumentProcessingOptions options)
        {
            try
            {
                if (!Guid.TryParse(tagContent, out var queryId))
                {
                    return "Invalid query ID format. Expected a GUID.";
                }

                // First get the query definition to determine columns
                var query = await _azureDevOpsService.GetQueryAsync(tagContent);
                if (query?.Columns == null || !query.Columns.Any())
                    return "No columns defined in query.";

                // Execute the query to get work item references
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

                var table = _htmlConverter.CreateTable(tableData.ToArray());
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