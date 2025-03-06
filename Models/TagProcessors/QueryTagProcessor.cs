using DocumentProcessor.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentProcessor.Models.TagProcessors
{
    public class QueryTagProcessor : ITagProcessor
    {
        private readonly IAzureDevOpsService _azureDevOpsService;
        private readonly IHtmlToWordConverter _htmlConverter;
        private const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string LIST_START_MARKER = "<LIST_START>";
        private const string LIST_END_MARKER = "<LIST_END>";
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

                var parts = tagContent.Split('|');


                string qID = parts[0];
                var requestedColumns = (parts.Length == 2 && parts[1].StartsWith("Columns:"))
                ? parts[1].Substring("Columns:".Length)
                 .Split(',', StringSplitOptions.RemoveEmptyEntries)
                 .Select(x => x.Trim())
                 .ToList()
                  : new List<string> { "*" };

                Console.WriteLine($"Processing query tag: {tagContent}");

                if (!Guid.TryParse(qID, out var queryId))
                {
                    return ProcessingResult.FromText("Invalid query ID format. Expected a GUID.");
                }

                // First get the query definition to determine columns
                var query = await _azureDevOpsService.GetQueryAsync(qID);
                if (query?.Columns == null || !query.Columns.Any())
                    return ProcessingResult.FromText("No columns defined in query.");

                // Execute the query to get work item references
                var queryResult = await _azureDevOpsService.ExecuteQueryAsync(qID);
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
                    query.Columns.Where(col => requestedColumns.Contains("*") || requestedColumns.Contains(col.Name))
                    .Select(c => c.Name).ToArray()
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

        public async Task<ProcessingResult> ProcessQueryAsListAsync(string tagContent)
        {
            try
            {
                var parts = tagContent.Split('|');


                string qID = parts[0];
                var requestedColumns = (parts.Length == 2 && parts[1].StartsWith("Columns:"))
                ? parts[1].Substring("Columns:".Length)
                 .Split(',', StringSplitOptions.RemoveEmptyEntries)
                 .Select(x => x.Trim())
                 .ToList()
                  : new List<string> { "*" };

                Console.WriteLine($"Processing query tag as list: {tagContent}");

                if (!Guid.TryParse(qID, out var queryId))
                {
                    return ProcessingResult.FromText("Invalid query ID format. Expected a GUID.");
                }

                // Get the query definition to determine columns
                var query = await _azureDevOpsService.GetQueryAsync(qID);
                if (query?.Columns == null || !query.Columns.Any())
                    return ProcessingResult.FromText("No columns defined in query.");

                // Execute the query to get work item references
                var queryResult = await _azureDevOpsService.ExecuteQueryAsync(qID);
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

                // Build bulleted list data
                var listItems = new List<string>();
                foreach (var workItem in workItems)
                {
                    var bulletText = string.Join(" - ", query.Columns
                        .Where(col => requestedColumns.Contains("*") || requestedColumns.Contains(col.Name))
                        .Select(col => GetFieldValue(workItem.Fields, col.ReferenceName)));
                    listItems.Add(bulletText);
                }


                // Convert the list to Word XML format
                var listXml = ConvertPlainTextListToWordFormat(listItems, 1);

                Console.WriteLine("Generated Word List XML:");
                Console.WriteLine(listXml);
                var wrappedXml = LIST_START_MARKER + listXml + LIST_END_MARKER;

                return ProcessingResult.FromText(wrappedXml);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing query {tagContent} as list: {ex.Message}");
                return ProcessingResult.FromText($"Error processing query: {ex.Message}");
            }
        }

        private string ConvertPlainTextListToWordFormat(List<string> items, int numId)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < items.Count; i++)
            {
                sb.AppendLine("<w:p>");
                sb.AppendLine($"<w:pPr><w:pStyle w:val=\"ListParagraph\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"{numId}\"/></w:numPr><w:ind w:left=\"720\"/></w:pPr>");
                sb.AppendLine("<w:r>");
                sb.AppendLine($"<w:t xml:space=\"preserve\">{System.Security.SecurityElement.Escape(items[i])}</w:t>");
                sb.AppendLine("</w:r>");
                sb.AppendLine("</w:p>");
            }
            return sb.ToString();
        }

        private static string GetFieldValue(IDictionary<string, object> fields, string fieldName)
        {
            return fields.TryGetValue(fieldName, out var value) ? value?.ToString() ?? string.Empty : string.Empty;
        }
    }
}