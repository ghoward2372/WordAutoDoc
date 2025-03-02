using DocumentProcessor.Services;
using System;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace DocumentProcessor.Models.TagProcessors
{
    public class WorkItemTagProcessor : ITagProcessor
    {
        private readonly IAzureDevOpsService _azureDevOpsService;
        private readonly IHtmlToWordConverter _htmlConverter;
        private const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public WorkItemTagProcessor(IAzureDevOpsService azureDevOpsService, IHtmlToWordConverter htmlConverter)
        {
            _azureDevOpsService = azureDevOpsService ?? throw new ArgumentNullException(nameof(azureDevOpsService));
            _htmlConverter = htmlConverter ?? throw new ArgumentNullException(nameof(htmlConverter));
        }

        public Task<string> ProcessTagAsync(string tagContent)
        {
            return ProcessTagAsync(tagContent, null);
        }

        public async Task<string> ProcessTagAsync(string tagContent, DocumentProcessingOptions? options)
        {
            // Validate work item ID first - don't catch the exception here
            if (!int.TryParse(tagContent, out int workItemId))
            {
                throw new ArgumentException($"Invalid work item ID: {tagContent}");
            }

            try
            {
                var documentText = await _azureDevOpsService.GetWorkItemDocumentTextAsync(workItemId, options?.FQDocumentField ?? string.Empty);
                if (string.IsNullOrEmpty(documentText))
                    return string.Empty;

                Console.WriteLine($"\n=== Processing Work Item {workItemId} ===");
                Console.WriteLine($"Retrieved content:\n{documentText}");

                // Check if content contains tables
                if (documentText.Contains("<table", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Found table in work item content, converting to Word format...");
                    var convertedContent = _htmlConverter.ConvertHtmlToWordFormat(documentText);

                    Console.WriteLine($"Converted content from HTML converter:\n{convertedContent}");

                    // Check if the converted content contains a table
                    if (convertedContent.Contains("<w:tbl"))
                    {
                        // Ensure proper XML structure for tables
                        if (!convertedContent.Contains("xmlns:w="))
                        {
                            convertedContent = convertedContent.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                        }

                        Console.WriteLine($"Final table XML with namespace:\n{convertedContent}");
                        return convertedContent;
                    }
                    else
                    {
                        Console.WriteLine("Warning: HTML table was found but no Word table was generated");
                    }
                }

                // For non-table content, convert HTML to Word format
                var processedContent = _htmlConverter.ConvertHtmlToWordFormat(documentText);
                Console.WriteLine($"Processed non-table content:\n{processedContent}");
                return processedContent;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing work item {tagContent}: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }
    }
}