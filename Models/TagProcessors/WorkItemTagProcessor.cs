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

                Console.WriteLine($"Retrieved work item content: {documentText}");

                // Check if content contains tables
                if (documentText.Contains("<table"))
                {
                    Console.WriteLine("Work item content contains tables, converting to Word format...");
                    var convertedContent = _htmlConverter.ConvertHtmlToWordFormat(documentText);

                    // Log the converted content for debugging
                    Console.WriteLine($"Converted content from HTML converter: {convertedContent}");

                    // Ensure proper XML structure for tables
                    if (convertedContent.Contains("<w:tbl"))
                    {
                        if (!convertedContent.Contains("xmlns:w="))
                        {
                            convertedContent = convertedContent.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                        }
                        Console.WriteLine($"Final table XML with namespace: {convertedContent}");
                        return convertedContent;
                    }
                }

                // For non-table content, convert HTML to Word format
                var processedContent = _htmlConverter.ConvertHtmlToWordFormat(documentText);
                Console.WriteLine($"Processed non-table content: {processedContent}");
                return processedContent;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing work item {tagContent}: {ex.Message}");
                throw;
            }
        }
    }
}
