using DocumentProcessor.Services;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Text;

namespace DocumentProcessor.Models.TagProcessors
{
    public class WorkItemTagProcessor : ITagProcessor
    {
        private readonly IAzureDevOpsService _azureDevOpsService;
        private readonly IHtmlToWordConverter _htmlConverter;
        private readonly TextBlockProcessor _textBlockProcessor;
        private const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public WorkItemTagProcessor(IAzureDevOpsService azureDevOpsService, IHtmlToWordConverter htmlConverter)
        {
            _azureDevOpsService = azureDevOpsService ?? throw new ArgumentNullException(nameof(azureDevOpsService));
            _htmlConverter = htmlConverter ?? throw new ArgumentNullException(nameof(htmlConverter));
            _textBlockProcessor = new TextBlockProcessor();
        }

        public Task<string> ProcessTagAsync(string tagContent)
        {
            return ProcessTagAsync(tagContent, null);
        }

        public async Task<string> ProcessTagAsync(string tagContent, DocumentProcessingOptions? options)
        {
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
                Console.WriteLine($"Raw document text:\n{documentText}");

                // Segment the text into blocks
                var blocks = _textBlockProcessor.SegmentText(documentText);
                Console.WriteLine($"Text segmented into {blocks.Count} blocks");

                var processedContent = new StringBuilder();
                var hasTableContent = false;

                // Process each block according to its type
                foreach (var block in blocks)
                {
                    Console.WriteLine($"\nProcessing block type: {block.Type}");
                    Console.WriteLine($"Block content length: {block.Content.Length}");

                    if (block.Type == TextBlockProcessor.BlockType.Table)
                    {
                        Console.WriteLine("Converting table block to Word format...");
                        var tableContent = _htmlConverter.ConvertHtmlToWordFormat(block.Content);
                        Console.WriteLine($"Initial converted table content:\n{tableContent}");

                        // Ensure proper XML structure for tables
                        if (tableContent.Contains("<w:tbl") && !tableContent.Contains("xmlns:w="))
                        {
                            tableContent = tableContent.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                            Console.WriteLine("Added XML namespace to table");
                        }

                        Console.WriteLine($"Final table XML:\n{tableContent}");
                        processedContent.Append(tableContent);
                        hasTableContent = true;
                    }
                    else
                    {
                        Console.WriteLine("Converting text block to Word format...");
                        var convertedText = _htmlConverter.ConvertHtmlToWordFormat(block.Content);
                        Console.WriteLine($"Converted text block:\n{convertedText}");
                        processedContent.Append(convertedText);
                    }
                }

                var result = processedContent.ToString();
                Console.WriteLine($"\n=== Final Content Status ===");
                Console.WriteLine($"Contains table XML: {result.Contains("<w:tbl")}");
                Console.WriteLine($"Total length: {result.Length}");
                Console.WriteLine($"Content preview: {(result.Length > 100 ? result.Substring(0, 100) + "..." : result)}");

                return result;
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