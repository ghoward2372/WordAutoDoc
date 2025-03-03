using DocumentProcessor.Services;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Text;
using System.Net;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentProcessor.Models.TagProcessors
{
    public class WorkItemTagProcessor : ITagProcessor
    {
        private readonly IAzureDevOpsService _azureDevOpsService;
        private readonly IHtmlToWordConverter _htmlConverter;
        private readonly ITextBlockProcessor _textBlockProcessor;
        private const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string TABLE_START_MARKER = "<TABLE_START>";
        private const string TABLE_END_MARKER = "<TABLE_END>";
        private const string LIST_START_MARKER = "<LIST_START>";
        private const string LIST_END_MARKER = "<LIST_END>";

        public WorkItemTagProcessor(
            IAzureDevOpsService azureDevOpsService,
            IHtmlToWordConverter htmlConverter,
            ITextBlockProcessor? textBlockProcessor = null)
        {
            _azureDevOpsService = azureDevOpsService ?? throw new ArgumentNullException(nameof(azureDevOpsService));
            _htmlConverter = htmlConverter ?? throw new ArgumentNullException(nameof(htmlConverter));
            _textBlockProcessor = textBlockProcessor ?? new TextBlockProcessor();
        }

        public Task<ProcessingResult> ProcessTagAsync(string tagContent)
        {
            return ProcessTagAsync(tagContent, null);
        }

        public async Task<ProcessingResult> ProcessTagAsync(string tagContent, DocumentProcessingOptions? options)
        {
            if (!int.TryParse(tagContent, out int workItemId))
            {
                return ProcessingResult.FromText($"[Invalid work item ID: {tagContent}]");
            }

            try
            {
                var documentText = await _azureDevOpsService.GetWorkItemDocumentTextAsync(workItemId, options?.FQDocumentField ?? string.Empty);
                if (string.IsNullOrEmpty(documentText))
                    return ProcessingResult.FromText("[Work Item not found or empty]");

                Console.WriteLine($"\n=== Processing Work Item {workItemId} ===");
                Console.WriteLine($"Raw document text:\n{documentText}");

                var processedContent = new StringBuilder();
                var blocks = _textBlockProcessor.SegmentText(documentText);
                Console.WriteLine($"Text segmented into {blocks.Count} blocks");

                foreach (var block in blocks)
                {
                    Console.WriteLine($"\nProcessing block type: {block.Type}");
                    Console.WriteLine($"Block content preview: {block.Content.Substring(0, Math.Min(100, block.Content.Length))}");

                    if (block.Type == TextBlockProcessor.BlockType.Table)
                    {
                        Console.WriteLine("Converting table block to Word format...");
                        var tableData = ExtractTableData(block.Content);
                        if (tableData.Length > 0)
                        {
                            var table = _htmlConverter.CreateTable(tableData);
                            var tableXml = table.OuterXml;

                            processedContent.AppendLine(TABLE_START_MARKER);
                            processedContent.AppendLine(tableXml);
                            processedContent.AppendLine(TABLE_END_MARKER);
                            Console.WriteLine("Table conversion complete");
                        }
                    }
                    else if (block.Type == TextBlockProcessor.BlockType.List)
                    {
                        Console.WriteLine("Converting list block to Word format...");
                        var listXml = _htmlConverter.CreateBulletList(block.Content);
                        processedContent.AppendLine(LIST_START_MARKER);
                        processedContent.AppendLine(listXml);
                        processedContent.AppendLine(LIST_END_MARKER);
                        Console.WriteLine("List conversion complete");
                    }
                    else
                    {
                        var convertedText = _htmlConverter.ConvertHtmlToWordFormat(block.Content);
                        processedContent.AppendLine(convertedText);
                    }
                }

                var result = processedContent.ToString();
                Console.WriteLine($"\n=== Final Content Status ===");
                Console.WriteLine($"Contains table XML: {result.Contains("<w:tbl")}");
                Console.WriteLine($"Contains list XML: {result.Contains("<w:p")}");
                Console.WriteLine($"Total length: {result.Length}");
                Console.WriteLine($"Content preview: {(result.Length > 100 ? result.Substring(0, 100) + "..." : result)}");

                return ProcessingResult.FromText(result);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing work item {tagContent}: {ex.Message}");
                return ProcessingResult.FromText($"[Error processing work item {tagContent}: {ex.Message}]");
            }
        }

        private string[][] ExtractTableData(string tableHtml)
        {
            Console.WriteLine($"Extracting data from table HTML...");
            var rows = new List<string[]>();

            var rowMatches = Regex.Matches(tableHtml, @"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            Console.WriteLine($"Found {rowMatches.Count} rows");

            foreach (Match rowMatch in rowMatches)
            {
                var cells = new List<string>();

                var cellMatches = Regex.Matches(rowMatch.Value, @"<(td|th)[^>]*>(.*?)</(?:td|th)>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                foreach (Match cellMatch in cellMatches)
                {
                    var cellContent = cellMatch.Groups[2].Value;
                    cellContent = Regex.Replace(cellContent, @"<[^>]+>", string.Empty);
                    cellContent = WebUtility.HtmlDecode(cellContent).Trim();
                    cells.Add(cellContent);
                }

                if (cells.Count > 0)
                {
                    rows.Add(cells.ToArray());
                }
            }

            if (rows.Count == 0)
            {
                Console.WriteLine("Warning: No valid data found in table");
                return new string[0][];
            }

            Console.WriteLine($"Extracted {rows.Count} rows with {rows[0].Length} columns");
            return rows.ToArray();
        }
    }
}