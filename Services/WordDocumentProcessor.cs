using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Models;
using DocumentProcessor.Models.TagProcessors;
using DocumentProcessor.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

namespace DocumentProcessor.Services
{
    public class WordDocumentProcessor
    {
        private readonly DocumentProcessingOptions _options;
        private readonly Dictionary<string, ITagProcessor> _tagProcessors;
        private const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string TABLE_START_MARKER = "<TABLE_START>";
        private const string TABLE_END_MARKER = "<TABLE_END>";

        public WordDocumentProcessor(DocumentProcessingOptions options)
        {
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _tagProcessors = new Dictionary<string, ITagProcessor>();

            // Always add AcronymTable processor
            _tagProcessors.Add("AcronymTable", new AcronymTableTagProcessor(options.AcronymProcessor, options.HtmlConverter));

            // Only add Azure DevOps related processors if service is available
            if (options.AzureDevOpsService != null)
            {
                _tagProcessors.Add("WorkItem", new WorkItemTagProcessor(options.AzureDevOpsService, options.HtmlConverter));
                _tagProcessors.Add("QueryID", new QueryTagProcessor(options.AzureDevOpsService, options.HtmlConverter));
            }
        }

        public async Task ProcessDocumentAsync()
        {
            try
            {
                Console.WriteLine($"\n=== Starting Document Processing ===");
                Console.WriteLine($"Source: {_options.SourcePath}");
                Console.WriteLine($"Output: {_options.OutputPath}");

                File.Copy(_options.SourcePath, _options.OutputPath, true);

                using (var doc = WordprocessingDocument.Open(_options.OutputPath, true))
                {
                    var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
                    var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body is missing");

                    await ProcessDocumentContentAsync(body);
                    mainPart.Document.Save();
                }

                Console.WriteLine("\n=== Document Processing Complete ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing document: {ex.Message}");
                throw;
            }
        }

        private async Task ProcessDocumentContentAsync(Body body)
        {
            var paragraphsToProcess = body.Elements<Paragraph>().ToList();
            Console.WriteLine($"Processing {paragraphsToProcess.Count} paragraphs");

            foreach (var paragraph in paragraphsToProcess)
            {
                var text = paragraph.InnerText;
                Console.WriteLine($"\nProcessing paragraph text: {text}");

                var processed = await ProcessTextAsync(text);

                if (processed.IsTable && processed.TableElement != null)
                {
                    try
                    {
                        var table = processed.TableElement;
                        Console.WriteLine("Inserting table into document...");

                        // Add namespace to table if missing
                        if (!table.OuterXml.Contains("xmlns:w="))
                        {
                            Console.WriteLine("Adding namespace to table XML");
                            var newTable = new Table();
                            newTable.InnerXml = table.OuterXml.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                            table = newTable;
                        }

                        // Insert table and remove original paragraph
                        paragraph.InsertBeforeSelf(table);
                        paragraph.Remove();
                        Console.WriteLine("Table inserted successfully");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error inserting table: {ex.Message}");
                        throw;
                    }
                }
                else if (text != processed.ProcessedText)
                {
                    // Check if this is content from a WorkItem tag that might contain tables
                    if (processed.ProcessedText.Contains(TABLE_START_MARKER))
                    {
                        InsertMixedContent(processed.ProcessedText, paragraph);
                    }
                    else
                    {
                        Console.WriteLine("Updating paragraph text");
                        paragraph.RemoveAllChildren();
                        paragraph.AppendChild(new Run(new Text(processed.ProcessedText)));
                    }
                }
            }
        }

        private void InsertMixedContent(string content, Paragraph paragraph)
        {
            try
            {
                // Split content by table markers
                var parts = content.Split(new[] { TABLE_START_MARKER, TABLE_END_MARKER },
                                       StringSplitOptions.RemoveEmptyEntries);

                // Keep track of our current position in the document
                OpenXmlElement currentElement = paragraph;

                foreach (var part in parts)
                {
                    var trimmedPart = part.Trim();
                    if (string.IsNullOrEmpty(trimmedPart)) continue;

                    if (trimmedPart.StartsWith("<w:tbl"))
                    {
                        // Handle table
                        Console.WriteLine("Processing embedded table");
                        var table = new Table();
                        var tableXml = trimmedPart;
                        if (!tableXml.Contains("xmlns:w="))
                        {
                            tableXml = tableXml.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                        }
                        table.InnerXml = tableXml;
                        currentElement.InsertAfterSelf(table);
                        currentElement = table;
                    }
                    else
                    {
                        // Handle text
                        Console.WriteLine("Processing text content");
                        var newParagraph = new Paragraph(new Run(new Text(trimmedPart)));
                        currentElement.InsertAfterSelf(newParagraph);
                        currentElement = newParagraph;
                    }
                }

                // Remove original paragraph if it's now empty
                if (paragraph.ChildElements.Count == 0)
                {
                    paragraph.Remove();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing mixed content: {ex.Message}");
                throw;
            }
        }

        private async Task<ProcessingResult> ProcessTextAsync(string text)
        {
            var result = ProcessingResult.FromText(text);

            foreach (var tagProcessor in _tagProcessors)
            {
                var pattern = RegexPatterns.GetTagPattern(tagProcessor.Key);
                var matches = pattern.Matches(text);

                foreach (Match match in matches)
                {
                    try
                    {
                        Console.WriteLine($"\nProcessing {tagProcessor.Key} tag: {match.Value}");
                        var tagContent = match.Groups[1].Value;
                        var processedContent = await tagProcessor.Value.ProcessTagAsync(tagContent, _options);

                        // If the tag processor returned a table, use it directly
                        if (processedContent.IsTable && processedContent.TableElement != null)
                        {
                            Console.WriteLine("Table found in processed content");
                            return processedContent;
                        }

                        // Otherwise, replace the tag with the processed text
                        text = text.Replace(match.Value, processedContent.ProcessedText);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing {tagProcessor.Key} tag: {ex.Message}");
                        text = text.Replace(match.Value, $"[Error processing {tagProcessor.Key} tag: {ex.Message}]");
                    }
                }
            }

            // Process acronyms only for non-table content
            result.ProcessedText = _options.AcronymProcessor.ProcessText(text);
            return result;
        }
    }
}