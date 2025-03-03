using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Models;
using DocumentProcessor.Models.TagProcessors;
using DocumentProcessor.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace DocumentProcessor.Services
{
    public class WordDocumentProcessor
    {
        private readonly DocumentProcessingOptions _options;
        private readonly Dictionary<string, ITagProcessor> _tagProcessors;
        private const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string TABLE_START_MARKER = "<TABLE_START>";
        private const string TABLE_END_MARKER = "<TABLE_END>";
        private const string LIST_START_MARKER = "<LIST_START>";
        private const string LIST_END_MARKER = "<LIST_END>";

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
                    if (processed.ProcessedText.Contains(TABLE_START_MARKER) || 
                        processed.ProcessedText.Contains(LIST_START_MARKER))
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
                Console.WriteLine("Inserting mixed content with tables and lists...");

                // Split content by markers
                var parts = new List<(string Content, string Type)>();
                var currentText = new StringBuilder();
                var lines = content.Split('\n');

                foreach (var line in lines)
                {
                    if (line.Trim() == TABLE_START_MARKER)
                    {
                        if (currentText.Length > 0)
                        {
                            parts.Add((currentText.ToString().Trim(), "text"));
                            currentText.Clear();
                        }
                        continue;
                    }
                    else if (line.Trim() == TABLE_END_MARKER)
                    {
                        if (currentText.Length > 0)
                        {
                            parts.Add((currentText.ToString().Trim(), "table"));
                            currentText.Clear();
                        }
                        continue;
                    }
                    else if (line.Trim() == LIST_START_MARKER)
                    {
                        if (currentText.Length > 0)
                        {
                            parts.Add((currentText.ToString().Trim(), "text"));
                            currentText.Clear();
                        }
                        continue;
                    }
                    else if (line.Trim() == LIST_END_MARKER)
                    {
                        if (currentText.Length > 0)
                        {
                            parts.Add((currentText.ToString().Trim(), "list"));
                            currentText.Clear();
                        }
                        continue;
                    }

                    currentText.AppendLine(line);
                }

                if (currentText.Length > 0)
                {
                    parts.Add((currentText.ToString().Trim(), "text"));
                }

                OpenXmlElement currentElement = paragraph;

                foreach (var part in parts)
                {
                    if (string.IsNullOrWhiteSpace(part.Content)) continue;

                    if (part.Type == "table")
                    {
                        Console.WriteLine("Processing embedded table");
                        var tableXml = part.Content;
                        if (!tableXml.Contains("xmlns:w="))
                        {
                            tableXml = tableXml.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                        }

                        var table = new Table();

                        // Parse table XML and load it into the table element
                        using (var stringReader = new StringReader(tableXml))
                        {
                            var xElement = XElement.Load(stringReader);
                            using (var reader = xElement.CreateReader())
                            {
                                table.Load(reader);
                            }
                        }

                        currentElement.InsertAfterSelf(table);
                        currentElement = table;
                        Console.WriteLine("Table inserted successfully");
                    }
                    else if (part.Type == "list")
                    {
                        Console.WriteLine("Processing embedded list");
                        foreach (var listParagraphXml in part.Content.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                        {
                            if (!string.IsNullOrWhiteSpace(listParagraphXml))
                            {
                                var paragraphXml = listParagraphXml.Trim();
                                if (!paragraphXml.Contains("xmlns:w="))
                                {
                                    paragraphXml = paragraphXml.Replace("<w:p>", $"<w:p xmlns:w=\"{WordMlNamespace}\">");
                                }

                                // Parse paragraph XML and create new paragraph
                                var newParagraph = new Paragraph();
                                using (var stringReader = new StringReader(paragraphXml))
                                {
                                    var xElement = XElement.Load(stringReader);
                                    using (var reader = xElement.CreateReader())
                                    {
                                        newParagraph.Load(reader);
                                    }
                                }

                                currentElement.InsertAfterSelf(newParagraph);
                                currentElement = newParagraph;
                            }
                        }
                        Console.WriteLine("List inserted successfully");
                    }
                    else
                    {
                        Console.WriteLine("Processing text content");
                        var newParagraph = new Paragraph(new Run(new Text(part.Content)));
                        currentElement.InsertAfterSelf(newParagraph);
                        currentElement = newParagraph;
                        Console.WriteLine("Text paragraph inserted");
                    }
                }

                paragraph.Remove();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing mixed content: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
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

                        if (processedContent.IsTable && processedContent.TableElement != null)
                        {
                            Console.WriteLine("Table found in processed content");
                            return processedContent;
                        }

                        text = text.Replace(match.Value, processedContent.ProcessedText);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing {tagProcessor.Key} tag: {ex.Message}");
                        text = text.Replace(match.Value, $"[Error processing {tagProcessor.Key} tag: {ex.Message}]");
                    }
                }
            }

            result.ProcessedText = _options.AcronymProcessor.ProcessText(text);
            return result;
        }
    }
}