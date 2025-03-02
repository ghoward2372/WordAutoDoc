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
using System.Xml;

namespace DocumentProcessor.Services
{
    public class WordDocumentProcessor
    {
        private readonly DocumentProcessingOptions _options;
        private readonly Dictionary<string, ITagProcessor> _tagProcessors;
        private const string WordMlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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
                Console.WriteLine("\n=== Starting Document Processing ===");
                Console.WriteLine($"Source document: {_options.SourcePath}");
                Console.WriteLine($"Output document: {_options.OutputPath}");

                File.Copy(_options.SourcePath, _options.OutputPath, true);
                Console.WriteLine("Created output document successfully");

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
                Console.WriteLine($"\n=== Document Processing Failed ===");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        private async Task ProcessDocumentContentAsync(Body body)
        {
            var paragraphsToProcess = body.Elements<Paragraph>().ToList();
            Console.WriteLine($"\n=== Processing {paragraphsToProcess.Count} Paragraphs ===");

            foreach (var paragraph in paragraphsToProcess)
            {
                var text = paragraph.InnerText;
                Console.WriteLine($"\nProcessing paragraph: {text}");

                var processed = await ProcessTextAsync(text);

                if (processed.IsTable && processed.TableElement != null)
                {
                    Console.WriteLine("\n=== Table Processing ===");
                    try
                    {
                        var table = processed.TableElement;
                        var rowCount = table.Elements<TableRow>().Count();
                        var columnCount = table.Elements<TableRow>().FirstOrDefault()?.Elements<TableCell>().Count() ?? 0;

                        Console.WriteLine($"Table structure details:");
                        Console.WriteLine($"- Total rows: {rowCount}");
                        Console.WriteLine($"- Columns per row: {columnCount}");

                        // Insert the table before the current paragraph
                        paragraph.InsertBeforeSelf(table);
                        paragraph.Remove();

                        Console.WriteLine("Table successfully inserted into document");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error during table insertion: {ex.Message}");
                        throw;
                    }
                }
                else if (text != processed.ProcessedText)
                {
                    Console.WriteLine("Updating paragraph with processed text");
                    paragraph.RemoveAllChildren();
                    paragraph.AppendChild(new Run(new Text(processed.ProcessedText)));
                }
            }
        }

        private async Task<ProcessingResult> ProcessTextAsync(string text)
        {
            var result = new ProcessingResult { ProcessedText = text };

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

                        if (IsTableXml(processedContent))
                        {
                            Console.WriteLine("Converting XML to table...");
                            result.IsTable = true;
                            result.TableElement = CreateTableFromXml(processedContent);
                            return result;
                        }

                        text = text.Replace(match.Value, processedContent);
                        Console.WriteLine("Tag processed successfully");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing {tagProcessor.Key} tag: {ex.Message}");
                        text = text.Replace(match.Value, $"[Error processing {tagProcessor.Key} tag]");
                    }
                }
            }

            result.ProcessedText = _options.AcronymProcessor.ProcessText(text);
            return result;
        }

        private bool IsTableXml(string content)
        {
            if (string.IsNullOrEmpty(content))
                return false;

            var trimmedContent = content.Trim();
            Console.WriteLine($"Checking if content is table XML: {trimmedContent.Substring(0, Math.Min(100, trimmedContent.Length))}...");
            return trimmedContent.Contains("<w:tbl");
        }

        private Table CreateTableFromXml(string tableXml)
        {
            try
            {
                Console.WriteLine($"Creating table from XML content. Length: {tableXml.Length}");
                Console.WriteLine($"XML Content: {tableXml}");

                var table = new Table();
                var doc = new XmlDocument();
                doc.LoadXml(tableXml);

                // Set up namespace manager for XPath
                var nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("w", WordMlNamespace);

                // Find table node using namespace-aware XPath
                var tableNode = doc.SelectSingleNode("//w:tbl", nsmgr);
                if (tableNode == null)
                {
                    throw new InvalidOperationException("No table found in XML content");
                }

                table.InnerXml = tableNode.InnerXml;
                var rowCount = table.Elements<TableRow>().Count();
                Console.WriteLine($"Table created successfully with {rowCount} rows");
                return table;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating table from XML: {ex.Message}");
                Console.WriteLine($"Table XML content: {tableXml}");
                throw;
            }
        }

        public string ExtractTextFromXml(string xml)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(xml))
                    return string.Empty;

                Console.WriteLine($"Processing XML content: {xml}");

                // First pass: Extract text specifically from Word text tags
                var matches = Regex.Matches(xml, @"<w:t(?:\s[^>]*)?>(.*?)</w:t>");
                if (matches.Count > 0)
                {
                    string textContent = string.Join(" ",
                        matches.Cast<Match>()
                                .Select(m => m.Groups[1].Value.Trim())
                                .Where(s => !string.IsNullOrWhiteSpace(s)));
                    Console.WriteLine($"Extracted Word text content: {textContent}");
                    return textContent;
                }

                // Fallback for non-Word XML: Remove all XML tags recursively
                string withoutTags = xml;
                string previousResult;
                do
                {
                    previousResult = withoutTags;
                    withoutTags = Regex.Replace(previousResult, @"<[^>]+>", string.Empty);
                } while (withoutTags != previousResult);

                // Clean up the result
                string decoded = System.Net.WebUtility.HtmlDecode(withoutTags);
                string normalized = Regex.Replace(decoded, @"\s+", " ").Trim();

                Console.WriteLine($"Final extracted text: {normalized}");
                return normalized;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting text from XML: {ex.Message}");
                return string.Empty;
            }
        }
    }
}