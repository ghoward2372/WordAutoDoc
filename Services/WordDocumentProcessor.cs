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
                Console.WriteLine($"\n=== Starting Document Processing ===");
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
                Console.WriteLine($"\nProcessing paragraph text: {text}");

                var processed = await ProcessTextAsync(text);

                if (processed.IsTable && processed.TableElement != null)
                {
                    Console.WriteLine("\n=== Table Processing ===");
                    try
                    {
                        var table = processed.TableElement;
                        Console.WriteLine($"Table XML before insertion: {table.OuterXml}");

                        // Add namespace to table if missing
                        if (!table.OuterXml.Contains("xmlns:w="))
                        {
                            Console.WriteLine("Adding namespace to table XML");
                            var newTable = new Table();
                            newTable.InnerXml = table.OuterXml.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                            table = newTable;
                        }

                        // Insert table and remove original paragraph
                        Console.WriteLine("Inserting table into document");
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
                else if (!processed.IsTable && text != processed.ProcessedText)
                {
                    Console.WriteLine("Updating paragraph with processed text");
                    var processedText = _options.AcronymProcessor.ProcessText(processed.ProcessedText);
                    paragraph.RemoveAllChildren();
                    paragraph.AppendChild(new Run(new Text(processedText)));
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

                        Console.WriteLine($"Processed content from {tagProcessor.Key} tag processor:");
                        Console.WriteLine(processedContent);

                        if (IsTableXml(processedContent))
                        {
                            Console.WriteLine("Table XML detected, creating Word table...");
                            result.IsTable = true;
                            result.TableElement = CreateTableFromXml(processedContent);
                            return result;
                        }

                        text = text.Replace(match.Value, processedContent);
                        Console.WriteLine($"Updated text after processing {tagProcessor.Key} tag: {text}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing {tagProcessor.Key} tag: {ex.Message}");
                        Console.WriteLine($"Stack trace: {ex.StackTrace}");
                        text = text.Replace(match.Value, $"[Error processing {tagProcessor.Key} tag: {ex.Message}]");
                    }
                }
            }

            result.ProcessedText = text;
            return result;
        }

        private bool IsTableXml(string content)
        {
            if (string.IsNullOrEmpty(content))
                return false;

            var trimmedContent = content.Trim();
            if (trimmedContent.Contains("<w:tbl"))
            {
                Console.WriteLine("Found table XML content:");
                Console.WriteLine(trimmedContent);
                return true;
            }
            return false;
        }

        private Table CreateTableFromXml(string tableXml)
        {
            try
            {
                Console.WriteLine($"Creating table from XML content:\n{tableXml}");

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

                var table = new Table();
                table.InnerXml = tableNode.InnerXml;

                // Verify table structure
                var rowCount = table.Elements<TableRow>().Count();
                var columnCount = table.Elements<TableRow>().FirstOrDefault()?.Elements<TableCell>().Count() ?? 0;
                Console.WriteLine($"Table created successfully with {rowCount} rows and {columnCount} columns per row");

                return table;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating table from XML: {ex.Message}");
                Console.WriteLine($"Table XML content:\n{tableXml}");
                throw;
            }
        }
    }
}