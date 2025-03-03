using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentProcessor.Models;
using DocumentProcessor.Models.TagProcessors;
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
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        private async Task ProcessDocumentContentAsync(Body body)
        {
            var paragraphsToProcess = body.Elements<Paragraph>().ToList();
            Console.WriteLine($"Processing {paragraphsToProcess.Count} paragraphs");

            foreach (var currentParagraph in paragraphsToProcess)
            {
                var text = currentParagraph.InnerText;
                Console.WriteLine($"\nProcessing paragraph text: {text}");

                var processed = await ProcessTextAsync(text);

                if (processed.IsTable && processed.TableElement != null)
                {
                    try
                    {
                        Console.WriteLine("Inserting table into document...");
                        body.InsertBefore(processed.TableElement, currentParagraph);
                        currentParagraph.Remove();
                        Console.WriteLine($"Table inserted successfully with OuterXml:\n{processed.TableElement.OuterXml}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error inserting table: {ex.Message}");
                        Console.WriteLine($"Stack trace: {ex.StackTrace}");
                        throw;
                    }
                }
                else if (text != processed.ProcessedText)
                {
                    if (processed.ProcessedText.Contains(TABLE_START_MARKER) ||
                        processed.ProcessedText.Contains(LIST_START_MARKER))
                    {
                        Console.WriteLine("Found special content markers, processing mixed content...");
                        InsertMixedContent(processed.ProcessedText, currentParagraph, body);
                    }
                    else
                    {
                        Console.WriteLine("Updating paragraph text");
                        currentParagraph.RemoveAllChildren();
                        currentParagraph.AppendChild(new Run(new Text(processed.ProcessedText)));
                    }
                }
            }
        }

        private void InsertMixedContent(string content, Paragraph currentParagraph, Body body)
        {
            try
            {
                Console.WriteLine("Inserting mixed content with tables and lists...");
                Console.WriteLine($"Content preview: {content.Substring(0, Math.Min(100, content.Length))}...");

                var parts = new List<(string Content, string Type)>();
                var currentText = new StringBuilder();
                var lines = content.Split('\n');

                foreach (var line in lines)
                {
                    var trimmedLine = line.Trim();
                    Console.WriteLine($"Processing line: {trimmedLine.Substring(0, Math.Min(50, trimmedLine.Length))}...");

                    if (trimmedLine == TABLE_START_MARKER)
                    {
                        if (currentText.Length > 0)
                        {
                            parts.Add((currentText.ToString().Trim(), "text"));
                            currentText.Clear();
                        }
                    }
                    else if (trimmedLine == TABLE_END_MARKER)
                    {
                        if (currentText.Length > 0)
                        {
                            parts.Add((currentText.ToString().Trim(), "table"));
                            currentText.Clear();
                        }
                    }
                    else if (trimmedLine == LIST_START_MARKER)
                    {
                        if (currentText.Length > 0)
                        {
                            parts.Add((currentText.ToString().Trim(), "text"));
                            currentText.Clear();
                        }
                    }
                    else if (trimmedLine == LIST_END_MARKER)
                    {
                        if (currentText.Length > 0)
                        {
                            parts.Add((currentText.ToString().Trim(), "list"));
                            currentText.Clear();
                        }
                    }
                    else
                    {
                        currentText.AppendLine(line);
                    }
                }

                if (currentText.Length > 0)
                {
                    parts.Add((currentText.ToString().Trim(), "text"));
                }

                Console.WriteLine($"Split content into {parts.Count} parts:");
                for (int i = 0; i < parts.Count; i++)
                {
                    Console.WriteLine($"Part {i + 1} - Type: {parts[i].Type}, Length: {parts[i].Content.Length}");
                }

                foreach (var part in parts)
                {
                    if (string.IsNullOrWhiteSpace(part.Content)) continue;

                    if (part.Type == "table")
                    {
                        Console.WriteLine("Processing embedded table");
                        try
                        {
                            var tableXml = part.Content;
                            Console.WriteLine($"Raw table XML:\n{tableXml}");

                            // Add namespace if missing
                            if (!tableXml.Contains("xmlns:w="))
                            {
                                tableXml = tableXml.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                                Console.WriteLine($"Added namespace. Updated XML:\n{tableXml}");
                            }

                            // Create new table with table properties
                            var table = new Table();
                            table.AppendChild(new TableProperties(
                                new TableStyle { Val = "TableGrid" },
                                new TableBorders(
                                    new TopBorder { Val = BorderValues.Single },
                                    new BottomBorder { Val = BorderValues.Single },
                                    new LeftBorder { Val = BorderValues.Single },
                                    new RightBorder { Val = BorderValues.Single },
                                    new InsideHorizontalBorder { Val = BorderValues.Single },
                                    new InsideVerticalBorder { Val = BorderValues.Single }
                                )
                            ));

                            // Parse and inject the table content
                            table.InnerXml = tableXml;

                            // Insert the table before removing the current paragraph
                            body.InsertBefore(table, currentParagraph);
                            Console.WriteLine($"Table inserted successfully. Table XML:\n{table.OuterXml}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error creating table: {ex.Message}");
                            Console.WriteLine($"Table XML: {part.Content}");
                            throw;
                        }
                    }
                    else if (part.Type == "list")
                    {
                        Console.WriteLine("Processing embedded list");
                        var listLines = part.Content.Split('\n');
                        foreach (var listLine in listLines)
                        {
                            if (!string.IsNullOrWhiteSpace(listLine))
                            {
                                try
                                {
                                    var newParagraph = new Paragraph();
                                    newParagraph.ParagraphProperties = new ParagraphProperties(
                                        new NumberingProperties(
                                            new NumberingLevelReference { Val = 0 },
                                            new NumberingId { Val = 1 }
                                        )
                                    );
                                    newParagraph.AppendChild(new Run(new Text(listLine.Trim())));
                                    body.InsertBefore(newParagraph, currentParagraph);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error creating list paragraph: {ex.Message}");
                                    Console.WriteLine($"Line content: {listLine}");
                                    throw;
                                }
                            }
                        }
                        Console.WriteLine("List processing complete");
                    }
                    else
                    {
                        Console.WriteLine("Processing text content");
                        var text = part.Content.Trim();
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            var newParagraph = new Paragraph(new Run(new Text(text)));
                            body.InsertBefore(newParagraph, currentParagraph);
                            Console.WriteLine("Text paragraph inserted");
                        }
                    }
                }

                currentParagraph.Remove();
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
                var matches = Regex.Matches(text, @"\[\[" + tagProcessor.Key + @":([^\]]+)\]\]");

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