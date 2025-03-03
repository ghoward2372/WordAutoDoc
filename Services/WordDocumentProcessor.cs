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

            _tagProcessors.Add("AcronymTable", new AcronymTableTagProcessor(options.AcronymProcessor, options.HtmlConverter));

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
                        Console.WriteLine($"Stack trace: {ex.StackTrace}");
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

                    Console.WriteLine($"Processing part type: {part.Type}");
                    Console.WriteLine($"Content preview: {part.Content.Substring(0, Math.Min(100, part.Content.Length))}...");

                    if (part.Type == "table")
                    {
                        Console.WriteLine("Processing embedded table");
                        var tableXml = part.Content;

                        try
                        {
                            var table = new Table();

                            // Ensure proper namespace
                            if (!tableXml.Contains("xmlns:w="))
                            {
                                tableXml = tableXml.Replace("<w:tbl>", $"<w:tbl xmlns:w=\"{WordMlNamespace}\">");
                            }

                            // Create proper table structure
                            var xElement = XElement.Parse(tableXml);
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

                            // Add rows and cells
                            foreach (var rowElement in xElement.Elements())
                            {
                                if (rowElement.Name.LocalName != "tr") continue;

                                var row = new TableRow();
                                foreach (var cellElement in rowElement.Elements())
                                {
                                    if (cellElement.Name.LocalName != "tc") continue;

                                    var cell = new TableCell();
                                    foreach (var paraElement in cellElement.Elements())
                                    {
                                        if (paraElement.Name.LocalName != "p") continue;

                                        var para = new Paragraph();
                                        foreach (var runElement in paraElement.Elements())
                                        {
                                            if (runElement.Name.LocalName != "r") continue;

                                            var run = new Run();
                                            var textElements = runElement.Elements().Where(e => e.Name.LocalName == "t");
                                            foreach (var textElement in textElements)
                                            {
                                                run.AppendChild(new Text(textElement.Value));
                                            }
                                            if (run.HasChildren)
                                            {
                                                para.AppendChild(run);
                                            }
                                        }
                                        if (para.HasChildren)
                                        {
                                            cell.AppendChild(para);
                                        }
                                    }
                                    if (cell.HasChildren)
                                    {
                                        row.AppendChild(cell);
                                    }
                                }
                                if (row.HasChildren)
                                {
                                    table.AppendChild(row);
                                }
                            }

                            if (table.HasChildren)
                            {
                                currentElement.InsertAfterSelf(table);
                                currentElement = table;
                                Console.WriteLine("Table inserted successfully");
                            }
                            else
                            {
                                Console.WriteLine("Warning: Table had no valid content to insert");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error creating table: {ex.Message}");
                            Console.WriteLine($"Table XML: {tableXml}");
                            throw;
                        }
                    }
                    else if (part.Type == "list")
                    {
                        Console.WriteLine("Processing embedded list");
                        foreach (var listParagraphXml in part.Content.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                        {
                            if (!string.IsNullOrWhiteSpace(listParagraphXml))
                            {
                                try
                                {
                                    var paragraphXml = listParagraphXml.Trim();
                                    if (!paragraphXml.Contains("xmlns:w="))
                                    {
                                        paragraphXml = paragraphXml.Replace("<w:p>", $"<w:p xmlns:w=\"{WordMlNamespace}\">");
                                    }

                                    var xElement = XElement.Parse(paragraphXml);
                                    var paragraph = new Paragraph();

                                    // Extract and set paragraph properties
                                    var pPrElement = xElement.Elements().FirstOrDefault(e => e.Name.LocalName == "pPr");
                                    if (pPrElement != null)
                                    {
                                        var numPrElement = pPrElement.Elements().FirstOrDefault(e => e.Name.LocalName == "numPr");
                                        if (numPrElement != null)
                                        {
                                            var ilvlElement = numPrElement.Elements().FirstOrDefault(e => e.Name.LocalName == "ilvl");
                                            var numIdElement = numPrElement.Elements().FirstOrDefault(e => e.Name.LocalName == "numId");

                                            if (ilvlElement?.Attribute("val")?.Value is string ilvl &&
                                                numIdElement?.Attribute("val")?.Value is string numId)
                                            {
                                                paragraph.ParagraphProperties = new ParagraphProperties(
                                                    new NumberingProperties(
                                                        new NumberingLevelReference { Val = int.Parse(ilvl) },
                                                        new NumberingId { Val = int.Parse(numId) }
                                                    )
                                                );
                                            }
                                        }
                                    }

                                    // Extract text content
                                    var runs = xElement.Descendants().Where(e => e.Name.LocalName == "t");
                                    foreach (var run in runs)
                                    {
                                        paragraph.AppendChild(new Run(new Text(run.Value)));
                                    }

                                    if (paragraph.HasChildren)
                                    {
                                        currentElement.InsertAfterSelf(paragraph);
                                        currentElement = paragraph;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error creating list paragraph: {ex.Message}");
                                    Console.WriteLine($"Paragraph XML: {listParagraphXml}");
                                    throw;
                                }
                            }
                        }
                        Console.WriteLine("List inserted successfully");
                    }
                    else
                    {
                        Console.WriteLine("Processing text content");
                        var text = part.Content.Trim();
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            var newParagraph = new Paragraph(new Run(new Text(text)));
                            currentElement.InsertAfterSelf(newParagraph);
                            currentElement = newParagraph;
                            Console.WriteLine("Text paragraph inserted");
                        }
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